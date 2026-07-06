import os from "os";
import { afterEach, describe, expect, it, vi } from "vitest";
import { executePlainPythonScript } from "../../automation/python";
import { evaluateCommand, type CommandEvaluation } from "../../security/sandbox";
import { addPythonExecutors } from "./pythonExecutor";

// @MOCK_INTERFACE: Python process and sandbox primitives are mocked so these tests cover executor contracts without spawning Python.
vi.mock("../../automation/python", () => ({
  executePlainPythonScript: vi.fn(),
}));

vi.mock("../../security/sandbox", () => ({
  evaluateCommand: vi.fn(),
}));

const executePlainPythonScriptMock = vi.mocked(executePlainPythonScript);
const evaluateCommandMock = vi.mocked(evaluateCommand);

function allowEvaluation(overrides: Partial<CommandEvaluation["cwd"]> = {}): CommandEvaluation {
  return {
    decision: "allow",
    evaluation: {
      decision: "allow",
      hits: [],
      violations: [],
      unparseable: [],
    },
    cwd: {
      allowed: true,
      effectiveWorkdir: "C:\\safe-workdir",
      redirected: false,
      ...overrides,
    },
    parsed: [],
  };
}

function createExecutors() {
  const executors = new Map();
  addPythonExecutors(executors);
  return executors;
}

describe("pythonExecutor", () => {
  afterEach(() => {
    vi.clearAllMocks();
  });

  it("registers canonical and legacy underscore tool names to the same executor", () => {
    const executors = createExecutors();

    expect(executors.get("python.execute")).toBeDefined();
    expect(executors.get("python_execute")).toBe(executors.get("python.execute"));
  });

  it("rejects missing code before sandbox evaluation or Python execution", async () => {
    const executors = createExecutors();

    const result = await executors.get("python.execute")!.execute({});

    expect(result.success).toBe(false);
    expect(result.error).toContain("code");
    expect(evaluateCommandMock).not.toHaveBeenCalled();
    expect(executePlainPythonScriptMock).not.toHaveBeenCalled();
  });

  it("executes code in the sandbox effective workdir and returns execution metadata", async () => {
    evaluateCommandMock.mockResolvedValueOnce(allowEvaluation({
      effectiveWorkdir: "D:\\project",
      redirected: true,
    }));
    executePlainPythonScriptMock.mockResolvedValueOnce({
      stdout: "ok",
      stderr: "",
      exitCode: 0,
      pythonPath: "python",
    });
    const executors = createExecutors();

    const result = await executors.get("python.execute")!.execute({
      code: "print('ok')",
      workdir: "C:\\blocked",
      timeout_ms: 1234,
    });

    expect(evaluateCommandMock).toHaveBeenCalledWith("python script.py", "C:\\blocked");
    expect(executePlainPythonScriptMock).toHaveBeenCalledWith("print('ok')", 1234, "D:\\project");
    expect(result).toMatchObject({
      success: true,
      data: {
        stdout: "ok",
        stderr: "",
        exitCode: 0,
        pythonPath: "python",
        decision: "allow",
        workdirRequested: "C:\\blocked",
        workdirEffective: "D:\\project",
        workdirRedirected: true,
      },
    });
    expect(result.error).toBeUndefined();
  });

  it("uses the home directory and default timeout when optional args are omitted", async () => {
    evaluateCommandMock.mockResolvedValueOnce(allowEvaluation({
      effectiveWorkdir: os.homedir(),
    }));
    executePlainPythonScriptMock.mockResolvedValueOnce({
      stdout: "",
      stderr: "",
      exitCode: 0,
      pythonPath: "python",
    });
    const executors = createExecutors();

    await executors.get("python.execute")!.execute({ code: "print(1)" });

    expect(evaluateCommandMock).toHaveBeenCalledWith("python script.py", os.homedir());
    expect(executePlainPythonScriptMock).toHaveBeenCalledWith("print(1)", 90000, os.homedir());
  });

  it("surfaces stderr when the Python script exits unsuccessfully", async () => {
    evaluateCommandMock.mockResolvedValueOnce(allowEvaluation());
    executePlainPythonScriptMock.mockResolvedValueOnce({
      stdout: "",
      stderr: "traceback",
      exitCode: 1,
      pythonPath: "python",
    });
    const executors = createExecutors();

    const result = await executors.get("python.execute")!.execute({ code: "raise Exception()" });

    expect(result).toMatchObject({
      success: false,
      error: "traceback",
      data: {
        stderr: "traceback",
        exitCode: 1,
      },
    });
  });
});
