import { EventEmitter } from "events";
import os from "os";
import type { ChildProcess } from "child_process";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { evaluateCommand, killProcessTree, runShellSpawn, type CommandEvaluation } from "../../security/sandbox";
import { SHELL_TOOL_DEFINITIONS } from "../registry/shell";
import { DEFAULT_SHELL_TIMEOUT_MS } from "./shellExecutionLimits";
import { addShellExecutors, executeShellCommand } from "./shellExecutor";

// @MOCK_INTERFACE: security/sandbox process primitives are mocked so shellExecutor tests can assert policy handling without spawning real processes.
vi.mock("../../security/sandbox", async (importOriginal) => {
  const actual = await importOriginal<typeof import("../../security/sandbox")>();
  return {
    ...actual,
    evaluateCommand: vi.fn(),
    runShellSpawn: vi.fn(),
    killProcessTree: vi.fn(),
  };
});

const evaluateCommandMock = vi.mocked(evaluateCommand);
const runShellSpawnMock = vi.mocked(runShellSpawn);
const killProcessTreeMock = vi.mocked(killProcessTree);

const parsedCommand = { raw: "cmd", tokens: ["cmd"] };

function mockSuccessfulSpawn(): void {
  runShellSpawnMock.mockImplementation((_command, _workdir, _timeoutMs, done) => {
    const child = new EventEmitter();
    setTimeout(() => done({ stdout: "ok", stderr: "", exitCode: 0 }), 0);
    return child as unknown as ChildProcess;
  });
  killProcessTreeMock.mockResolvedValue(undefined);
}

function allowEvaluation(workdir = os.tmpdir()): CommandEvaluation {
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
      effectiveWorkdir: workdir,
      redirected: false,
    },
    parsed: [],
  };
}

function promptEvaluation(workdir = os.tmpdir()): CommandEvaluation {
  return {
    ...allowEvaluation(workdir),
    decision: "prompt",
    evaluation: {
      decision: "prompt",
      hits: [{
        matchedPrefix: ["git", "status"],
        command: parsedCommand,
        rule: {
          first: "git",
          rest: [{ kind: "single", value: "status" }],
          decision: "prompt",
          justification: "needs review",
        },
      }],
      violations: [],
      unparseable: [],
    },
    cwd: {
      allowed: true,
      effectiveWorkdir: workdir,
      redirected: true,
    },
  };
}

function forbiddenEvaluation(workdir = os.tmpdir()): CommandEvaluation {
  return {
    ...allowEvaluation(workdir),
    decision: "forbidden",
    violationMessage: "blocked by policy",
    evaluation: {
      decision: "forbidden",
      hits: [],
      violations: [{
        matchedPrefix: ["Remove-Item"],
        command: parsedCommand,
        rule: {
          first: "Remove-Item",
          rest: [],
          decision: "forbidden",
          justification: "destructive",
        },
      }],
      unparseable: [],
    },
  };
}

describe("shellExecutor", () => {
  beforeEach(() => {
    mockSuccessfulSpawn();
  });

  afterEach(() => {
    vi.clearAllMocks();
    vi.useRealTimers();
  });

  it("reuses sandbox evaluation passed by the tool executor context", async () => {
    evaluateCommandMock.mockRejectedValue(new Error("should not evaluate twice"));
    const executors = new Map();
    addShellExecutors(executors);

    const result = await executors.get("shell.execute")!.execute(
      { command: "echo ok", workdir: os.tmpdir() },
      { sandboxEvaluation: allowEvaluation() }
    );

    expect(result.success).toBe(true);
    expect(evaluateCommandMock).not.toHaveBeenCalled();
    expect(runShellSpawnMock).toHaveBeenCalledWith(
      "echo ok",
      os.tmpdir(),
      DEFAULT_SHELL_TIMEOUT_MS,
      expect.any(Function)
    );
  });

  it("keeps the registry timeout description aligned with the executor default", () => {
    const parameters = SHELL_TOOL_DEFINITIONS[0].parameters as {
      properties: { timeout_ms: { description: string } };
    };
    const timeoutDescription = parameters.properties.timeout_ms.description;

    expect(timeoutDescription).toContain(String(DEFAULT_SHELL_TIMEOUT_MS));
  });

  it("returns policy details without spawning when the command is forbidden", async () => {
    const executors = new Map();
    addShellExecutors(executors);

    const result = await executors.get("shell.execute")!.execute(
      { command: "Remove-Item -Recurse .", workdir: os.tmpdir() },
      { sandboxEvaluation: forbiddenEvaluation() }
    );

    expect(result).toMatchObject({
      success: false,
      error: "blocked by policy",
      data: {
        decision: "forbidden",
        workdirRedirected: false,
        violations: [{ matched: ["Remove-Item"], justification: "destructive" }],
      },
    });
    expect(runShellSpawnMock).not.toHaveBeenCalled();
  });

  it("executes prompt-approved commands in the sandbox effective workdir", async () => {
    const executors = new Map();
    addShellExecutors(executors);
    const effectiveWorkdir = os.tmpdir();

    const result = await executors.get("shell.execute")!.execute(
      { command: "git status", workdir: "C:\\blocked", timeout_ms: 1234 },
      { sandboxEvaluation: promptEvaluation(effectiveWorkdir) }
    );

    expect(result.success).toBe(true);
    expect(result.data).toMatchObject({
      decision: "prompt",
      workdirRequested: "C:\\blocked",
      workdirEffective: effectiveWorkdir,
      workdirRedirected: true,
      matchedRules: [{ matched: ["git", "status"], decision: "prompt", justification: "needs review" }],
    });
    expect(runShellSpawnMock).toHaveBeenCalledWith("git status", effectiveWorkdir, 1234, expect.any(Function));
  });

  it("rejects commands when the effective workdir does not exist", async () => {
    const executors = new Map();
    addShellExecutors(executors);
    const missingWorkdir = `${os.tmpdir()}\\missing-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;

    const result = await executors.get("shell.execute")!.execute(
      { command: "echo ok", workdir: missingWorkdir },
      { sandboxEvaluation: allowEvaluation(missingWorkdir) }
    );

    expect(result).toMatchObject({
      success: false,
      error: `工作目录不存在: ${missingWorkdir}`,
    });
    expect(runShellSpawnMock).not.toHaveBeenCalled();
  });

  it("kills the process tree when command execution exceeds the watchdog timeout", async () => {
    vi.useFakeTimers();
    const child = new EventEmitter();
    runShellSpawnMock.mockImplementationOnce(() => child as unknown as ChildProcess);

    const resultPromise = executeShellCommand(allowEvaluation(), "slow", os.tmpdir(), 10);
    await vi.advanceTimersByTimeAsync(510);
    const result = await resultPromise;

    expect(killProcessTreeMock).toHaveBeenCalledWith(child);
    expect(result).toEqual({
      stdout: "",
      stderr: "命令执行超时（0.01s），已强杀进程树",
      exitCode: -1,
    });
  });
});
