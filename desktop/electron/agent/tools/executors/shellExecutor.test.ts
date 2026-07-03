import { EventEmitter } from "events";
import os from "os";
import { describe, expect, it, vi } from "vitest";
import { evaluateCommand, runShellSpawn, type CommandEvaluation } from "../../security/sandbox";
import { addShellExecutors } from "./shellExecutor";

vi.mock("../../security/sandbox", async (importOriginal) => {
  const actual = await importOriginal<typeof import("../../security/sandbox")>();
  return {
    ...actual,
    evaluateCommand: vi.fn(),
    runShellSpawn: vi.fn((_command, _workdir, _timeoutMs, done) => {
      const child = new EventEmitter();
      setTimeout(() => done({ stdout: "ok", stderr: "", exitCode: 0 }), 0);
      return child;
    }),
    killProcessTree: vi.fn(async () => {}),
  };
});

const evaluateCommandMock = vi.mocked(evaluateCommand);
const runShellSpawnMock = vi.mocked(runShellSpawn);

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

describe("shellExecutor", () => {
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
      30000,
      expect.any(Function)
    );
  });
});
