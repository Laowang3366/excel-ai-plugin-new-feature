import type { SqliteDatabase } from "../storage/nodeSqlite";
import {
  fieldAad,
  protectFieldValue,
  protectRequiredField,
} from "../../main-modules/localDataProtection/fieldCrypto";
import { mapGoal } from "./stateRuntimeMappers";
import type { RuntimeGoalRecord } from "./stateRuntimeTypes";

export function upsertGoalInDb(goalsDb: SqliteDatabase, goal: RuntimeGoalRecord): void {
  goalsDb
    .prepare(
      `INSERT INTO goals (
      goal_id, thread_id, objective, status, token_budget, token_usage,
      created_at, updated_at, completed_at, payload_json
    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    ON CONFLICT(goal_id) DO UPDATE SET
      thread_id = excluded.thread_id,
      objective = excluded.objective,
      status = excluded.status,
      token_budget = excluded.token_budget,
      token_usage = excluded.token_usage,
      updated_at = excluded.updated_at,
      completed_at = excluded.completed_at,
      payload_json = excluded.payload_json`,
    )
    .run(
      goal.goalId,
      goal.threadId ?? null,
      protectRequiredField(goal.objective, fieldAad("goals", "goals", goal.goalId, "objective")),
      goal.status,
      goal.tokenBudget ?? null,
      goal.tokenUsage ?? null,
      goal.createdAt,
      goal.updatedAt,
      goal.completedAt ?? null,
      goal.payload
        ? protectRequiredField(
            JSON.stringify(goal.payload),
            fieldAad("goals", "goals", goal.goalId, "payload_json"),
          )
        : null,
    );
}

export function getGoalFromDb(goalsDb: SqliteDatabase, goalId: string): RuntimeGoalRecord | null {
  const row = goalsDb.prepare(`SELECT * FROM goals WHERE goal_id = ?`).get(goalId) as
    Record<string, any> | undefined;

  return row ? mapGoal(row) : null;
}
