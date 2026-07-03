declare module "node:sqlite" {
  export type SupportedValueType = string | number | bigint | boolean | Uint8Array | null;

  export interface StatementResultingChanges {
    changes: number;
    lastInsertRowid: number | bigint;
  }

  export class StatementSync {
    readonly sourceSQL: string;
    readonly expandedSQL: string;
    run(...params: SupportedValueType[]): StatementResultingChanges;
    get(...params: SupportedValueType[]): Record<string, unknown> | undefined;
    all(...params: SupportedValueType[]): Array<Record<string, unknown>>;
  }

  export class DatabaseSync {
    constructor(path: string);
    exec(sql: string): void;
    prepare(sql: string): StatementSync;
    close(): void;
  }
}
