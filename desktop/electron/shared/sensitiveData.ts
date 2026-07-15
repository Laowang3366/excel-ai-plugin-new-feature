import { createHash } from "crypto";

const REDACTION = "[REDACTED]";
const MAX_REDACTION_DEPTH = 12;
const MAX_REDACTION_ARRAY_ITEMS = 100;
const MAX_REDACTION_OBJECT_FIELDS = 100;

const HIGH_CONFIDENCE_SECRET_PATTERNS: Array<{
  kind: string;
  source: string;
  flags: string;
}> = [
  {
    kind: "private-key",
    source:
      "-----BEGIN ((?:RSA |EC |OPENSSH |DSA |ENCRYPTED )?PRIVATE KEY)-----[\\s\\S]*?-----END \\1-----",
    flags: "i",
  },
  { kind: "aws-access-key", source: "\\b(?:AKIA|ASIA)[A-Z0-9]{16}\\b", flags: "" },
  { kind: "github-token", source: "\\bgh(?:p|o|u|s|r)_[A-Za-z0-9]{30,}\\b", flags: "" },
  { kind: "slack-token", source: "\\bxox(?:b|p|a|r|s)-[A-Za-z0-9-]{20,}\\b", flags: "" },
  { kind: "google-api-key", source: "\\bAIza[A-Za-z0-9_-]{30,}\\b", flags: "" },
  { kind: "openai-style-key", source: "\\bsk-[A-Za-z0-9_-]{20,}\\b", flags: "" },
  {
    kind: "jwt",
    source: "\\beyJ[A-Za-z0-9_-]{8,}\\.[A-Za-z0-9_-]{8,}\\.[A-Za-z0-9_-]{8,}\\b",
    flags: "",
  },
];

const SENSITIVE_FIELD_SUFFIXES = [
  "apikey",
  "token",
  "password",
  "secret",
  "privatekey",
  "authorization",
  "cookie",
  "credentials",
];

const SENSITIVE_FIELD_NAMES = new Set([
  "customheaders",
  "headers",
  "remotecompactapikey",
  "mineruapitoken",
]);

export function findHighConfidenceSensitiveData(texts: string[]): string[] {
  const combined = texts.filter(Boolean).join("\n");
  if (!combined) return [];

  return HIGH_CONFIDENCE_SECRET_PATTERNS.filter(({ source, flags }) =>
    new RegExp(source, flags).test(combined),
  ).map(({ kind }) => kind);
}

export function redactSensitiveText(value: string, maxLength = 16_384): string {
  let redacted = value;
  for (const { kind, source, flags } of HIGH_CONFIDENCE_SECRET_PATTERNS) {
    redacted = redacted.replace(new RegExp(source, `${flags}g`), `[REDACTED:${kind}]`);
  }
  if (redacted.length <= maxLength) return redacted;
  return `${redacted.slice(0, Math.max(0, maxLength - 14))}...[TRUNCATED]`;
}

export function isSensitiveFieldName(fieldName: string): boolean {
  const normalized = fieldName.replace(/[^a-z0-9]/gi, "").toLowerCase();
  return (
    SENSITIVE_FIELD_NAMES.has(normalized) ||
    SENSITIVE_FIELD_SUFFIXES.some((suffix) => normalized.endsWith(suffix))
  );
}

export function redactSensitiveValue(value: unknown): unknown {
  return redactValue(value, new WeakSet<object>(), 0);
}

export function summarizeValueForAudit(value: unknown): string {
  const stats: AuditShapeStats = {
    arrays: 0,
    booleans: 0,
    nulls: 0,
    numbers: 0,
    objects: 0,
    strings: 0,
    stringCharacters: 0,
    undefinedValues: 0,
    otherValues: 0,
    redactedFields: 0,
  };
  const keys = new Set<string>();
  const seen = new WeakSet<object>();
  collectAuditShape(value, stats, keys, seen, 0);
  const shape = {
    type: describeType(value),
    ...(Array.isArray(value) ? { length: value.length } : {}),
    ...(isPlainObject(value) ? { fields: Object.keys(value).length } : {}),
    keys: Array.from(keys).sort().slice(0, 16),
    stats,
  };
  const fingerprint = createHash("sha256").update(JSON.stringify(shape)).digest("hex");
  return JSON.stringify({ ...shape, fingerprint });
}

function redactValue(value: unknown, seen: WeakSet<object>, depth: number): unknown {
  if (typeof value === "string") return redactSensitiveText(value);
  if (value === null || typeof value !== "object") return value;
  if (depth >= MAX_REDACTION_DEPTH) return "[MaxDepth]";
  if (seen.has(value)) return "[Circular]";
  seen.add(value);

  if (value instanceof Error) {
    return {
      name: value.name,
      message: redactSensitiveText(value.message),
      stack: value.stack ? redactSensitiveText(value.stack) : undefined,
    };
  }

  if (Array.isArray(value)) {
    const items = value
      .slice(0, MAX_REDACTION_ARRAY_ITEMS)
      .map((item) => redactValue(item, seen, depth + 1));
    if (value.length > MAX_REDACTION_ARRAY_ITEMS) {
      items.push(`[${value.length - MAX_REDACTION_ARRAY_ITEMS} more items]`);
    }
    return items;
  }

  const entries = Object.entries(value).slice(0, MAX_REDACTION_OBJECT_FIELDS);
  const output: Record<string, unknown> = {};
  for (const [key, entryValue] of entries) {
    output[redactSensitiveText(key, 256)] = isSensitiveFieldName(key)
      ? REDACTION
      : redactValue(entryValue, seen, depth + 1);
  }
  if (Object.keys(value).length > MAX_REDACTION_OBJECT_FIELDS) {
    output.__truncatedFields = Object.keys(value).length - MAX_REDACTION_OBJECT_FIELDS;
  }
  return output;
}

interface AuditShapeStats {
  arrays: number;
  booleans: number;
  nulls: number;
  numbers: number;
  objects: number;
  strings: number;
  stringCharacters: number;
  undefinedValues: number;
  otherValues: number;
  redactedFields: number;
}

function collectAuditShape(
  value: unknown,
  stats: AuditShapeStats,
  keys: Set<string>,
  seen: WeakSet<object>,
  depth: number,
): void {
  if (depth > MAX_REDACTION_DEPTH) return;
  if (value === null) {
    stats.nulls += 1;
    return;
  }
  if (typeof value === "string") {
    stats.strings += 1;
    stats.stringCharacters += value.length;
    return;
  }
  if (typeof value === "number") {
    stats.numbers += 1;
    return;
  }
  if (typeof value === "boolean") {
    stats.booleans += 1;
    return;
  }
  if (typeof value === "undefined") {
    stats.undefinedValues += 1;
    return;
  }
  if (typeof value !== "object") {
    stats.otherValues += 1;
    return;
  }
  if (seen.has(value)) return;
  seen.add(value);

  if (Array.isArray(value)) {
    stats.arrays += 1;
    for (const item of value.slice(0, MAX_REDACTION_ARRAY_ITEMS)) {
      collectAuditShape(item, stats, keys, seen, depth + 1);
    }
    return;
  }

  stats.objects += 1;
  for (const [key, entryValue] of Object.entries(value).slice(0, MAX_REDACTION_OBJECT_FIELDS)) {
    keys.add(redactSensitiveText(key, 64));
    if (isSensitiveFieldName(key)) {
      stats.redactedFields += 1;
      continue;
    }
    collectAuditShape(entryValue, stats, keys, seen, depth + 1);
  }
}

function describeType(value: unknown): string {
  if (value === null) return "null";
  if (Array.isArray(value)) return "array";
  return typeof value;
}

function isPlainObject(value: unknown): value is Record<string, unknown> {
  return value !== null && typeof value === "object" && !Array.isArray(value);
}
