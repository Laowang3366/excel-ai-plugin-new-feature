export type RemoteDataOperation =
  | "web-search"
  | "ocr"
  | "invoice-extraction"
  | "embedding";

export type RemoteDataPolicyErrorCode =
  | "remote_data_processing_disabled"
  | "sensitive_data_detected";

export interface RemoteDataTransferSummary {
  operation: RemoteDataOperation;
  service: string;
  destination: string;
  dataSummary: string;
}

export class RemoteDataPolicyError extends Error {
  readonly code: RemoteDataPolicyErrorCode;
  readonly operation: RemoteDataOperation;
  readonly sensitiveKinds: string[];

  constructor(
    code: RemoteDataPolicyErrorCode,
    operation: RemoteDataOperation,
    message: string,
    sensitiveKinds: string[] = [],
  ) {
    super(message);
    this.name = "RemoteDataPolicyError";
    this.code = code;
    this.operation = operation;
    this.sensitiveKinds = sensitiveKinds;
  }
}

const HIGH_CONFIDENCE_SECRET_PATTERNS: Array<{ kind: string; pattern: RegExp }> = [
  { kind: "private-key", pattern: /-----BEGIN (?:RSA |EC |OPENSSH )?PRIVATE KEY-----/i },
  { kind: "aws-access-key", pattern: /\b(?:AKIA|ASIA)[A-Z0-9]{16}\b/ },
  { kind: "github-token", pattern: /\bgh(?:p|o|u|s|r)_[A-Za-z0-9]{30,}\b/ },
  { kind: "slack-token", pattern: /\bxox(?:b|p|a|r|s)-[A-Za-z0-9-]{20,}\b/ },
  { kind: "google-api-key", pattern: /\bAIza[A-Za-z0-9_-]{30,}\b/ },
  { kind: "openai-style-key", pattern: /\bsk-[A-Za-z0-9_-]{20,}\b/ },
  { kind: "jwt", pattern: /\beyJ[A-Za-z0-9_-]{8,}\.[A-Za-z0-9_-]{8,}\.[A-Za-z0-9_-]{8,}\b/ },
];

export function isRemoteDataProcessingEnabled(value: unknown): boolean {
  return value === true;
}

export function assertRemoteDataProcessingAllowed(options: {
  enabled: boolean;
  operation: RemoteDataOperation;
  texts?: string[];
}): void {
  if (!options.enabled) {
    throw new RemoteDataPolicyError(
      "remote_data_processing_disabled",
      options.operation,
      "远程数据处理已关闭；请在常规设置中明确开启后再执行此操作",
    );
  }

  const sensitiveKinds = findHighConfidenceSensitiveData(options.texts || []);
  if (sensitiveKinds.length > 0) {
    throw new RemoteDataPolicyError(
      "sensitive_data_detected",
      options.operation,
      `检测到高置信敏感凭据（${sensitiveKinds.join("、")}），已在发送前阻止`,
      sensitiveKinds,
    );
  }
}

export function findHighConfidenceSensitiveData(texts: string[]): string[] {
  const combined = texts.filter(Boolean).join("\n");
  if (!combined) return [];
  return HIGH_CONFIDENCE_SECRET_PATTERNS
    .filter(({ pattern }) => pattern.test(combined))
    .map(({ kind }) => kind);
}

export function toRemoteDataPolicyResult(error: unknown): {
  success: false;
  error: string;
  data: {
    code: RemoteDataPolicyErrorCode;
    operation: RemoteDataOperation;
    sensitiveKinds: string[];
  };
} | null {
  if (!(error instanceof RemoteDataPolicyError)) return null;
  return {
    success: false,
    error: error.message,
    data: {
      code: error.code,
      operation: error.operation,
      sensitiveKinds: error.sensitiveKinds,
    },
  };
}
