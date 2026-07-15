import { findHighConfidenceSensitiveData } from "./sensitiveData";

export type RemoteDataOperation = "web-search" | "ocr" | "invoice-extraction" | "embedding";

export type RemoteDataPolicyErrorCode =
  "remote_data_processing_disabled" | "sensitive_data_detected";

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

export { findHighConfidenceSensitiveData } from "./sensitiveData";

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
