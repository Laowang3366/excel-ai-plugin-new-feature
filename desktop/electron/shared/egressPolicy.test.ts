import { describe, expect, it } from "vitest";

import {
  assertRemoteDataProcessingAllowed,
  findHighConfidenceSensitiveData,
  RemoteDataPolicyError,
} from "./egressPolicy";

describe("egressPolicy", () => {
  it("fails closed when remote processing is disabled", () => {
    expect(() => assertRemoteDataProcessingAllowed({
      enabled: false,
      operation: "web-search",
      texts: ["public query"],
    })).toThrowError(expect.objectContaining({
      code: "remote_data_processing_disabled",
    }));
  });

  it("blocks high-confidence credentials before remote processing", () => {
    expect(() => assertRemoteDataProcessingAllowed({
      enabled: true,
      operation: "embedding",
      texts: ["token sk-1234567890abcdefghijklmnop"],
    })).toThrowError(expect.objectContaining({
      code: "sensitive_data_detected",
      sensitiveKinds: ["openai-style-key"],
    }));
  });

  it("does not flag ordinary documentation queries", () => {
    expect(findHighConfidenceSensitiveData([
      "How do I configure an OpenAI API key without pasting its value?",
    ])).toEqual([]);
  });

  it("uses a typed policy error", () => {
    try {
      assertRemoteDataProcessingAllowed({ enabled: false, operation: "ocr" });
    } catch (error) {
      expect(error).toBeInstanceOf(RemoteDataPolicyError);
    }
  });
});
