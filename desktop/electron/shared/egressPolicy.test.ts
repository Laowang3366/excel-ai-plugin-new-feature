import { describe, expect, it } from "vitest";

import {
  assertRemoteDataProcessingAllowed,
  findHighConfidenceSensitiveData,
  RemoteDataPolicyError,
} from "./egressPolicy";

describe("egressPolicy", () => {
  it("blocks high-confidence credentials before remote processing", () => {
    expect(() =>
      assertRemoteDataProcessingAllowed({
        operation: "embedding",
        texts: ["token sk-1234567890abcdefghijklmnop"],
      }),
    ).toThrowError(
      expect.objectContaining({
        code: "sensitive_data_detected",
        sensitiveKinds: ["openai-style-key"],
      }),
    );
  });

  it("does not flag ordinary documentation queries", () => {
    expect(
      findHighConfidenceSensitiveData([
        "How do I configure an OpenAI API key without pasting its value?",
      ]),
    ).toEqual([]);
  });

  it("uses a typed policy error", () => {
    try {
      assertRemoteDataProcessingAllowed({
        operation: "ocr",
        texts: ["sk-1234567890abcdefghijklmnop"],
      });
    } catch (error) {
      expect(error).toBeInstanceOf(RemoteDataPolicyError);
    }
  });
});
