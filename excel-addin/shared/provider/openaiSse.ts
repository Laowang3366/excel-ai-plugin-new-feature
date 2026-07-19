/** Browser-safe SSE line parser over Uint8Array chunks (CRLF or LF). */

export type SseParseResult =
  | { kind: "data"; data: string }
  | { kind: "done" };

export class SseByteParser {
  private buffer = "";

  push(chunk: Uint8Array): SseParseResult[] {
    this.buffer += new TextDecoder().decode(chunk, { stream: true });
    // Normalize CRLF to LF for event splitting
    this.buffer = this.buffer.replace(/\r\n/g, "\n").replace(/\r/g, "\n");
    const out: SseParseResult[] = [];
    while (true) {
      const sep = this.buffer.indexOf("\n\n");
      if (sep < 0) break;
      const rawEvent = this.buffer.slice(0, sep);
      this.buffer = this.buffer.slice(sep + 2);
      const dataLines: string[] = [];
      for (const line of rawEvent.split("\n")) {
        if (line.startsWith("data:")) {
          dataLines.push(line.slice(5).replace(/^ /, ""));
        }
      }
      if (dataLines.length === 0) continue;
      const data = dataLines.join("\n");
      if (data === "[DONE]") out.push({ kind: "done" });
      else out.push({ kind: "data", data });
    }
    return out;
  }

  flush(): SseParseResult[] {
    // Incomplete trailing event without blank line is ignored (incomplete).
    this.buffer = "";
    return [];
  }
}
