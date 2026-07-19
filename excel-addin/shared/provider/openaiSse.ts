/** Browser-safe SSE parser over Uint8Array chunks (CRLF or LF, multi-byte safe). */

export type SseParseResult =
  | { kind: "data"; data: string }
  | { kind: "done" };

export class SseByteParser {
  private readonly decoder = new TextDecoder();
  private buffer = "";

  push(chunk: Uint8Array): SseParseResult[] {
    this.buffer += this.decoder.decode(chunk, { stream: true });
    return this.drain(false);
  }

  /** Finish decoder stream and parse any trailing data event without a blank line. */
  flush(): SseParseResult[] {
    this.buffer += this.decoder.decode();
    return this.drain(true);
  }

  private drain(atEof: boolean): SseParseResult[] {
    // Keep a trailing bare CR until the next chunk can complete CRLF.
    let holdCr = false;
    if (!atEof && this.buffer.endsWith("\r")) {
      this.buffer = this.buffer.slice(0, -1);
      holdCr = true;
    }
    this.buffer = this.buffer.replace(/\r\n/g, "\n").replace(/\r/g, "\n");
    if (holdCr) this.buffer += "\r";

    const out: SseParseResult[] = [];
    while (true) {
      const sep = this.buffer.indexOf("\n\n");
      if (sep < 0) break;
      const rawEvent = this.buffer.slice(0, sep);
      this.buffer = this.buffer.slice(sep + 2);
      const parsed = this.parseEvent(rawEvent);
      if (parsed) out.push(parsed);
    }

    if (atEof && this.buffer.length > 0) {
      // Accept a final event that lacked the trailing blank line.
      let trailing = this.buffer;
      if (trailing.endsWith("\r")) trailing = trailing.slice(0, -1);
      trailing = trailing.replace(/\r\n/g, "\n").replace(/\r/g, "\n");
      // Trim one trailing newline if present (event body without blank separator).
      if (trailing.endsWith("\n")) trailing = trailing.slice(0, -1);
      const parsed = this.parseEvent(trailing);
      if (parsed) out.push(parsed);
      this.buffer = "";
    }
    return out;
  }

  private parseEvent(rawEvent: string): SseParseResult | null {
    if (!rawEvent) return null;
    const dataLines: string[] = [];
    for (const line of rawEvent.split("\n")) {
      if (line.startsWith("data:")) {
        dataLines.push(line.slice(5).replace(/^ /, ""));
      }
    }
    if (dataLines.length === 0) return null;
    const data = dataLines.join("\n");
    if (data === "[DONE]") return { kind: "done" };
    return { kind: "data", data };
  }
}
