const DEFAULT_LIMIT_BYTES = 1024 * 1024;
const GBK_COMPATIBLE_CHARSETS = new Set(["gbk", "gb2312", "gb18030"]);
const UTF8_CHARSETS = new Set(["utf-8", "utf8"]);

export function jsonBodyParser(options = {}) {
  const limitBytes = options.limitBytes || DEFAULT_LIMIT_BYTES;

  return (req, res, next) => {
    if (!isJsonRequest(req)) {
      next();
      return;
    }

    const chunks = [];
    let totalLength = 0;
    let tooLarge = false;

    req.on("data", (chunk) => {
      totalLength += chunk.length;
      if (totalLength > limitBytes) {
        tooLarge = true;
        req.destroy();
        return;
      }
      chunks.push(chunk);
    });

    req.on("end", () => {
      if (tooLarge) {
        res.status(413).json({ error: "请求体过大" });
        return;
      }

      try {
        req.body = parseJsonBody(Buffer.concat(chunks), req.headers["content-type"]);
        next();
      } catch {
        res.status(400).json({ error: "JSON 请求体格式错误" });
      }
    });

    req.on("error", (err) => {
      if (tooLarge) {
        res.status(413).json({ error: "请求体过大" });
        return;
      }
      next(err);
    });
  };
}

export function parseJsonBody(buffer, contentType = "") {
  if (!buffer || buffer.length === 0) return {};

  const charset = getCharset(contentType);
  const encodings = getDecodingOrder(charset);
  let lastError;

  for (const encoding of encodings) {
    try {
      const text = decodeBuffer(buffer, encoding);
      return JSON.parse(text);
    } catch (err) {
      lastError = err;
    }
  }

  throw lastError;
}

function isJsonRequest(req) {
  const method = String(req.method || "").toUpperCase();
  if (method === "GET" || method === "HEAD") return false;
  const contentType = String(req.headers["content-type"] || "");
  return /^application\/(?:[\w.+-]+\+)?json\b/i.test(contentType.trim());
}

function getCharset(contentType) {
  const match = /;\s*charset=([^;]+)/i.exec(String(contentType || ""));
  return match ? match[1].trim().toLowerCase().replace(/^"|"$/g, "") : "";
}

function getDecodingOrder(charset) {
  if (GBK_COMPATIBLE_CHARSETS.has(charset)) return ["gb18030"];
  if (UTF8_CHARSETS.has(charset) || !charset) return ["utf-8", "gb18030"];
  return [charset, "utf-8", "gb18030"];
}

function decodeBuffer(buffer, encoding) {
  const label = GBK_COMPATIBLE_CHARSETS.has(encoding) ? "gb18030" : encoding;
  const decoder = new TextDecoder(label, { fatal: true });
  return decoder.decode(buffer);
}
