/**
 * 自定义 JSON 请求体解析中间件
 *
 * 为什么不用 express.json()？
 * 插件端的 HTTP 请求可能使用 GBK/GB2312 编码发送 JSON 请求体，
 * 而 express.json() 仅支持 UTF-8 编码。此中间件增加了多编码支持。
 *
 * 工作流程：
 * 1. 通过 Content-Type 头检测请求是否为 JSON
 * 2. 收集请求体 chunk，限制大小（默认 1MB）
 * 3. 根据 Content-Type 中的 charset 选择合适的编码解码
 * 4. 若未指定 charset，尝试 UTF-8 -> GB18030 回退解码
 * 5. 将解析后的对象注入 req.body
 *
 * 安全考虑：
 * - 请求体大小限制（limitBytes）防止内存耗尽攻击
 * - 解码失败返回 400，不会 crash
 * - 超过限制立即销毁连接（req.destroy()）
 */

/** 默认请求体大小限制：1MB */
const DEFAULT_LIMIT_BYTES = 1024 * 1024;

/** GBK 兼容字符集集合（均使用 GB18030 解码） */
const GBK_COMPATIBLE_CHARSETS = new Set(["gbk", "gb2312", "gb18030"]);

/** UTF-8 字符集集合 */
const UTF8_CHARSETS = new Set(["utf-8", "utf8"]);

/**
 * 创建 JSON 解析中间件
 *
 * @param {object} [options]
 * @param {number} [options.limitBytes=1048576] - 请求体大小上限（字节）
 * @returns {import('express').RequestHandler}
 */
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
        req.destroy(); // 立即终止连接，不再接收更多数据
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

/**
 * 解析 JSON 请求体（支持多编码）
 *
 * @param {Buffer} buffer - 原始请求体二进制
 * @param {string} contentType - Content-Type 头
 * @returns {object} 解析后的 JSON 对象
 */
export function parseJsonBody(buffer, contentType = "") {
  if (!buffer || buffer.length === 0) return {};

  const charset = getCharset(contentType);
  const encodings = getDecodingOrder(charset);
  let lastError;

  // 按优先级依次尝试解码，直到成功或所有编码都失败
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

/**
 * 判断请求是否为 JSON 请求
 *
 * 检查方法（排除 GET/HEAD）和 Content-Type 头。
 */
function isJsonRequest(req) {
  const method = String(req.method || "").toUpperCase();
  if (method === "GET" || method === "HEAD") return false;
  const contentType = String(req.headers["content-type"] || "");
  return /^application\/(?:[\w.+-]+\+)?json\b/i.test(contentType.trim());
}

/** 从 Content-Type 头中提取 charset */
function getCharset(contentType) {
  const match = /;\s*charset=([^;]+)/i.exec(String(contentType || ""));
  return match ? match[1].trim().toLowerCase().replace(/^"|"$/g, "") : "";
}

/**
 * 确定解码尝试顺序
 *
 * - 如果指定了 GBK 系列编码，只尝试 GB18030
 * - 如果指定了 UTF-8 或未指定，先尝试 UTF-8 再回退 GB18030
 * - 其他编码，先尝试指定编码再回退 UTF-8 和 GB18030
 */
function getDecodingOrder(charset) {
  if (GBK_COMPATIBLE_CHARSETS.has(charset)) return ["gb18030"];
  if (UTF8_CHARSETS.has(charset) || !charset) return ["utf-8", "gb18030"];
  return [charset, "utf-8", "gb18030"];
}

/**
 * 使用指定编码解码 Buffer
 *
 * 使用 TextDecoder API（Node.js 16+ 内置），fatal: true 表示解码失败时抛出错误而非替换字符。
 */
function decodeBuffer(buffer, encoding) {
  const label = GBK_COMPATIBLE_CHARSETS.has(encoding) ? "gb18030" : encoding;
  const decoder = new TextDecoder(label, { fatal: true });
  return decoder.decode(buffer);
}
