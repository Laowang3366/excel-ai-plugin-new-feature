/**
 * 文本清理工具函数
 *
 * 主要解决国内 LLM（Kimi、DeepSeek、讯飞等）推理/思考模式下，
 * reasoning_content 的 token 级别空格问题：
 *
 *   输入: "我已 经 读取 了 Sheet 2 ! D 2 : F 8"
 *   期望: "我已经读取了Sheet2!D2:F8"
 *
 * 规则：
 * 1. 两个 CJK 字符之间的空格 → 移除
 * 2. CJK 字符与英文/数字之间的空格 → 移除
 * 3. CJK 字符与标点之间的空格 → 移除
 * 4. 英文单词之间的空格 → 保留
 * 5. Markdown 语法（表格、列表等）中的必要空格 → 保留
 */

/** 判断字符是否为 CJK 统一表意文字 */
function isCJK(char: string): boolean {
  const code = char.codePointAt(0)!;
  return (
    (code >= 0x4e00 && code <= 0x9fff) ||   // CJK 统一汉字
    (code >= 0x3400 && code <= 0x4dbf) ||   // CJK 扩展A
    (code >= 0x20000 && code <= 0x2a6df) || // CJK 扩展B-F
    (code >= 0xf900 && code <= 0xfaff) ||   // CJK 兼容汉字
    (code >= 0x2f800 && code <= 0x2fa1f) || // CJK 兼容补充
    (code >= 0x3000 && code <= 0x303f) ||   // CJK 符号和标点（。、「」等）
    (code >= 0xff00 && code <= 0xffef)      // 全角字符
  );
}

/** 判断字符是否为常见标点（中英文） */
function isPunctuation(char: string): boolean {
  const code = char.codePointAt(0)!;
  return (
    // 英文标点
    code === 0x21 || // !
    code === 0x3a || // :
    code === 0x3b || // ;
    code === 0x2c || // ,
    code === 0x2e || // .
    code === 0x3f || // ?
    code === 0x28 || // (
    code === 0x29 || // )
    code === 0x5b || // [
    code === 0x5d || // ]
    code === 0x7b || // {
    code === 0x7d || // }
    // 中文标点
    code === 0x3001 || // 、
    code === 0x3002 || // 。
    code === 0xff01 || // ！
    code === 0xff0c || // ，
    code === 0xff1a || // ：
    code === 0xff1b || // ；
    code === 0xff1f || // ？
    code === 0x201c || code === 0x201d || // ""
    code === 0x2018 || code === 0x2019     // ''
  );
}

/**
 * 清理推理文本中的 token 级别空格
 *
 * 核心逻辑：遍历文本，当遇到空格时，检查其前后字符：
 * - 如果空格两侧都是 CJK 字符 → 移除空格
 * - 如果空格一侧是 CJK，另一侧是英文/数字/标点 → 移除空格
 * - 如果空格两侧都是英文/数字 → 保留空格（正常英文分词）
 * - Markdown 表格行（含 | ）中的空格需要特殊处理
 */
export function cleanReasoningText(text: string): string {
  if (!text) return text;

  // 先处理 Markdown 表格：表格行中的空格需要保留（| col1 | col2 |）
  // 策略：将表格行单独处理，非表格行做 CJK 空格清理
  const lines = text.split("\n");
  const cleanedLines = lines.map((line) => {
    // 检测 Markdown 表格行（以 | 开头或包含多个 |）
    const pipeCount = (line.match(/\|/g) || []).length;
    if (pipeCount >= 2) {
      // 表格行：只清理单元格内部的 CJK 空格，保留 | 周围的空格
      return cleanTableCellSpaces(line);
    }
    return cleanCJKSpaces(line);
  });

  // 合并行，再压缩连续空行
  return cleanedLines.join("\n").replace(/\n{2,}/g, "\n\n");
}

/**
 * 清理一行文本中 CJK 字符间的多余空格
 *
 * 针对国内 LLM（Kimi、DeepSeek 等）reasoning_content 的 token 级别空格问题。
 * 这些模型将推理文本按 token 输出，token 之间会插入空格：
 *   "我已 经 读取 了 Sheet 2 ! D 2 : F 8"
 * 期望结果：
 *   "我已经读取了Sheet2!D2:F8"
 *
 * 核心判断：空格是否属于英文单词分词？
 * - 如果空格两侧都是较长的英文单词（≥2字母），保留空格
 * - 其他情况（CJK、数字、标点、单字母）移除空格
 */
function cleanCJKSpaces(line: string): string {
  const chars = [...line]; // 正确处理 Unicode 代理对
  const result: string[] = [];

  for (let i = 0; i < chars.length; i++) {
    const ch = chars[i];

    // 遇到空格时，检查是否需要移除
    if (ch === " " || ch === "\u3000") { // 半角空格或全角空格
      const prevChar = result.length > 0 ? result[result.length - 1] : "";
      const nextChar = i + 1 < chars.length ? chars[i + 1] : "";

      // 如果空格在行首或行尾，移除
      if (!prevChar || !nextChar) {
        continue;
      }

      // 如果前后有 CJK 字符，移除空格
      if (isCJK(prevChar) || isCJK(nextChar)) {
        continue;
      }

      // 如果前后有标点，移除空格
      if (isPunctuation(prevChar) || isPunctuation(nextChar)) {
        continue;
      }

      // 字母与数字之间（如 D 2, Sheet 2, A 1）→ 移除空格
      // 这是 token 级别分词的典型特征
      if ((isAlpha(prevChar) && isDigit(nextChar)) ||
          (isDigit(prevChar) && isAlpha(nextChar))) {
        continue;
      }

      // 数字与数字之间（如 1 000）→ 移除空格（token 级别）
      if (isDigit(prevChar) && isDigit(nextChar)) {
        continue;
      }

      // 单字母与单字母之间（如 A B C）→ 可能是 token 级别，移除
      if (isAlpha(prevChar) && isAlpha(nextChar)) {
        // 检查是否属于更长的英文单词
        // 向前扫描：prevChar 之前是否连续有字母（形成单词）
        const prevWordLen = countPrevAlphaLen(result, result.length - 1);
        // 向后扫描：nextChar 之后是否连续有字母（形成单词）
        const nextWordLen = countNextAlphaLen(chars, i + 1);

        // 如果两侧都是较长的英文单词（≥2字母），保留空格
        // 例如 "the quick" → prevWord=3, nextWord=5 → 保留
        // 例如 "D 2" → prevWord=1, nextWord=0 → 移除
        if (prevWordLen >= 2 && nextWordLen >= 2) {
          result.push(" ");
          continue;
        }
        // 否则移除（token 级别空格）
        continue;
      }

      // 其他情况保留空格
      result.push(" ");
    } else {
      result.push(ch);
    }
  }

  return result.join("");
}

/** 判断字符是否为字母 */
function isAlpha(char: string): boolean {
  const code = char.codePointAt(0)!;
  return (code >= 0x41 && code <= 0x5a) || (code >= 0x61 && code <= 0x7a);
}

/** 判断字符是否为数字 */
function isDigit(char: string): boolean {
  const code = char.codePointAt(0)!;
  return code >= 0x30 && code <= 0x39;
}

/** 从 result 数组的指定位置向前统计连续字母长度 */
function countPrevAlphaLen(result: string[], fromIdx: number): number {
  let count = 0;
  for (let i = fromIdx; i >= 0; i--) {
    if (isAlpha(result[i])) {
      count++;
    } else {
      break;
    }
  }
  return count;
}

/** 从 chars 数组的指定位置向后统计连续字母长度 */
function countNextAlphaLen(chars: string[], fromIdx: number): number {
  let count = 0;
  for (let i = fromIdx; i < chars.length; i++) {
    if (isAlpha(chars[i])) {
      count++;
    } else {
      break;
    }
  }
  return count;
}

/**
 * 清理 Markdown 表格行中的空格
 *
 * 表格行格式: | col1 | col2 | col3 |
 * 策略：按 | 分割，对每个单元格内容做 CJK 空格清理，再重新拼接
 */
function cleanTableCellSpaces(line: string): string {
  // 分割表格单元格
  const parts = line.split("|");
  const cleanedParts = parts.map((part, idx) => {
    // 首尾空元素（行首/行尾的 | 产生的）保留
    if (idx === 0 && part === "") return part;
    if (idx === parts.length - 1 && part === "") return part;

    // 分隔线行（如 |---|---|）不做处理
    if (/^[-:\s]+$/.test(part)) return part;

    // 对单元格内容做 CJK 空格清理，但保留 | 两侧的空格
    const trimmed = part;
    // 保留单元格前后各一个空格（Markdown 表格格式需要）
    const leadingSpace = trimmed.startsWith(" ") ? " " : "";
    const trailingSpace = trimmed.endsWith(" ") ? " " : "";
    const inner = trimmed.trim();

    return leadingSpace + cleanCJKSpaces(inner) + trailingSpace;
  });

  return cleanedParts.join("|");
}
