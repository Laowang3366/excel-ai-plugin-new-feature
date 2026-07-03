/**
 * 图片附件解析 — 将本地文件路径转为 base64 data URI
 *
 * 从 agentLoop.ts 提取的纯函数，无类依赖。
 */

import * as fs from "fs";
import * as path from "path";
import type { ChatMessage, ContentPart } from "../providers/aiClient";

/** MIME 类型映射 */
const IMAGE_MIME_MAP: Record<string, string> = {
  ".png": "image/png",
  ".jpg": "image/jpeg",
  ".jpeg": "image/jpeg",
  ".gif": "image/gif",
  ".webp": "image/webp",
  ".bmp": "image/bmp",
};

/**
 * 解析 ChatMessage 中的本地图片路径，替换为 base64 data URI
 *
 * messageBuilder 为图片附件生成 file:// 路径占位，
 * 此函数在发送给 AI 前将它们解析为实际数据。
 */
export async function resolveImageAttachments(messages: ChatMessage[]): Promise<void> {
  for (const msg of messages) {
    if (msg.role !== "user" || !Array.isArray(msg.content)) continue;

    for (let i = 0; i < msg.content.length; i++) {
      const part = msg.content[i];
      if (part.type !== "image_url") continue;

      const url = part.image_url.url;
      if (!url.startsWith("file://")) continue;

      const filePath = url.slice("file://".length);
      try {
        const buffer = await fs.promises.readFile(filePath);
        const ext = path.extname(filePath).toLowerCase();
        const mimeType = IMAGE_MIME_MAP[ext] || "image/png";
        const base64 = buffer.toString("base64");

        part.image_url.url = `data:${mimeType};base64,${base64}`;
      } catch (err: any) {
        (msg.content as ContentPart[])[i] = {
          type: "text",
          text: `[图片读取失败: ${path.basename(filePath)} — ${err.message}]`,
        };
      }
    }
  }
}
