const fs = require("node:fs");
const path = require("node:path");

const MAX_ENTRY_BYTES = 480 * 1024;
const MAX_CHUNK_BYTES = 500 * 1024;
const MIN_ASYNC_CHUNKS = 2;

function readEntryAsset(indexHtml) {
  const scriptTag = indexHtml.match(/<script\b[^>]*\btype=["']module["'][^>]*>/iu)?.[0];
  const source = scriptTag?.match(/\bsrc=["']([^"']+\.js)["']/iu)?.[1];
  if (!source) throw new Error("Renderer index.html 缺少 module 入口脚本");
  return source.replace(/^\.\//u, "");
}

function readModulePreloads(indexHtml) {
  return new Set(
    [...indexHtml.matchAll(/<link\b[^>]*\brel=["']modulepreload["'][^>]*>/giu)]
      .map((match) => match[0].match(/\bhref=["']([^"']+\.js)["']/iu)?.[1])
      .filter(Boolean)
      .map((assetPath) => path.basename(assetPath)),
  );
}

function inspectRendererBundle(distDir) {
  const resolvedDistDir = path.resolve(distDir);
  const indexHtml = fs.readFileSync(path.join(resolvedDistDir, "index.html"), "utf8");
  const entryAsset = readEntryAsset(indexHtml);
  const modulePreloads = readModulePreloads(indexHtml);
  const assetsDir = path.join(resolvedDistDir, "assets");
  const chunks = fs
    .readdirSync(assetsDir)
    .filter((fileName) => fileName.endsWith(".js"))
    .map((fileName) => ({
      fileName,
      bytes: fs.statSync(path.join(assetsDir, fileName)).size,
    }));
  const entryFileName = path.basename(entryAsset);
  const entry = chunks.find((chunk) => chunk.fileName === entryFileName);
  if (!entry) throw new Error(`Renderer 入口文件不存在: ${entryAsset}`);
  if (entry.bytes > MAX_ENTRY_BYTES) {
    throw new Error(`Renderer 首屏入口超出预算: ${entry.bytes} > ${MAX_ENTRY_BYTES} bytes`);
  }
  const oversizedChunk = chunks.find((chunk) => chunk.bytes > MAX_CHUNK_BYTES);
  if (oversizedChunk) {
    throw new Error(
      `Renderer chunk 超出预算: ${oversizedChunk.fileName} (${oversizedChunk.bytes} bytes)`,
    );
  }
  const asyncChunkCount = chunks.filter(
    (chunk) => chunk.fileName !== entryFileName && !modulePreloads.has(chunk.fileName),
  ).length;
  if (asyncChunkCount < MIN_ASYNC_CHUNKS) {
    throw new Error(`Renderer 异步 chunk 数量不足: ${asyncChunkCount} < ${MIN_ASYNC_CHUNKS}`);
  }
  return { entry, asyncChunkCount, chunks };
}

if (require.main === module) {
  const report = inspectRendererBundle(path.resolve(__dirname, "../dist"));
  console.log(
    `Renderer bundle budget passed: entry=${(report.entry.bytes / 1024).toFixed(2)} KiB, asyncChunks=${report.asyncChunkCount}`,
  );
}

module.exports = {
  MAX_ENTRY_BYTES,
  MAX_CHUNK_BYTES,
  MIN_ASYNC_CHUNKS,
  inspectRendererBundle,
};
