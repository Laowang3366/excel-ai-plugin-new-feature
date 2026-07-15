import * as zlib from "node:zlib";

export async function zstdCompress(content: Buffer): Promise<Buffer> {
  const compress = (
    zlib as typeof zlib & {
      zstdCompress?: (
        buffer: Buffer,
        callback: (error: Error | null, result: Buffer) => void,
      ) => void;
    }
  ).zstdCompress;
  if (!compress) throw new Error("当前 Node 运行时不支持 zstd 压缩");
  return new Promise((resolve, reject) => {
    compress(content, (error, result) => {
      if (error) reject(error);
      else resolve(result);
    });
  });
}

export async function zstdDecompress(content: Buffer): Promise<Buffer> {
  const decompress = (
    zlib as typeof zlib & {
      zstdDecompress?: (
        buffer: Buffer,
        callback: (error: Error | null, result: Buffer) => void,
      ) => void;
    }
  ).zstdDecompress;
  if (!decompress) throw new Error("当前 Node 运行时不支持 zstd 解压");
  return new Promise((resolve, reject) => {
    decompress(content, (error, result) => {
      if (error) reject(error);
      else resolve(result);
    });
  });
}
