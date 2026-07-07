const fs = require("fs");
const path = require("path");
const zlib = require("zlib");

const ROOT = path.resolve(__dirname, "..");
const BUILD_DIR = path.join(ROOT, "build");
const PUBLIC_DIR = path.join(ROOT, "public");

function rgb(r, g, b) {
  return { r, g, b };
}

function mix(a, b, t) {
  return rgb(
    Math.round(a.r + (b.r - a.r) * t),
    Math.round(a.g + (b.g - a.g) * t),
    Math.round(a.b + (b.b - a.b) * t)
  );
}

function createCanvas(width, height, fill) {
  return {
    width,
    height,
    pixels: Array.from({ length: width * height }, () => ({ ...fill })),
  };
}

function createTransparentCanvas(width, height) {
  return {
    width,
    height,
    pixels: Array.from({ length: width * height }, () => ({ r: 0, g: 0, b: 0, a: 0 })),
  };
}

function setPixel(canvas, x, y, color) {
  if (x < 0 || y < 0 || x >= canvas.width || y >= canvas.height) return;
  canvas.pixels[y * canvas.width + x] = { a: 255, ...color };
}

function fillRect(canvas, x, y, width, height, color) {
  for (let py = y; py < y + height; py++) {
    for (let px = x; px < x + width; px++) {
      setPixel(canvas, px, py, color);
    }
  }
}

function strokeRect(canvas, x, y, width, height, color) {
  fillRect(canvas, x, y, width, 1, color);
  fillRect(canvas, x, y + height - 1, width, 1, color);
  fillRect(canvas, x, y, 1, height, color);
  fillRect(canvas, x + width - 1, y, 1, height, color);
}

function verticalGradient(canvas, top, bottom) {
  for (let y = 0; y < canvas.height; y++) {
    const color = mix(top, bottom, y / Math.max(1, canvas.height - 1));
    fillRect(canvas, 0, y, canvas.width, 1, color);
  }
}

function drawGridMark(canvas, x, y, cell, gap, colorA, colorB) {
  for (let row = 0; row < 3; row++) {
    for (let col = 0; col < 3; col++) {
      const isAccent = (row === 0 && col === 0) || (row === 1 && col === 2) || (row === 2 && col === 1);
      fillRect(canvas, x + col * (cell + gap), y + row * (cell + gap), cell, cell, isAccent ? colorA : colorB);
    }
  }
}

function drawWenGeMark(canvas, x, y, scale, light, mid) {
  const s = scale;
  fillRect(canvas, x, y, 7 * s, 2 * s, light);
  fillRect(canvas, x, y, 2 * s, 14 * s, light);
  fillRect(canvas, x, y + 6 * s, 7 * s, 2 * s, light);
  fillRect(canvas, x + 5 * s, y + 6 * s, 2 * s, 8 * s, light);

  const gx = x + 10 * s;
  fillRect(canvas, gx, y, 10 * s, 2 * s, mid);
  fillRect(canvas, gx, y, 2 * s, 14 * s, mid);
  fillRect(canvas, gx + 8 * s, y, 2 * s, 14 * s, mid);
  fillRect(canvas, gx, y + 6 * s, 10 * s, 2 * s, mid);
  fillRect(canvas, gx, y + 12 * s, 10 * s, 2 * s, mid);
}

function drawSidebar(filePath) {
  const canvas = createCanvas(164, 314, rgb(18, 83, 64));
  verticalGradient(canvas, rgb(12, 73, 56), rgb(32, 138, 103));

  fillRect(canvas, 0, 0, 164, 314, rgb(15, 86, 66));
  for (let y = 0; y < canvas.height; y++) {
    const color = mix(rgb(13, 75, 58), rgb(38, 147, 109), y / (canvas.height - 1));
    fillRect(canvas, 0, y, 164, 1, color);
  }

  fillRect(canvas, 18, 24, 48, 48, rgb(238, 255, 247));
  fillRect(canvas, 24, 30, 36, 36, rgb(25, 119, 89));
  drawGridMark(canvas, 28, 34, 8, 3, rgb(224, 252, 239), rgb(149, 225, 190));

  drawWenGeMark(canvas, 20, 96, 3, rgb(244, 255, 250), rgb(198, 244, 222));
  fillRect(canvas, 20, 164, 112, 3, rgb(203, 244, 221));
  fillRect(canvas, 20, 178, 92, 3, rgb(169, 226, 197));
  fillRect(canvas, 20, 192, 72, 3, rgb(142, 211, 180));

  strokeRect(canvas, 20, 232, 86, 52, rgb(165, 228, 198));
  fillRect(canvas, 30, 242, 66, 5, rgb(232, 255, 246));
  fillRect(canvas, 30, 255, 44, 5, rgb(193, 240, 216));
  fillRect(canvas, 30, 268, 56, 5, rgb(193, 240, 216));

  drawBmp(canvas, filePath);
}

function drawHeader(filePath) {
  const canvas = createCanvas(150, 57, rgb(248, 251, 249));
  fillRect(canvas, 0, 0, 150, 57, rgb(248, 251, 249));
  fillRect(canvas, 0, 55, 150, 2, rgb(219, 228, 222));

  fillRect(canvas, 8, 9, 38, 38, rgb(31, 138, 103));
  fillRect(canvas, 13, 14, 28, 28, rgb(238, 255, 247));
  drawGridMark(canvas, 17, 18, 6, 2, rgb(25, 119, 89), rgb(132, 211, 175));

  drawWenGeMark(canvas, 57, 12, 2, rgb(22, 96, 72), rgb(31, 138, 103));
  fillRect(canvas, 57, 43, 74, 3, rgb(91, 113, 101));

  drawBmp(canvas, filePath);
}

function drawIconCanvas(size) {
  const canvas = createTransparentCanvas(size, size);
  const margin = Math.max(2, Math.round(size * 0.1));
  const inner = size - margin * 2;
  const radius = Math.max(2, Math.round(size * 0.08));
  const bgTop = rgb(18, 117, 86);
  const bgBottom = rgb(35, 154, 112);

  for (let y = margin; y < size - margin; y++) {
    const t = (y - margin) / Math.max(1, inner - 1);
    const color = mix(bgTop, bgBottom, t);
    for (let x = margin; x < size - margin; x++) {
      const left = x - margin;
      const right = size - margin - 1 - x;
      const top = y - margin;
      const bottom = size - margin - 1 - y;
      const inCorner =
        (left >= radius || top >= radius || (left - radius) ** 2 + (top - radius) ** 2 <= radius ** 2) &&
        (right >= radius || top >= radius || (right - radius) ** 2 + (top - radius) ** 2 <= radius ** 2) &&
        (left >= radius || bottom >= radius || (left - radius) ** 2 + (bottom - radius) ** 2 <= radius ** 2) &&
        (right >= radius || bottom >= radius || (right - radius) ** 2 + (bottom - radius) ** 2 <= radius ** 2);
      if (inCorner) setPixel(canvas, x, y, color);
    }
  }

  const gridMargin = Math.round(size * 0.24);
  const gridSize = size - gridMargin * 2;
  const gap = Math.max(1, Math.round(size * 0.035));
  const cell = Math.floor((gridSize - gap * 2) / 3);
  const start = Math.round((size - (cell * 3 + gap * 2)) / 2);
  drawGridMark(
    canvas,
    start,
    start,
    cell,
    gap,
    rgb(239, 255, 248),
    rgb(155, 226, 190)
  );

  const shine = rgb(255, 255, 255);
  for (let i = 0; i < Math.max(1, Math.round(size * 0.018)); i++) {
    fillRect(canvas, margin + Math.round(size * 0.12), margin + Math.round(size * 0.1) + i, Math.round(size * 0.42), 1, shine);
  }
  return canvas;
}

function encodeIcoImage(canvas) {
  const width = canvas.width;
  const height = canvas.height;
  const headerSize = 40;
  const xorStride = width * 4;
  const andStride = Math.ceil(width / 32) * 4;
  const imageSize = headerSize + xorStride * height + andStride * height;
  const buffer = Buffer.alloc(imageSize);

  buffer.writeUInt32LE(headerSize, 0);
  buffer.writeInt32LE(width, 4);
  buffer.writeInt32LE(height * 2, 8);
  buffer.writeUInt16LE(1, 12);
  buffer.writeUInt16LE(32, 14);
  buffer.writeUInt32LE(0, 16);
  buffer.writeUInt32LE(xorStride * height + andStride * height, 20);
  buffer.writeInt32LE(2835, 24);
  buffer.writeInt32LE(2835, 28);

  let offset = headerSize;
  for (let y = height - 1; y >= 0; y--) {
    for (let x = 0; x < width; x++) {
      const color = canvas.pixels[y * width + x];
      buffer[offset++] = color.b;
      buffer[offset++] = color.g;
      buffer[offset++] = color.r;
      buffer[offset++] = color.a ?? 255;
    }
  }

  for (let y = height - 1; y >= 0; y--) {
    for (let xByte = 0; xByte < andStride; xByte++) {
      let value = 0;
      for (let bit = 0; bit < 8; bit++) {
        const x = xByte * 8 + bit;
        if (x >= width) continue;
        const alpha = canvas.pixels[y * width + x].a ?? 255;
        if (alpha === 0) value |= 0x80 >> bit;
      }
      buffer[offset++] = value;
    }
  }

  return buffer;
}

function crc32(buffer) {
  let crc = 0xffffffff;
  for (const byte of buffer) {
    crc ^= byte;
    for (let bit = 0; bit < 8; bit++) {
      crc = (crc >>> 1) ^ (0xedb88320 & -(crc & 1));
    }
  }
  return (crc ^ 0xffffffff) >>> 0;
}

function pngChunk(type, data) {
  const typeBuffer = Buffer.from(type, "ascii");
  const chunk = Buffer.alloc(12 + data.length);
  chunk.writeUInt32BE(data.length, 0);
  typeBuffer.copy(chunk, 4);
  data.copy(chunk, 8);
  chunk.writeUInt32BE(crc32(Buffer.concat([typeBuffer, data])), 8 + data.length);
  return chunk;
}

function drawPng(canvas, filePath) {
  const raw = Buffer.alloc((canvas.width * 4 + 1) * canvas.height);
  let offset = 0;
  for (let y = 0; y < canvas.height; y++) {
    raw[offset++] = 0;
    for (let x = 0; x < canvas.width; x++) {
      const color = canvas.pixels[y * canvas.width + x];
      raw[offset++] = color.r;
      raw[offset++] = color.g;
      raw[offset++] = color.b;
      raw[offset++] = color.a ?? 255;
    }
  }

  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(canvas.width, 0);
  ihdr.writeUInt32BE(canvas.height, 4);
  ihdr[8] = 8;
  ihdr[9] = 6;
  ihdr[10] = 0;
  ihdr[11] = 0;
  ihdr[12] = 0;

  const png = Buffer.concat([
    Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
    pngChunk("IHDR", ihdr),
    pngChunk("IDAT", zlib.deflateSync(raw, { level: 9 })),
    pngChunk("IEND", Buffer.alloc(0)),
  ]);

  fs.mkdirSync(path.dirname(filePath), { recursive: true });
  fs.writeFileSync(filePath, png);
}

function drawIco(filePath) {
  const sizes = [16, 24, 32, 48, 64, 128, 256];
  const images = sizes.map((size) => ({ size, data: encodeIcoImage(drawIconCanvas(size)) }));
  const headerSize = 6;
  const dirSize = 16 * images.length;
  const totalSize = headerSize + dirSize + images.reduce((sum, image) => sum + image.data.length, 0);
  const buffer = Buffer.alloc(totalSize);

  buffer.writeUInt16LE(0, 0);
  buffer.writeUInt16LE(1, 2);
  buffer.writeUInt16LE(images.length, 4);

  let imageOffset = headerSize + dirSize;
  images.forEach((image, index) => {
    const dirOffset = headerSize + index * 16;
    buffer[dirOffset] = image.size === 256 ? 0 : image.size;
    buffer[dirOffset + 1] = image.size === 256 ? 0 : image.size;
    buffer[dirOffset + 2] = 0;
    buffer[dirOffset + 3] = 0;
    buffer.writeUInt16LE(1, dirOffset + 4);
    buffer.writeUInt16LE(32, dirOffset + 6);
    buffer.writeUInt32LE(image.data.length, dirOffset + 8);
    buffer.writeUInt32LE(imageOffset, dirOffset + 12);
    image.data.copy(buffer, imageOffset);
    imageOffset += image.data.length;
  });

  fs.mkdirSync(path.dirname(filePath), { recursive: true });
  fs.writeFileSync(filePath, buffer);
}

function drawBmp(canvas, filePath) {
  const rowSize = Math.ceil((canvas.width * 3) / 4) * 4;
  const pixelArraySize = rowSize * canvas.height;
  const fileSize = 54 + pixelArraySize;
  const buffer = Buffer.alloc(fileSize);

  buffer.write("BM", 0, "ascii");
  buffer.writeUInt32LE(fileSize, 2);
  buffer.writeUInt32LE(54, 10);
  buffer.writeUInt32LE(40, 14);
  buffer.writeInt32LE(canvas.width, 18);
  buffer.writeInt32LE(canvas.height, 22);
  buffer.writeUInt16LE(1, 26);
  buffer.writeUInt16LE(24, 28);
  buffer.writeUInt32LE(0, 30);
  buffer.writeUInt32LE(pixelArraySize, 34);
  buffer.writeInt32LE(2835, 38);
  buffer.writeInt32LE(2835, 42);

  for (let y = 0; y < canvas.height; y++) {
    const sourceY = canvas.height - 1 - y;
    const rowOffset = 54 + y * rowSize;
    for (let x = 0; x < canvas.width; x++) {
      const color = canvas.pixels[sourceY * canvas.width + x];
      const offset = rowOffset + x * 3;
      buffer[offset] = color.b;
      buffer[offset + 1] = color.g;
      buffer[offset + 2] = color.r;
    }
  }

  fs.mkdirSync(path.dirname(filePath), { recursive: true });
  fs.writeFileSync(filePath, buffer);
}

drawSidebar(path.join(BUILD_DIR, "installer-sidebar.bmp"));
drawHeader(path.join(BUILD_DIR, "installer-header.bmp"));
drawIco(path.join(BUILD_DIR, "icon.ico"));
drawIco(path.join(BUILD_DIR, "installerIcon.ico"));
drawIco(path.join(BUILD_DIR, "uninstallerIcon.ico"));
drawPng(drawIconCanvas(256), path.join(PUBLIC_DIR, "icon.png"));

console.log("Generated installer assets in", BUILD_DIR);
