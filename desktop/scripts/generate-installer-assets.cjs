const fs = require("fs");
const path = require("path");

const ROOT = path.resolve(__dirname, "..");
const BUILD_DIR = path.join(ROOT, "build");

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

function setPixel(canvas, x, y, color) {
  if (x < 0 || y < 0 || x >= canvas.width || y >= canvas.height) return;
  canvas.pixels[y * canvas.width + x] = { ...color };
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

console.log("Generated installer assets in", BUILD_DIR);
