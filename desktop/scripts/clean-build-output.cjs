const fs = require("fs");
const path = require("path");

const root = path.resolve(__dirname, "..");
const outputDirs = ["dist", "dist-electron", "release"];

for (const dir of outputDirs) {
  const target = path.resolve(root, dir);
  const relative = path.relative(root, target);
  if (relative.startsWith("..") || path.isAbsolute(relative)) {
    throw new Error(`Refusing to clean outside project root: ${target}`);
  }
  fs.rmSync(target, { recursive: true, force: true });
}
