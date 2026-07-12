import { hashPassword } from "../src/auth.mjs";

const password = process.argv[2];
if (!password || password.length < 12) {
  console.error("用法: npm run hash-password -- <至少12位密码>");
  process.exit(1);
}
console.log(await hashPassword(password));
