import { hashPassword } from "../src/auth.mjs";

if (process.argv[2]) {
  console.error("不要通过命令行参数传递密码；请通过标准输入提供");
  process.exit(1);
}
if (process.stdin.isTTY) {
  console.error("用法: 从安全输入源通过标准输入传入至少 12 位密码");
  process.exit(1);
}
let password = "";
for await (const chunk of process.stdin) password += chunk;
password = password.replace(/[\r\n]+$/, "");
if (!password || password.length < 12) {
  console.error("密码至少需要 12 位");
  process.exit(1);
}
console.log(await hashPassword(password));
