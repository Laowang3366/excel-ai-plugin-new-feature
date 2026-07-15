import {
  randomBytes,
  scrypt as scryptCallback,
  timingSafeEqual,
} from "node:crypto";
import { promisify } from "node:util";

const scrypt = promisify(scryptCallback);
const SESSION_MAX_AGE_SECONDS = 8 * 60 * 60;

export async function hashPassword(password, salt = randomBytes(16)) {
  const derived = await scrypt(password, salt, 32);
  return `scrypt$${Buffer.from(salt).toString("base64")}$${Buffer.from(derived).toString("base64")}`;
}

export async function verifyPassword(password, encoded) {
  try {
    const [algorithm, saltBase64, hashBase64] = encoded.split("$");
    if (algorithm !== "scrypt" || !saltBase64 || !hashBase64) return false;
    const expected = Buffer.from(hashBase64, "base64");
    const actual = Buffer.from(
      await scrypt(
        password,
        Buffer.from(saltBase64, "base64"),
        expected.length,
      ),
    );
    return (
      actual.length === expected.length && timingSafeEqual(actual, expected)
    );
  } catch {
    return false;
  }
}

export function setAdminSession(reply) {
  const expiresAt = Date.now() + SESSION_MAX_AGE_SECONDS * 1000;
  reply.setCookie(
    "wenge_admin",
    `${expiresAt}.${randomBytes(18).toString("base64url")}`,
    {
      path: "/",
      httpOnly: true,
      sameSite: "strict",
      secure: process.env.NODE_ENV === "production",
      signed: true,
      maxAge: SESSION_MAX_AGE_SECONDS,
    },
  );
}

export function clearAdminSession(reply) {
  reply.clearCookie("wenge_admin", { path: "/" });
}

export function hasValidAdminSession(request) {
  const signed = request.cookies.wenge_admin;
  if (!signed) return false;
  const unsigned = request.unsignCookie(signed);
  if (!unsigned.valid || !unsigned.value) return false;
  const [expiresAt] = unsigned.value.split(".", 1);
  return Number(expiresAt) > Date.now();
}
