/**
 * Portable Vite HTTP-mode detection for dev server.
 * Windows cmd.exe cannot run `FOO=1 vite`; use npm_lifecycle_event or VITE_DEV_HTTP.
 * @param {NodeJS.ProcessEnv | Record<string, string | undefined>} [env]
 * @returns {boolean}
 */
export function shouldUseViteDevHttp(env = process.env) {
  if (env.VITE_DEV_HTTP === "1") return true;
  if (env.npm_lifecycle_event === "dev:http") return true;
  return false;
}
