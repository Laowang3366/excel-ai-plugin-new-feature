/**
 * Resolve Vite HTTPS server options via office-addin-dev-certs.
 * Only used for `vite` / `vite preview` serve modes — never for build/test.
 */
import fs from "node:fs";
import path from "node:path";
import os from "node:os";

const certDir = path.join(os.homedir(), ".office-addin-dev-certs");
const caPath = path.join(certDir, "ca.crt");
const certPath = path.join(certDir, "localhost.crt");
const keyPath = path.join(certDir, "localhost.key");

export async function resolveDevHttpsOptions() {
  // Prefer existing files first so non-interactive environments can reuse generated
  // certs without forcing a CA trust install (Windows trust is a separate step).
  if (fs.existsSync(certPath) && fs.existsSync(keyPath)) {
    return {
      cert: fs.readFileSync(certPath),
      key: fs.readFileSync(keyPath),
      ...(fs.existsSync(caPath) ? { ca: fs.readFileSync(caPath) } : {}),
    };
  }

  const devCerts = await import("office-addin-dev-certs");
  const options = await devCerts.getHttpsServerOptions();
  return {
    cert: options.cert,
    key: options.key,
    ...(options.ca ? { ca: options.ca } : {}),
  };
}
