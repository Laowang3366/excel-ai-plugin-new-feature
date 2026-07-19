import path from "node:path";
import { defineConfig, type Plugin } from "vite";
import react from "@vitejs/plugin-react";
import { resolveViteBase } from "./scripts/basePath.mjs";
import { resolveDevHttpsOptions } from "./scripts/viteHttps.mjs";

function officeDevHttpsPlugin(): Plugin {
  return {
    name: "office-dev-https",
    apply: "serve",
    async config() {
      // Only when explicitly not disabled; default HTTPS for Office sideload.
      if (process.env.VITE_DEV_HTTP === "1") {
        return {};
      }
      try {
        const https = await resolveDevHttpsOptions();
        return {
          server: { https },
          preview: { https },
        };
      } catch (err) {
        const message = err instanceof Error ? err.message : String(err);
        throw new Error(
          [
            "Failed to enable HTTPS for Vite dev server via office-addin-dev-certs.",
            message,
            "Install/trust certs: npm run certs:install",
            "Or pure HTTP browser debug: npm run dev:http",
          ].join("\n"),
        );
      }
    },
  };
}

export default defineConfig({
  base: resolveViteBase(process.env),
  plugins: [react(), officeDevHttpsPlugin()],
  resolve: {
    alias: {
      "@shared": path.resolve(__dirname, "shared"),
    },
  },
  server: {
    port: 3000,
    strictPort: true,
  },
  preview: {
    port: 3000,
    strictPort: true,
  },
  build: {
    outDir: "dist",
    emptyOutDir: true,
  },
});
