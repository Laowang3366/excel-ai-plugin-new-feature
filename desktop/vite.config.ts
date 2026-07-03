import { defineConfig } from "vite";
import electron from "vite-plugin-electron/simple";
import react from "@vitejs/plugin-react";
import { resolve } from "path";

export default defineConfig({
  plugins: [
    react(),
    electron({
      main: {
        entry: "electron/main.ts",
        vite: {
          build: {
            rolldownOptions: {
              external: ["better-sqlite3"],
            },
          },
        },
      },
      preload: {
        input: "electron/preload.ts",
      },
      renderer: {},
    }),
  ],
  resolve: {
    alias: {
      "@": resolve(__dirname, "src"),
    },
  },
  server: {
    port: 5173,
  },
  build: {
    outDir: "dist",
  },
});
