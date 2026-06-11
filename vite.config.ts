import react from "@vitejs/plugin-react";
import { resolve } from "node:path";
import { defineConfig } from "vite";

export default defineConfig(async ({ command }) => ({
  plugins: [react()],
  server: {
    port: 3003,
    strictPort: false,
    https:
      command === "serve"
        ? await (await import("office-addin-dev-certs")).getHttpsServerOptions(365)
        : undefined
  },
  build: {
    outDir: "dist",
    sourcemap: true,
    rollupOptions: {
      input: {
        index: resolve(__dirname, "index.html"),
        taskpane: resolve(__dirname, "taskpane.html")
      }
    }
  }
}));
