import react from "@vitejs/plugin-react";
import { getHttpsServerOptions } from "office-addin-dev-certs";
import { resolve } from "node:path";
import { defineConfig } from "vite";

export default defineConfig(async () => ({
  plugins: [react()],
  server: {
    port: 3003,
    strictPort: false,
    https: await getHttpsServerOptions(365)
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
