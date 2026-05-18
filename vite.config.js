import { defineConfig, loadEnv } from "vite";
import react from "@vitejs/plugin-react";

// BUG FIX 1: Vite non carica .env in process.env per il config file.
// Bisogna usare loadEnv() esplicitamente.
export default defineConfig(({ mode }) => {
  const env = loadEnv(mode, process.cwd(), "");

  return {
    plugins: [react()],
    server: {
      proxy: {
        "/api/chat": {
          target: "https://api.anthropic.com",
          changeOrigin: true,
          rewrite: () => "/v1/messages",
          configure: (proxy) => {
            proxy.on("proxyReq", (proxyReq) => {
              proxyReq.setHeader("x-api-key", env.ANTHROPIC_API_KEY || "");
              proxyReq.setHeader("anthropic-version", "2023-06-01");
              proxyReq.setHeader("anthropic-beta", "web-search-2025-03-05");
            });
          },
        },
      },
    },
  };
});
