import path from "path"
import tailwindcss from "@tailwindcss/vite"
import react from "@vitejs/plugin-react"
import { defineConfig } from "vite"

// https://vite.dev/config/
export default defineConfig({
  plugins: [react(), tailwindcss()],
  server: {
    proxy: {
      "/v1": {
        target: process.env.VITE_PROXY_TARGET ?? "http://localhost:4000",
        changeOrigin: true,
      },
      "/openai": {
        target: process.env.VITE_PROXY_TARGET ?? "http://localhost:4000",
        changeOrigin: true,
      },
      "/__viz": {
        target: process.env.VITE_PROXY_TARGET ?? "http://localhost:4000",
        changeOrigin: true,
      },
    },
  },
  resolve: {
    alias: {
      "@": path.resolve(__dirname, "./src"),
    },
  },
})
