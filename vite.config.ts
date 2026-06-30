import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";
import basicSsl from "@vitejs/plugin-basic-ssl";

export default defineConfig({
  plugins: [react(), basicSsl()],
  server: {
    port: 5173,
    host: "127.0.0.1",
    https: true
  },
  build: {
    outDir: "dist"
  }
});
