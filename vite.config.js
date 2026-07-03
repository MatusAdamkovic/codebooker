import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";

// base: "./" makes the built site work from any subfolder
// (GitHub Pages, a university server directory, etc.)
export default defineConfig({
  plugins: [react()],
  base: "./",
});
