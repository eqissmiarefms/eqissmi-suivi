import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";

export default defineConfig({
  plugins: [react()],
  base: "/eqissmi-suivi/", // remplacez par le nom exact de votre dépôt GitHub, entre slashes
});