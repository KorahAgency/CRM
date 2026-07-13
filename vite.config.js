import { defineConfig } from "vite"
import react from "@vitejs/plugin-react"

export default defineConfig({
  plugins: [react()],
  build: {
    chunkSizeWarningLimit: 700,
    rollupOptions: {
      output: {
        manualChunks: {
          charts: ["recharts"],
          supabase: ["@supabase/supabase-js"],
          files: ["papaparse", "read-excel-file"]
        }
      }
    }
  }
})
