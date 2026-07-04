import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';
export default defineConfig({
  plugins: [react()],
  build: {
    outDir: 'dist',
    chunkSizeWarningLimit: 1500,
    rollupOptions: { output: { manualChunks: { 'firebase': ['firebase/app','firebase/database'], 'react-core': ['react','react-dom'] } } }
  }
});
