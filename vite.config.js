import { defineConfig } from 'vite';

export default defineConfig({
  // Relative asset URLs so the built dist/ folder works from any sub-path on a static host
  base: './',
  server: {
    port: 3000,
    open: false
  },
  define: {
    // Some libraries might check for process.env
    'process.env': {}
  }
});
