import { defineConfig } from 'vite';

export default defineConfig({
  server: {
    port: 3000,
    open: false
  },
  define: {
    // Some libraries might check for process.env
    'process.env': {}
  }
});
