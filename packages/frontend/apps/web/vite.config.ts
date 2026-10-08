import {sqliteServer} from './sqlite-server';
import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';
import {clipperServer} from './clipper-server';
import { pdfAssets } from './pdf-assets';
export default defineConfig({
  plugins: [react(), pdfAssets(), clipperServer(), sqliteServer()],
  optimizeDeps: { include: ['pdfjs-dist'] },
  build: {
    rollupOptions: {
      output: {
        manualChunks(id) {
          if (!id.includes('node_modules')) return;
          if (/\/(react|react-dom|scheduler)\//.test(id)) return 'react';
          if (/\/(yjs|y-prosemirror|y-protocols|lib0)\//.test(id)) return 'crdt';
          if (id.includes('/prosemirror-')) return 'editor';
          if (/\/(markdown-it|entities|linkify-it|mdurl|uc.micro|punycode.js)\//.test(id)) return 'markdown';
        },
      },
    },
  },
});
