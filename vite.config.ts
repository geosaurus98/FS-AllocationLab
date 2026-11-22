/// <reference types="vitest" />
import { defineConfig, type Plugin } from 'vite';
import react from '@vitejs/plugin-react';

/**
 * `vite build --mode offline` produces ONE self-contained HTML file (all JS/CSS inlined) that advisers
 * can open directly from disk with no server and no network access.
 */
function singleFile(): Plugin {
  return {
    name: 'allocationlab-single-file',
    enforce: 'post',
    generateBundle(_opts, bundle) {
      const html = Object.values(bundle).find((f) => f.type === 'asset' && f.fileName.endsWith('.html'));
      if (!html || html.type !== 'asset') return;
      let source = String(html.source);
      for (const [name, file] of Object.entries(bundle)) {
        if (file.type === 'chunk' && file.isEntry) {
          const re = new RegExp(`<script([^>]*) src="[^"]*${name.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}"[^>]*></script>`);
          const code = file.code.replace(/<\/script/gi, '<\\/script');
          source = source.replace(re, () => `<script type="module">${code}</script>`);
          delete bundle[name];
        } else if (file.type === 'asset' && name.endsWith('.css')) {
          const re = new RegExp(`<link[^>]*href="[^"]*${name.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}"[^>]*>`);
          source = source.replace(re, () => `<style>${String(file.source)}</style>`);
          delete bundle[name];
        }
      }
      html.source = source;
    },
  };
}

export default defineConfig(({ mode }) => ({
  plugins: [react(), ...(mode === 'offline' ? [singleFile()] : [])],
  base: mode === 'offline' ? './' : '/',
  build:
    mode === 'offline'
      ? {
          outDir: 'dist-offline',
          assetsInlineLimit: 100_000_000,
          cssCodeSplit: false,
          modulePreload: false,
          rollupOptions: { output: { inlineDynamicImports: true } },
        }
      : undefined,
  test: {
    globals: true,
    environment: 'node',
    include: ['src/**/*.test.ts', 'src/**/*.test.tsx'],
  },
}));
