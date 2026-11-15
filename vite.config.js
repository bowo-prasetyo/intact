import { defineConfig } from 'vite';
import vue from '@vitejs/plugin-vue';

// `base: './'` makes the build work from any GitHub Pages sub-path (https://<user>.github.io/<repo>/).
export default defineConfig({
  base: './',
  plugins: [vue()],
});
