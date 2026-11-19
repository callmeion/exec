import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'
import tailwind from '@tailwindcss/vite'
import { viteSingleFile } from 'vite-plugin-singlefile'
// base './' + single file => dist/index.html opens by double-click (file://) and also deploys on IIS
export default defineConfig({ base: './', plugins: [react(), tailwind(), viteSingleFile()] })
