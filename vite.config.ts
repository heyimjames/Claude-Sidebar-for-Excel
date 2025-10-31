import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';
import fs from 'fs';
import path from 'path';
import os from 'os';

const certPath = path.join(os.homedir(), '.office-addin-dev-certs');

export default defineConfig({
  plugins: [react()],
  root: 'src',
  publicDir: '../public',
  build: {
    outDir: '../dist',
    emptyOutDir: true,
    rollupOptions: {
      input: {
        taskpane: path.resolve(__dirname, 'src/taskpane/index.html'),
        commands: path.resolve(__dirname, 'src/commands.html'),
      },
    },
  },
  server: {
    port: 3000,
    https: fs.existsSync(certPath)
      ? {
          key: fs.readFileSync(path.join(certPath, 'localhost.key')),
          cert: fs.readFileSync(path.join(certPath, 'localhost.crt')),
        }
      : undefined,
  },
});
