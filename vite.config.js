import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';

export default defineConfig({
  plugins: [react()],
  base: '/excel-previewer/',
  build: {
    // Target modern browsers — enables smaller output than the default 'modules'
    // baseline and avoids unnecessary legacy transforms.
    target: 'es2020',
    rollupOptions: {
      output: {
        manualChunks: (id) => {
          // React and react-dom are stable across deploys — putting them in their
          // own chunk means returning users never re-download them after an app update.
          if (id.includes('node_modules/react/') || id.includes('node_modules/react-dom/')) {
            return 'vendor-react';
          }
          // react-router-dom changes rarely — separate chunk for the same reason.
          if (id.includes('node_modules/react-router') || id.includes('node_modules/router')) {
            return 'vendor-router';
          }
          // xlsx stays in the lazy tool chunk automatically (it's only imported
          // inside the excel-previewer tool which is React.lazy'd). No rule needed.
        },
      },
    },
  },
});
