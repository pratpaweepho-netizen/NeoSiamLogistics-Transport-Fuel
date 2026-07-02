import path from 'path';
import { defineConfig } from 'vite';

export default defineConfig(() => {
    return {
      resolve: {
        alias: {
          '@': path.resolve(__dirname, '.'),
        }
      },
      build: {
        rollupOptions: {
          output: {
            // Split the heavy vendors into their own long-lived cache chunks so app changes
            // don't force users to re-download firebase/xlsx.
            manualChunks: {
              firebase: ['firebase/compat/app', 'firebase/compat/database'],
              xlsx: ['xlsx'],
            },
          },
        },
      },
    };
});
