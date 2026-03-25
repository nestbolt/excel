import { defineConfig } from 'vitest/config';

export default defineConfig({
  test: {
    globals: false,
    root: '.',
    include: ['test/**/*.spec.ts'],
    coverage: {
      provider: 'v8',
      include: ['src/**/*.ts'],
      exclude: [
        'src/index.ts',
        'src/**/index.ts',
        'src/interfaces/**',
        'src/**/*.interface.ts',
        'src/**/interfaces.ts',
        'src/storage/storage.types.ts',
        'src/storage/storage-driver.interface.ts',
      ],
    },
  },
});
