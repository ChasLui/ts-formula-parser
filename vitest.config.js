import { defineConfig } from 'vitest/config';

export default defineConfig({
  test: {
    // Test file pattern
    include: ['test/**/*.ts'],
    // Exclude test data files and utility files
    exclude: ['**/testcase.ts', 'test/utils.ts', 'test/formulas/index.ts', 'node_modules/**'],
    // Global timeout settings
    testTimeout: 20000,
    hookTimeout: 10000,
    // Coverage configuration
    coverage: {
      provider: 'v8',
      reporter: ['text', 'html', 'lcov'],
      reportsDirectory: 'docs/coverage',
      include: [
        'formulas/functions/**',
        'formulas/operators.ts',
        'grammar/**',
        'index.ts'
      ],
      exclude: [
        'test/**',
        'docs/**',
        'build/**',
        '*.config.js'
      ]
    },
    // Environment configuration
    environment: 'node',
    // Set globals to false, maintain explicit imports
    globals: false,
    // Support ESM modules
    pool: 'forks'
  }
});