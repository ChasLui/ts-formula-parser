import { defineConfig } from 'vitest/config';

export default defineConfig({
  test: {
    // 测试文件模式
    include: ['test/**/*.js'],
    // 排除测试数据文件和工具文件
    exclude: ['**/testcase.js', 'test/utils.js', 'test/formulas/index.js', 'node_modules/**'],
    // 全局超时设置
    testTimeout: 20000,
    hookTimeout: 10000,
    // 覆盖率配置
    coverage: {
      provider: 'v8',
      reporter: ['text', 'html', 'lcov'],
      reportsDirectory: 'docs/coverage',
      include: [
        'formulas/functions/**',
        'formulas/operators.js',
        'grammar/**',
        'index.js'
      ],
      exclude: [
        'test/**',
        'docs/**',
        'build/**',
        '*.config.js'
      ]
    },
    // 环境配置
    environment: 'node',
    // 设置globals为false，保持显式导入
    globals: false,
    // 支持ESM模块
    pool: 'forks'
  }
});