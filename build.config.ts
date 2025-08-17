import { defineBuildConfig } from 'unbuild';

export default defineBuildConfig([
  {
    entries: ['./index'],
    outDir: 'build',
    clean: true,
    declaration: false,
    rollup: {
      emitCJS: true,
    },
  },
  {
    name: 'minified',
    entries: ['./index'],
    outDir: 'build/min',
    clean: true,
    declaration: false,
    rollup: {
      emitCJS: true,
      esbuild: {
        minify: true,
      },
    },
  },
]);


