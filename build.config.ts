import { defineBuildConfig } from "unbuild";

const ENTRIES = ["./index.ts"] as const;
const OUT_DIR = "build" as const;
const PACKAGE_NAME = "FormulaParser" as const;

const NODE_TARGET = "es2020" as const;
const BROWSER_TARGET = "es2018" as const;

function createNodeCjsMin() {
  return {
    name: "cjs-min",
    entries: ENTRIES as unknown as string[],
    outDir: OUT_DIR,
    clean: false,
    declaration: false,
    rollup: {
      output: {
        format: "cjs" as const,
        entryFileNames: "index.cjs.min.js",
        exports: "auto" as const,
      },
      esbuild: {
        platform: "node" as const,
        target: NODE_TARGET,
        minify: true,
        treeShaking: true,
      },
    },
  } as const;
}

type BrowserFormat = "umd" | "iife" | "esm";

function createBrowserBundle(format: BrowserFormat, minify = false) {
  const fileSuffix = minify ? ".min" : "";
  const fileName = `index.${format}${fileSuffix}.js`;
  const isNamed = format === "umd" || format === "iife";

  return {
    name: `${format}${minify ? "-min" : ""}`,
    entries: ENTRIES as unknown as string[],
    outDir: OUT_DIR,
    clean: false,
    declaration: false,
    externals: [],
    rollup: {
      inlineDependencies: true,
      output: {
        format,
        ...(isNamed ? { 
          name: PACKAGE_NAME, 
          exports: "auto" as const,
          globals: {
            'bahttext': 'bahttext',
            'bessel': 'bessel', 
            'jstat': 'jStat',
            'chevrotain': 'chevrotain'
          }
        } : { 
          exports: "auto" as const 
        }),
        entryFileNames: fileName,
      },
      esbuild: {
        platform: "browser" as const,
        target: BROWSER_TARGET,
        ...(minify ? { minify: true } : {}),
      },
    },
  } as const;
}

export default defineBuildConfig([
  // Node: emit ESM + CJS (corresponding to main/module/exports)
  {
    entries: ENTRIES as unknown as string[],
    outDir: OUT_DIR,
    clean: true,
    declaration: false,
    rollup: {
      emitCJS: true,
      output: {
        exports: "auto" as const,
      },
    },
  },
  // Node CJS minified
  createNodeCjsMin(),
  // Browser bundles
  createBrowserBundle("umd", false),
  createBrowserBundle("umd", true),
  createBrowserBundle("iife", false),
  createBrowserBundle("iife", true),
  createBrowserBundle("esm", false),
  createBrowserBundle("esm", true),
]);
