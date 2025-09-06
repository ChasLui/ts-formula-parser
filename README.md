![GitHub](https://img.shields.io/github/license/ChasLui/ts-formula-parser)
[![npm (tag)](https://img.shields.io/npm/v/ts-formula-parser/latest)](https://www.npmjs.com/package/ts-formula-parser)
[![npm](https://img.shields.io/npm/dt/ts-formula-parser)](https://www.npmjs.com/package/ts-formula-parser)
[![Coverage Status](https://coveralls.io/repos/github/ChasLui/ts-formula-parser/badge.svg?branch=master)](https://coveralls.io/github/ChasLui/ts-formula-parser?branch=master)
[![CI](https://github.com/ChasLui/ts-formula-parser/actions/workflows/ci.yml/badge.svg)](https://github.com/ChasLui/ts-formula-parser/actions/workflows/ci.yml)

## [A Fast Excel Formula Parser & Evaluator](https://github.com/ChasLui/ts-formula-parser)

English | [中文简体](./README-zh-CN.md)

A fast and reliable excel formula parser in TypeScript/JavaScript with full **ESM** support. Using **LL(1)** parser.

### 🔥 Key Features

- ✅ **ESM First**: Native ES Module support with TypeScript definitions
- ⚡ **High Performance**: 3x faster than other formula parsers
- 🧮 **283+ Excel Functions**: Comprehensive Excel function support
- 🏗️ **Multiple Build Formats**: ESM, CJS, UMD, IIFE with minified versions
- 🔒 **Type Safe**: Full TypeScript support with detailed type definitions
- 📦 **Zero Config**: Works out-of-the-box with modern Node.js (>=22.0.0)

### [Demo](https://lesterlyu.github.io/#/demo/fast-formula-parser)

### [Documentation](https://chaslui.github.io/ts-formula-parser/index.html)

### [Grammar Diagram](https://chaslui.github.io/ts-formula-parser/generated_diagrams.html)

### Supports 283+ Formulas

```
ABS, ACOS, ACOSH, ACOT, ACOTH, ADDRESS, AND, ARABIC, AREAS, ASC, ASIN, ASINH, ATAN, ATAN2, ATANH, AVEDEV, AVERAGE, AVERAGEA, AVERAGEIF, BAHTTEXT, BASE, BESSELI, BESSELJ, BESSELK, BESSELY, BETA.DIST, BETA.INV, BIN2DEC, BIN2HEX, BIN2OCT, BINOM.DIST, BINOM.DIST.RANGE, BINOM.INV, BITAND, BITLSHIFT, BITOR,
BITRSHIFT, BITXOR, CEILING, CEILING.MATH, CEILING.PRECISE, CHAR, CHISQ.DIST, CHISQ.DIST.RT, CHISQ.INV, CHISQ.INV.RT, CHISQ.TEST, CLEAN, CODE, COLUMN, COLUMNS, COMBIN, COMBINA, COMPLEX, CONCAT, CONCATENATE, CONFIDENCE.NORM, CONFIDENCE.T, CORREL, COS, COSH, COT, COTH, COUNT, COUNTIF, COVARIANCE.P,
COVARIANCE.S, CSC, CSCH, DATE, DATEDIF, DATEVALUE, DAY, DAYS, DAYS360, DBCS, DEC2BIN, DEC2HEX, DEC2OCT, DECIMAL, DEGREES, DELTA, DEVSQ, DOLLAR, EDATE, ENCODEURL, EOMONTH, ERF, ERFC, ERROR.TYPE, EVEN, EXACT, EXP, EXPON.DIST, F.DIST, F.DIST.RT, F.INV, F.INV.RT, F.TEST, FACT, FACTDOUBLE, FALSE, FIND, FINDB,
FISHER, FISHERINV, FIXED, FLOOR, FLOOR.MATH, FLOOR.PRECISE, FORECAST, FORECAST.LINEAR, FREQUENCY, GAMMA, GAMMA.DIST, GAMMA.INV, GAMMALN, GAMMALN.PRECISE, GAUSS, GCD, GEOMEAN, GESTEP, GROWTH, HARMEAN, HEX2BIN, HEX2DEC, HEX2OCT, HLOOKUP, HOUR, HYPGEOM.DIST, IF, IFERROR, IFNA, IFS, IMABS, IMAGINARY, IMARGUMENT,
IMCONJUGATE, IMCOS, IMCOSH, IMCOT, IMCSC, IMCSCH, IMDIV, IMEXP, IMLN, IMLOG10, IMLOG2, IMPOWER, IMPRODUCT, IMREAL, IMSEC, IMSECH, IMSIN, IMSINH, IMSQRT, IMSUB, IMSUM, IMTAN, INDEX, INT, INTERCEPT, ISBLANK, ISERR, ISERROR, ISEVEN, ISLOGICAL, ISNA, ISNONTEXT, ISNUMBER, ISO.CEILING, ISOWEEKNUM, ISREF, ISTEXT,
KURT, LCM, LEFT, LEFTB, LN, LOG, LOG10, LOGNORM.DIST, LOGNORM.INV, LOWER, MDETERM, MID, MIDB, MINUTE, MMULT, MOD, MONTH, MROUND, MULTINOMIAL, MUNIT, N, NA, NEGBINOM.DIST, NETWORKDAYS, NETWORKDAYS.INTL, NORM.DIST, NORM.INV, NORM.S.DIST, NORM.S.INV, NOT, NOW, NUMBERVALUE, OCT2BIN, OCT2DEC, OCT2HEX, ODD, OR,
PHI, PI, POISSON.DIST, POWER, PRODUCT, PROPER, QUOTIENT, RADIANS, RAND, RANDBETWEEN, REPLACE, REPLACEB, REPT, RIGHT, RIGHTB, ROMAN, ROUND, ROUNDDOWN, ROUNDUP, ROW, ROWS, SEARCH, SEARCHB, SEC, SECH, SECOND, SERIESSUM, SIGN, SIN, SINH, SQRT, SQRTPI, STANDARDIZE, SUM, SUMIF, SUMPRODUCT, SUMSQ, SUMX2MY2,
SUMX2PY2, SUMXMY2, T, T.DIST, T.DIST.2T, T.DIST.RT, T.INV, T.INV.2T, TAN, TANH, TEXT, TIME, TIMEVALUE, TODAY, TRANSPOSE, TRIM, UPPER, TRUE, TRUNC, TYPE, UNICHAR, UNICODE, VLOOKUP, WEBSERVICE, WEEKDAY, WEEKNUM, WEIBULL.DIST, WORKDAY, WORKDAY.INTL, XOR, YEAR, YEARFRAC
```

### Bundle Sizes

| Format | Uncompressed | Minified | Gzipped+Minified |
|--------|-------------|----------|------------------|
| ESM | 228KB | 108KB | ~30KB |
| CJS | 230KB | 121KB | ~32KB |
| UMD | 258KB | 109KB | ~30KB |
| IIFE | 258KB | 108KB | ~30KB |

### Requirements

- **Node.js**: >=22.0.0
- **Package Manager**: npm, yarn, or **pnpm** (recommended)
- **Module System**: ESM (ES Modules) - CommonJS also supported via build outputs

### Testing Framework

This project uses **Vitest** as the testing framework:

- Main test files located in `test/` directory
- Formula-specific tests located in `test/formulas/`
- Uses `@vitest/coverage-v8` for coverage reporting
- Test data stored in JSON files (e.g., `test/formulas2.json`)
- All test files are written in TypeScript for better type safety

Key testing commands:
```bash
# Run all tests
pnpm test

# Run formula-specific tests
pnpm test:f

# Run tests in watch mode
pnpm test:watch

# Generate coverage report
pnpm run coverage
```

### Background

Inspired by [XLParser](https://github.com/spreadsheetlab/XLParser/blob/master/src/XLParser/ExcelFormulaGrammar.cs)
and the paper ["A Grammar for Spreadsheet Formulas Evaluated on Two Large Datasets" by Efthimia Aivaloglou, David Hoepelman and Felienne Hermans](https://fenia266781730.files.wordpress.com/2019/01/07335408.pdf).

Note: The grammar in my implementation is different from theirs. My implementation gets rid of ambiguities to boost the performance.

### What is not supported

- [External reference](https://support.office.com/en-ie/article/create-an-external-reference-link-to-a-cell-range-in-another-workbook-c98d1803-dd75-4668-ac6a-d7cca2a9b95f)
  - Anything with `[` and `]`
- Ambiguous old styles
  - Sheet name contains `:`, e.g. `SUM('1003:1856'!D6)`
  - Sheet name with space that is not quoted, e.g. `I am a sheet!A1`
- `SUM(Sheet2:Sheet3!A1:C3)`
- You tell me

### Performance

- **3x faster** than the optimized [formula-parser](https://github.com/LesterLyu/formula-parser)
- **LL(1) parsing** for optimal performance
- **Optimized bundle sizes** across all build formats:
  - ESM: 228KB / 108KB minified
  - CJS: 230KB / 121KB minified  
  - UMD: 258KB / 109KB minified
  - IIFE: 258KB / 108KB minified

### Dependency

- [Chevrotain](https://github.com/SAP/chevrotain) , thanks to this great parser building toolkit.

### Build Outputs

The package provides multiple build formats for different use cases:

| Format | File | Size | Use Case |
|--------|------|------|----------|
| ESM | `build/index.mjs` | 228KB | Modern Node.js/bundlers |
| CJS | `build/index.cjs` | 230KB | Legacy Node.js |
| UMD | `build/index.umd.min.js` | 109KB | Browser globals |
| IIFE | `build/index.iife.min.js` | 108KB | Direct browser use |
| ESM Browser | `build/index.esm.min.js` | 108KB | Modern browsers |

### Examples

- Install

  ```sh
  # Using pnpm (recommended)
  pnpm add ts-formula-parser
  
  # Using npm
  npm install ts-formula-parser
  
  # Using yarn
  yarn add ts-formula-parser
  ```

- Import (ESM - Recommended)

  ```js
  // Default import with named exports
  import FormulaParser, {
    FormulaHelpers,
    DepParser,
    SSF,
    FormulaError,
    MAX_ROW,
    MAX_COLUMN,
  } from "ts-formula-parser";
  
  // Or named imports only
  import { 
    FormulaParser, 
    FormulaHelpers, 
    FormulaError 
  } from "ts-formula-parser";
  ```

- Import (CommonJS - Legacy)

  ```js
  const FormulaParser = require("ts-formula-parser");
  const { FormulaHelpers, FormulaError, MAX_ROW, MAX_COLUMN } = FormulaParser;
  ```

- Browser Usage

  ```html
  <!-- UMD build -->
  <script src="/node_modules/ts-formula-parser/build/index.umd.min.js"></script>
  
  <!-- ESM build -->
  <script type="module">
    import FormulaParser from '/node_modules/ts-formula-parser/build/index.mjs';
  </script>
  ```

- Basic Usage

  ```js
  const data = [
    // A  B  C
    [1, 2, 3], // row 1
    [4, 5, 6], // row 2
  ];

  const parser = new FormulaParser({
    // External functions, this will override internal functions with same name
    functions: {
      CHAR: (number) => {
        number = FormulaHelpers.accept(number, 'number');
        if (number > 255 || number < 1) throw FormulaError.VALUE;
        return String.fromCharCode(number);
      },
    },

    // Variable used in formulas (defined name)
    // Should only return range reference or cell reference
    onVariable: (name, sheetName) => {
      // If it is a range reference (A1:B2)
      return {
        sheet: "sheet name",
        from: {
          row: 1,
          col: 1,
        },
        to: {
          row: 2,
          col: 2,
        },
      };
      // If it is a cell reference (A1)
      return {
        sheet: "sheet name",
        row: 1,
        col: 1,
      };
    },

    // retrieve cell value
    onCell: ({ sheet, row, col }) => {
      // using 1-based index
      // return the cell value, see possible types in next section.
      return data[row - 1][col - 1];
    },

    // retrieve range values
    onRange: (ref) => {
      // using 1-based index
      // Be careful when ref.to.col is MAX_COLUMN or ref.to.row is MAX_ROW, this will result in
      // unnecessary loops in this approach.
      const arr = [];
      for (let row = ref.from.row; row <= ref.to.row; row++) {
        const innerArr = [];
        if (data[row - 1]) {
          for (let col = ref.from.col; col <= ref.to.col; col++) {
            innerArr.push(data[row - 1][col - 1]);
          }
        }
        arr.push(innerArr);
      }
      return arr;
    },
  });

  // position is required for evaluating certain formulas, e.g. ROW()
  const position = { row: 1, col: 1, sheet: "Sheet1" };

  // parse the formula, the position of where the formula is located is required
  // for some functions.
  console.log(parser.parse("SUM(A:C)", position));
  // print 21

  // you can specify if the return value can be an array, this is helpful when dealing
  // with an array formula
  console.log(parser.parse("MMULT({1,5;2,3},{1,2;2,3})", position, true));
  // print [ [ 11, 17 ], [ 8, 13 ] ]
  ```

- Custom Async functions

  > Remember to use `await parser.parseAsync(...)` instead of `parser.parse(...)`

  ```js
  const position = { row: 1, col: 1, sheet: "Sheet1" };
  const parser = new FormulaParser({
    onCell: (ref) => {
      return 1;
    },
    functions: {
      DEMO_FUNC: async () => {
        return [
          [1, 2, 3],
          [4, 5, 6],
        ];
      },
    },
  });
  console.log(await parser.parseAsync("A1 + IMPORT_CSV())", position));
  // print 2
  console.log(await parser.parseAsync("SUM(DEMO_FUNC(), 1))", position));
  // print 22
  ```

- Custom function requires parser context (e.g. location of the formula)

  ```js
  const position = { row: 1, col: 1, sheet: "Sheet1" };
  const parser = new FormulaParser({
    functionsNeedContext: {
      // the first argument is the context
      // the followings are the arguments passed to the function
      ROW_PLUS_COL: (context, ...args) => {
        return context.position.row + context.position.col;
      },
    },
  });
  console.log(parser.parse("SUM(ROW_PLUS_COL(), 1)", position));
  // print 3
  ```

### Development

This project uses **pnpm** for package management, **TypeScript** for type safety, **Vitest** for testing, and **unbuild** for creating multiple output formats.

```bash
# Install dependencies
pnpm install

# Run tests (using Vitest)
pnpm test

# Run tests in watch mode
pnpm test:watch

# Build all formats
pnpm build

# Development build (stub mode)
pnpm build:dev

# Type checking
pnpm typecheck

# Lint code
pnpm lint

# Generate documentation
pnpm run docs

# Coverage report
pnpm run coverage

# Performance test
pnpm run perf

# CI pipeline (typecheck + test + build)
pnpm run ci
```

- Parse Formula Dependency

  > This is helpful for building `dependency graph/tree`.

  ```js
  import { DepParser } from "ts-formula-parser";
  const depParser = new DepParser({
    // onVariable is the only thing you need provide if the formula contains variables
    onVariable: (variable) => {
      return "VAR1" === variable
        ? { from: { row: 1, col: 1 }, to: { row: 2, col: 2 } }
        : { row: 1, col: 1 };
    },
  });

  // position of the formula should be provided
  const position = { row: 1, col: 1, sheet: "Sheet1" };

  // Return an array of references (range reference or cell reference)
  // This gives [{row: 1, col: 1, sheet: 'Sheet1'}]
  depParser.parse("A1+1", position);

  // This gives [{sheet: 'Sheet1', from: {row: 1, col: 1}, to: {row: 3, col: 3}}]
  depParser.parse("A1:C3", position);

  // This gives [{from: {row: 1, col: 1}, to: {row: 2, col: 2}}]
  depParser.parse("VAR1 + 1", position);

  // Complex formula
  depParser.parse(
    'IF(MONTH($K$1)<>MONTH($K$1-(WEEKDAY($K$1,1)-(start_day-1))-IF((WEEKDAY($K$1,1)-(start_day-1))<=0,7,0)+(ROW(O5)-ROW($K$3))*7+(COLUMN(O5)-COLUMN($K$3)+1)),"",$K$1-(WEEKDAY($K$1,1)-(start_day-1))-IF((WEEKDAY($K$1,1)-(start_day-1))<=0,7,0)+(ROW(O5)-ROW($K$3))*7+(COLUMN(O5)-COLUMN($K$3)+1))',
    position
  );
  // This gives the following result
  const result = [
    {
      col: 11,
      row: 1,
      sheet: "Sheet1",
    },
    {
      col: 1,
      row: 1,
      sheet: "Sheet1",
    },
    {
      col: 15,
      row: 5,
      sheet: "Sheet1",
    },
    {
      col: 11,
      row: 3,
      sheet: "Sheet1",
    },
  ];
  ```

### Formula data types in JavaScript

> The following data types are used in excel formulas and these are the only valid data types a formula or a function can return.

- Number (date uses number): `1234`
- String: `'some string'`
- Boolean: `true`, `false`
- Array: `[[1, 2, true, 'str']]`
- Range Reference: (1-based index)

  ```js
  const ref = {
      sheet: String,
      from: {
          row: Number,
          col: Number,
      },
      to: {
          row: Number,
          col: Number,
      },
  }
  ```

- Cell Reference: (1-based index)

  ```js
  const ref = {
      sheet: String,
      row: Number,
      col: Number,
  }
  ```

- [Union (e.g. (A1:C3, E1:G6))](https://github.com/LesterLyu/fast-formula-parser/blob/master/grammar/type/collection.js)
- [FormulaError](https://lesterlyu.github.io/fast-formula-parser/FormulaError.html)
  - `FormulaError.DIV0`: `#DIV/0!`
  - `FormulaError.NA`: `#N/A`
  - `FormulaError.NAME`: `#NAME?`
  - `FormulaError.NULL`: `#NULL!`
  - `FormulaError.NUM`: `#NUM!`
  - `FormulaError.REF`: `#REF!`
  - `FormulaError.VALUE`: `#VALUE!`

### TypeScript Support

> Full TypeScript support with comprehensive type definitions

```typescript
import FormulaParser, { 
  FormulaHelpers, 
  FormulaError, 
  MAX_ROW, 
  MAX_COLUMN,
  DepParser,
  SSF 
} from "ts-formula-parser";

// All types are properly defined
const parser = new FormulaParser({
  onCell: ({ sheet, row, col }) => {
    // TypeScript knows the exact shape of this object
    return data[row - 1][col - 1];
  },
  
  functions: {
    CUSTOM_FUNC: (arg1: number, arg2: string): number => {
      // Custom functions with proper typing
      return FormulaHelpers.accept(arg1, 'number') as number;
    }
  }
});

// Parse with position context
const result = parser.parse('SUM(A1:C3)', {
  sheet: 'Sheet1',
  row: 1, 
  col: 1
});
```

### Error handling

- Lexing/Parsing Error

  > Error location is available at `error.details.errorLocation`

  ```js
  try {
    parser.parse("SUM(1))", position);
  } catch (e) {
    console.log(e);
    // #ERROR!:
    // SUM(1))
    //       ^
    // Error at position 1:7
    // Redundant input, expecting EOF but found: )

    expect(e).to.be.instanceof(FormulaError);
    expect(e.details.errorLocation.line).to.eq(1);
    expect(e.details.errorLocation.column).to.eq(7);
    expect(e.name).to.eq("#ERROR!");
    expect(e.details.name).to.eq("NotAllInputParsedException");
  }
  ```

- Error from internal/external functions or unexpected error from the parser

  > The error will be wrapped into `FormulaError`. The exact error is in `error.details`.

  ```js
  const parser = new FormulaParser({
    functions: {
      BAD_FN: () => {
        throw new SyntaxError();
      },
    },
  });

  try {
    parser.parse("SUM(1))", position);
  } catch (e) {
    expect(e).to.be.instanceof(FormulaError);
    expect(e.name).to.eq("#ERROR!");
    expect(e.details.name).to.eq("SyntaxError");
  }
  ```

### Migration from fast-formula-parser

If you're migrating from `fast-formula-parser`, here are the key changes:

1. **Package name**: `fast-formula-parser` → `ts-formula-parser`
2. **ESM First**: Update your imports to use ES modules
3. **Node.js version**: Requires Node.js >=22.0.0
4. **TypeScript**: Full type definitions included
5. **Bundle paths**: Updated build output paths

```js
// Old (fast-formula-parser)
const FormulaParser = require('fast-formula-parser');

// New (ts-formula-parser)
import FormulaParser from 'ts-formula-parser';
```

### Thanks

- Forked from the [LesterLyu/fast-formula-parser](https://github.com/LesterLyu/fast-formula-parser) repo
- Built with [Chevrotain](https://github.com/SAP/chevrotain) parser toolkit
- Uses [unbuild](https://github.com/unjs/unbuild) for modern build system
