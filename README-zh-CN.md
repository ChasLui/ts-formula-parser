![GitHub](https://img.shields.io/github/license/ChasLui/ts-formula-parser)
[![npm (tag)](https://img.shields.io/npm/v/ts-formula-parser/latest)](https://www.npmjs.com/package/ts-formula-parser)
[![npm](https://img.shields.io/npm/dt/ts-formula-parser)](https://www.npmjs.com/package/ts-formula-parser)
[![Coverage Status](https://coveralls.io/repos/github/ChasLui/ts-formula-parser/badge.svg?branch=master)](https://coveralls.io/github/ChasLui/ts-formula-parser?branch=master)
[![CI](https://github.com/ChasLui/ts-formula-parser/actions/workflows/ci.yml/badge.svg)](https://github.com/ChasLui/ts-formula-parser/actions/workflows/ci.yml)

## [快速的 Excel 公式解析器和求值器](https://github.com/ChasLui/ts-formula-parser)

[English](./README.md) | 中文简体

一个快速且可靠的 TypeScript/JavaScript Excel 公式解析器，完全支持 **ESM**。使用 **LL(1)** 解析器。

### 🔥 核心特性

- ✅ **ESM 优先**: 原生 ES Module 支持，包含 TypeScript 定义
- ⚡ **高性能**: 比其他公式解析器快 3 倍
- 🧮 **283+ Excel 函数**: 全面的 Excel 函数支持
- 🏗️ **多种构建格式**: ESM、CJS、UMD、IIFE 及其压缩版本
- 🔒 **类型安全**: 完整的 TypeScript 支持和详细的类型定义
- 📦 **零配置**: 在现代 Node.js (>=22.0.0) 中开箱即用

### [演示](https://lesterlyu.github.io/#/demo/fast-formula-parser)

### [文档](https://chaslui.github.io/ts-formula-parser/index.html)

### [语法图表](https://chaslui.github.io/ts-formula-parser/generated_diagrams.html)

### 支持 283+ 种公式

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

### 包大小

| 格式 | 未压缩 | 压缩版 | Gzip+压缩版 |
|------|---------|--------|-------------|
| ESM | 228KB | 108KB | ~30KB |
| CJS | 230KB | 121KB | ~32KB |
| UMD | 258KB | 109KB | ~30KB |
| IIFE | 258KB | 108KB | ~30KB |

### 系统要求

- **Node.js**: >=22.0.0
- **包管理器**: npm、yarn 或 **pnpm**（推荐）
- **模块系统**: ESM（ES 模块）- 通过构建输出也支持 CommonJS

### 测试框架

本项目使用 **Vitest** 作为测试框架：

- 主要测试文件位于 `test/` 目录
- 公式特定测试位于 `test/formulas/`
- 使用 `@vitest/coverage-v8` 进行覆盖率报告
- 测试数据存储在 JSON 文件中（例如 `test/formulas2.json`）
- 所有测试文件都用 TypeScript 编写以获得更好的类型安全性

主要测试命令：
```bash
# 运行所有测试
pnpm test

# 运行公式特定测试
pnpm test:f

# 以监听模式运行测试
pnpm test:watch

# 生成覆盖率报告
pnpm run coverage
```

### 背景

灵感来自 [XLParser](https://github.com/spreadsheetlab/XLParser/blob/master/src/XLParser/ExcelFormulaGrammar.cs)
和论文 ["A Grammar for Spreadsheet Formulas Evaluated on Two Large Datasets" by Efthimia Aivaloglou, David Hoepelman and Felienne Hermans](https://fenia266781730.files.wordpress.com/2019/01/07335408.pdf)。

注意：我的实现中的语法与他们的不同。我的实现去除了歧义以提升性能。

### 不支持的功能

- [外部引用](https://support.office.com/en-ie/article/create-an-external-reference-link-to-a-cell-range-in-another-workbook-c98d1803-dd75-4668-ac6a-d7cca2a9b95f)
  - 任何包含 `[` 和 `]` 的内容
- 有歧义的旧式语法
  - 工作表名称包含 `:`，例如 `SUM('1003:1856'!D6)`
  - 带有空格但未加引号的工作表名称，例如 `I am a sheet!A1`
- `SUM(Sheet2:Sheet3!A1:C3)`
- 你告诉我的其他情况

### 性能

- **快 3 倍**: 比优化后的 [formula-parser](https://github.com/LesterLyu/formula-parser) 快 3 倍
- **LL(1) 解析**: 最优性能的解析算法
- **优化的包大小**: 各种构建格式都经过优化：
  - ESM: 228KB / 108KB 压缩版
  - CJS: 230KB / 121KB 压缩版  
  - UMD: 258KB / 109KB 压缩版
  - IIFE: 258KB / 108KB 压缩版

### 构建输出

该包为不同用例提供多种构建格式：

| 格式 | 文件 | 大小 | 用途 |
|------|------|------|------|
| ESM | `build/index.mjs` | 228KB | 现代 Node.js/打包工具 |
| CJS | `build/index.cjs` | 230KB | 传统 Node.js |
| UMD | `build/index.umd.min.js` | 109KB | 浏览器全局变量 |
| IIFE | `build/index.iife.min.js` | 108KB | 直接浏览器使用 |
| ESM Browser | `build/index.esm.min.js` | 108KB | 现代浏览器 |

### 依赖项

**核心依赖：**

- `chevrotain`: 解析器构建工具包（词法分析/解析）
- `jstat`: 统计函数
- `bessel`: 数学函数
- `bahttext`: 文本格式化工具

**开发依赖：**

- `vitest` + `@vitest/coverage-v8`: 测试框架和覆盖率报告
- `unbuild`: 构建系统
- `typescript`: 类型定义
- `jsdoc`: 文档生成

### 示例

- 安装

  ```sh
  # 使用 pnpm（推荐）
  pnpm add ts-formula-parser
  
  # 使用 npm
  npm install ts-formula-parser
  
  # 使用 yarn
  yarn add ts-formula-parser
  ```

- 导入（ESM - 推荐）

  ```js
  // 默认导入和命名导出
  import FormulaParser, {
    FormulaHelpers,
    DepParser,
    SSF,
    FormulaError,
    MAX_ROW,
    MAX_COLUMN,
  } from "ts-formula-parser";
  
  // 或只使用命名导入
  import { 
    FormulaParser, 
    FormulaHelpers, 
    FormulaError 
  } from "ts-formula-parser";
  ```

- 导入（CommonJS - 旧版）

  ```js
  const FormulaParser = require("ts-formula-parser");
  const { FormulaHelpers, FormulaError, MAX_ROW, MAX_COLUMN } = FormulaParser;
  ```

- 浏览器使用

  ```html
  <!-- UMD 构建 -->
  <script src="/node_modules/ts-formula-parser/build/index.umd.min.js"></script>
  
  <!-- ESM 构建 -->
  <script type="module">
    import FormulaParser from '/node_modules/ts-formula-parser/build/index.mjs';
  </script>
  ```

- 基本用法

  ```js
  const data = [
    // A  B  C
    [1, 2, 3], // 第 1 行
    [4, 5, 6], // 第 2 行
  ];

  const parser = new FormulaParser({
    // 外部函数，这将覆盖同名的内部函数
    functions: {
      CHAR: (number) => {
        number = FormulaHelpers.accept(number, 'number');
        if (number > 255 || number < 1) throw FormulaError.VALUE;
        return String.fromCharCode(number);
      },
    },

    // 公式中使用的变量（已定义名称）
    // 应该只返回范围引用或单元格引用
    onVariable: (name, sheetName) => {
      // 如果是范围引用 (A1:B2)
      return {
        sheet: "工作表名称",
        from: {
          row: 1,
          col: 1,
        },
        to: {
          row: 2,
          col: 2,
        },
      };
      // 如果是单元格引用 (A1)
      return {
        sheet: "工作表名称",
        row: 1,
        col: 1,
      };
    },

    // 获取单元格值
    onCell: ({ sheet, row, col }) => {
      // 使用基于 1 的索引
      // 返回单元格值，请参阅下一节中的可能类型。
      return data[row - 1][col - 1];
    },

    // 获取范围值
    onRange: (ref) => {
      // 使用基于 1 的索引
      // 当 ref.to.col 是 MAX_COLUMN 或 ref.to.row 是 MAX_ROW 时要小心，
      // 这种方法会导致不必要的循环。
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

  // 计算某些公式需要位置信息，例如 ROW()
  const position = { row: 1, col: 1, sheet: "Sheet1" };

  // 解析公式，需要公式所在位置的信息才能计算某些函数。
  console.log(parser.parse("SUM(A:C)", position));
  // 输出 21

  // 你可以指定返回值是否可以是数组，这在处理数组公式时很有用
  console.log(parser.parse("MMULT({1,5;2,3},{1,2;2,3})", position, true));
  // 输出 [ [ 11, 17 ], [ 8, 13 ] ]
  ```

- 自定义异步函数

  > 记住使用 `await parser.parseAsync(...)` 而不是 `parser.parse(...)`

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
  // 输出 2
  console.log(await parser.parseAsync("SUM(DEMO_FUNC(), 1))", position));
  // 输出 22
  ```

- 需要解析器上下文的自定义函数（例如公式的位置）

  ```js
  const position = { row: 1, col: 1, sheet: "Sheet1" };
  const parser = new FormulaParser({
    functionsNeedContext: {
      // 第一个参数是上下文
      // 后面是传递给函数的参数
      ROW_PLUS_COL: (context, ...args) => {
        return context.position.row + context.position.col;
      },
    },
  });
  console.log(parser.parse("SUM(ROW_PLUS_COL(), 1)", position));
  // 输出 3
  ```

- 解析公式依赖关系

  > 这对构建 `依赖图/树` 很有帮助。

  ```js
  import { DepParser } from "ts-formula-parser";
  const depParser = new DepParser({
    // 如果公式包含变量，onVariable 是你唯一需要提供的
    onVariable: (variable) => {
      return "VAR1" === variable
        ? { from: { row: 1, col: 1 }, to: { row: 2, col: 2 } }
        : { row: 1, col: 1 };
    },
  });

  // 应该提供公式的位置
  const position = { row: 1, col: 1, sheet: "Sheet1" };

  // 返回引用数组（范围引用或单元格引用）
  // 这会返回 [{row: 1, col: 1, sheet: 'Sheet1'}]
  depParser.parse("A1+1", position);

  // 这会返回 [{sheet: 'Sheet1', from: {row: 1, col: 1}, to: {row: 3, col: 3}}]
  depParser.parse("A1:C3", position);

  // 这会返回 [{from: {row: 1, col: 1}, to: {row: 2, col: 2}}]
  depParser.parse("VAR1 + 1", position);

  // 复杂公式
  depParser.parse(
    'IF(MONTH($K$1)<>MONTH($K$1-(WEEKDAY($K$1,1)-(start_day-1))-IF((WEEKDAY($K$1,1)-(start_day-1))<=0,7,0)+(ROW(O5)-ROW($K$3))*7+(COLUMN(O5)-COLUMN($K$3)+1)),"",$K$1-(WEEKDAY($K$1,1)-(start_day-1))-IF((WEEKDAY($K$1,1)-(start_day-1))<=0,7,0)+(ROW(O5)-ROW($K$3))*7+(COLUMN(O5)-COLUMN($K$3)+1))',
    position
  );
  // 这会返回以下结果
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

### JavaScript 中的公式数据类型

> 以下数据类型用于 Excel 公式，这些是公式或函数可以返回的唯一有效数据类型。

- 数字（日期使用数字）：`1234`
- 字符串：`'some string'`
- 布尔值：`true`、`false`
- 数组：`[[1, 2, true, 'str']]`
- 范围引用：（基于 1 的索引）

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

- 单元格引用：（基于 1 的索引）

  ```js
  const ref = {
      sheet: String,
      row: Number,
      col: Number,
  }
  ```

- [联合体（例如 (A1:C3, E1:G6)）](https://github.com/LesterLyu/fast-formula-parser/blob/master/grammar/type/collection.js)
- [FormulaError](https://lesterlyu.github.io/fast-formula-parser/FormulaError.html)
  - `FormulaError.DIV0`：`#DIV/0!`
  - `FormulaError.NA`：`#N/A`
  - `FormulaError.NAME`：`#NAME?`
  - `FormulaError.NULL`：`#NULL!`
  - `FormulaError.NUM`：`#NUM!`
  - `FormulaError.REF`：`#REF!`
  - `FormulaError.VALUE`：`#VALUE!`

### TypeScript 支持

> 完整的 TypeScript 支持和全面的类型定义

```typescript
import FormulaParser, { 
  FormulaHelpers, 
  FormulaError, 
  MAX_ROW, 
  MAX_COLUMN,
  DepParser,
  SSF 
} from "ts-formula-parser";

// 所有类型都有正确的定义
const parser = new FormulaParser({
  onCell: ({ sheet, row, col }) => {
    // TypeScript 知道此对象的确切形状
    return data[row - 1][col - 1];
  },
  
  functions: {
    CUSTOM_FUNC: (arg1: number, arg2: string): number => {
      // 具有正确类型的自定义函数
      return FormulaHelpers.accept(arg1, 'number') as number;
    }
  }
});

// 使用位置上下文解析
const result = parser.parse('SUM(A1:C3)', {
  sheet: 'Sheet1',
  row: 1, 
  col: 1
});
```

### 错误处理

- 词法/解析错误

  > 错误位置可通过 `error.details.errorLocation` 获得

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

- 来自内部/外部函数的错误或解析器的意外错误

  > 错误将被包装到 `FormulaError` 中。确切的错误在 `error.details` 中。

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

### 开发

本项目使用 **pnpm** 进行包管理，**TypeScript** 确保类型安全，**Vitest** 进行测试，**unbuild** 创建多种输出格式。

```bash
# 安装依赖
pnpm install

# 运行测试（使用 Vitest）
pnpm test

# 监听模式运行测试
pnpm test:watch

# 构建所有格式
pnpm build

# 开发构建（stub 模式）
pnpm build:dev

# 类型检查
pnpm typecheck

# 代码检查
pnpm lint

# 生成文档
pnpm run docs

# 覆盖率报告
pnpm run coverage

# 性能测试
pnpm run perf

# CI 流水线（类型检查 + 测试 + 构建）
pnpm run ci
```

### 从 fast-formula-parser 迁移

如果你正在从 `fast-formula-parser` 迁移，以下是主要变更：

1. **包名**: `fast-formula-parser` → `ts-formula-parser`
2. **ESM 优先**: 更新你的导入以使用 ES 模块
3. **Node.js 版本**: 需要 Node.js >=22.0.0
4. **TypeScript**: 包含完整的类型定义
5. **构建路径**: 更新的构建输出路径

```js
// 旧版（fast-formula-parser）
const FormulaParser = require('fast-formula-parser');

// 新版（ts-formula-parser）
import FormulaParser from 'ts-formula-parser';
```

### 致谢

- 从 [LesterLyu/fast-formula-parser](https://github.com/LesterLyu/fast-formula-parser) 仓库 Fork 而来
- 使用 [Chevrotain](https://github.com/SAP/chevrotain) 解析器工具包构建
- 使用 [unbuild](https://github.com/unjs/unbuild) 现代构建系统
