import { FormulaParser } from './grammar/hooks';
import { DepParser } from './grammar/dependency/hooks';
import SSF from './ssf/ssf';
import FormulaError from './formulas/error';
import { FormulaHelpers } from './formulas/helpers';

// const funs = new FormulaParser().supportedFunctions();
// console.log('Supported:', funs.join(', '),
//     `\nTotal: ${funs.length}/477, ${funs.length/477*100}% implemented.`);

// Define constants first
const MAX_ROW = 1048576;
const MAX_COLUMN = 16384;

Object.assign(FormulaParser, {
    MAX_ROW,
    MAX_COLUMN,
    SSF,
    DepParser,
    FormulaError,
    FormulaHelpers
});

export default FormulaParser;
export { FormulaParser, DepParser, SSF, FormulaError, FormulaHelpers, MAX_ROW, MAX_COLUMN };