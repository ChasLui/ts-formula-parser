import FormulaError from '../formulas/error';
import type { ErrorDetails } from '../types';
import { Address } from '../formulas/helpers';
import { NotAllInputParsedException } from 'chevrotain';
import { Prefix, Postfix, Infix, Operators, CompareOperator, ConcatOperator, MathOperator } from '../formulas/operators';
import Collection from './type/collection';

const MAX_ROW = 1048576;
const MAX_COLUMN = 16384;

// Define interfaces for cell/range references
interface CellRef {
    row: number;
    col: number;
    sheet?: string | number;
    address?: string;
}

interface RangeRef {
    from: CellRef;
    to: CellRef;
    sheet?: string | number;
}

interface RefObject {
    ref: CellRef | RangeRef;
}

interface ExtractedValue {
    val: any;
    isArray: boolean;
}

// Define context interface for both FormulaParser and DepParser
interface ParserContext {
    retrieveRef(valueOrRef: any): any;
    async?: boolean;
}

class Utils {
    private context: ParserContext;

    constructor(context: ParserContext) {
        this.context = context;
    }

    columnNameToNumber(columnName: string): number {
        return Address.columnNameToNumber(columnName);
    }

    /**
     * Parse the cell address only.
     * @param cellAddress - Cell address string
     * @return Object with ref property containing cell information
     */
    parseCellAddress(cellAddress: string): RefObject {
        const res = cellAddress.match(/([$]?)([A-Za-z]{1,3})([$]?)([1-9][0-9]*)/);
        if (!res) {
            throw new Error(`Invalid cell address: ${cellAddress}`);
        }
        
        return {
            ref: {
                address: res[0],
                col: this.columnNameToNumber(res[2]),
                row: +res[4]
            },
        };
    }

    parseRow(row: string | number): RefObject {
        const rowNum = +row;
        if (!Number.isInteger(rowNum)) {
            throw Error('Row number must be integer.');
        }
        return {
            ref: {
                col: undefined as any,
                row: +row
            },
        };
    }

    parseCol(col: string): RefObject {
        return {
            ref: {
                col: this.columnNameToNumber(col),
                row: undefined as any,
            },
        };
    }

    parseColRange(col1: string, col2: string): RefObject {
        const col1Num = this.columnNameToNumber(col1);
        const col2Num = this.columnNameToNumber(col2);
        return {
            ref: {
                from: {
                    col: Math.min(col1Num, col2Num),
                    row: null as any
                },
                to: {
                    col: Math.max(col1Num, col2Num),
                    row: null as any
                }
            } as RangeRef
        };
    }

    parseRowRange(row1: number, row2: number): RefObject {
        return {
            ref: {
                from: {
                    col: null as any,
                    row: Math.min(row1, row2),
                },
                to: {
                    col: null as any,
                    row: Math.max(row1, row2),
                }
            } as RangeRef
        };
    }

    private _applyPrefix(prefixes: string[], val: any, isArray: boolean): any {
        if (this.isFormulaError(val)) {
            return val;
        }
        return Prefix.unaryOp(prefixes, val, isArray);
    }

    async applyPrefixAsync(prefixes: string[], value: Promise<any>): Promise<any> {
        const { val, isArray } = this.extractRefValue(await value);
        return this._applyPrefix(prefixes, val, isArray);
    }

    /**
     * Apply + or - unary prefix.
     * @param prefixes - Array of prefix operators
     * @param value - Value to apply prefixes to
     * @return Modified value
     */
    applyPrefix(prefixes: string[], value: any): any {
        if (this.context.async) {
            return this.applyPrefixAsync(prefixes, value);
        } else {
            const { val, isArray } = this.extractRefValue(value);
            return this._applyPrefix(prefixes, val, isArray);
        }
    }

    private _applyPostfix(val: any, isArray: boolean, postfix: string): any {
        if (this.isFormulaError(val)) {
            return val;
        }
        return Postfix.percentOp(val, postfix, isArray);
    }

    async applyPostfixAsync(value: Promise<any>, postfix: string): Promise<any> {
        const { val, isArray } = this.extractRefValue(await value);
        return this._applyPostfix(val, isArray, postfix);
    }

    applyPostfix(value: any, postfix: string): any {
        if (this.context.async) {
            return this.applyPostfixAsync(value, postfix);
        } else {
            const { val, isArray } = this.extractRefValue(value);
            return this._applyPostfix(val, isArray, postfix);
        }
    }

    private _applyInfix(res1: ExtractedValue, infix: string, res2: ExtractedValue): any {
        const val1 = res1.val, isArray1 = res1.isArray;
        const val2 = res2.val, isArray2 = res2.isArray;
        if (this.isFormulaError(val1)) {
            return val1;
        }
        if (this.isFormulaError(val2)) {
            return val2;
        }
        if (Operators.compareOp.includes(infix as CompareOperator)) {
            return Infix.compareOp(val1, infix, val2, isArray1, isArray2);
        } else if (Operators.concatOp.includes(infix as ConcatOperator)) {
            return Infix.concatOp(val1, infix, val2, isArray1, isArray2);
        } else if (Operators.mathOp.includes(infix as MathOperator)) {
            return Infix.mathOp(val1, infix, val2, isArray1, isArray2);
        } else {
            throw new Error(`Unrecognized infix: ${infix}`);
        }
    }

    async applyInfixAsync(value1: Promise<any>, infix: string, value2: Promise<any>): Promise<any> {
        const res1 = this.extractRefValue(await value1);
        const res2 = this.extractRefValue(await value2);
        return this._applyInfix(res1, infix, res2);
    }

    applyInfix(value1: any, infix: string, value2: any): any {
        if (this.context.async) {
            return this.applyInfixAsync(value1, infix, value2);
        } else {
            const res1 = this.extractRefValue(value1);
            const res2 = this.extractRefValue(value2);
            return this._applyInfix(res1, infix, res2);
        }
    }

    applyIntersect(refs: any[]): any {
        if (this.isFormulaError(refs[0])) {
            return refs[0];
        }
        if (!refs[0].ref) {
            throw Error(`Expecting a reference, but got ${refs[0]}.`);
        }
        // a intersection will keep track of references, value won't be retrieved here.
        let maxRow: number, maxCol: number, minRow: number, minCol: number, sheet: string | number | undefined, res: RefObject;
        
        // first time setup
        const ref = refs.shift().ref as CellRef | RangeRef;
        sheet = ref.sheet;
        if (!('from' in ref)) {
            // check whole row/col reference
            if ((ref as CellRef).row === undefined || (ref as CellRef).col === undefined) {
                throw Error('Cannot intersect the whole row or column.');
            }

            // cell ref
            maxRow = minRow = (ref as CellRef).row;
            maxCol = minCol = (ref as CellRef).col;
        } else {
            // range ref
            maxRow = Math.max(ref.from.row, ref.to.row);
            minRow = Math.min(ref.from.row, ref.to.row);
            maxCol = Math.max(ref.from.col, ref.to.col);
            minCol = Math.min(ref.from.col, ref.to.col);
        }

        let err: FormulaError | undefined;
        for (const refItem of refs) {
            if (this.isFormulaError(refItem)) {
                err = refItem;
                break;
            }
            const currentRef = refItem.ref as CellRef | RangeRef;
            if (!currentRef) {
                throw Error(`Expecting a reference, but got ${refItem}.`);
            }
            if (!('from' in currentRef)) {
                if ((currentRef as CellRef).row === undefined || (currentRef as CellRef).col === undefined) {
                    throw Error('Cannot intersect the whole row or column.');
                }
                // cell ref
                const cellRef = currentRef as CellRef;
                if (cellRef.row > maxRow || cellRef.row < minRow || cellRef.col > maxCol || cellRef.col < minCol
                    || sheet !== cellRef.sheet) {
                    err = FormulaError.NULL;
                }
                maxRow = minRow = cellRef.row;
                maxCol = minCol = cellRef.col;
            } else {
                // range ref
                const rangeRef = currentRef as RangeRef;
                const refMaxRow = Math.max(rangeRef.from.row, rangeRef.to.row);
                const refMinRow = Math.min(rangeRef.from.row, rangeRef.to.row);
                const refMaxCol = Math.max(rangeRef.from.col, rangeRef.to.col);
                const refMinCol = Math.min(rangeRef.from.col, rangeRef.to.col);
                if (refMinRow > maxRow || refMaxRow < minRow || refMinCol > maxCol || refMaxCol < minCol
                    || sheet !== rangeRef.sheet) {
                    err = FormulaError.NULL;
                }
                // update
                maxRow = Math.min(maxRow, refMaxRow);
                minRow = Math.max(minRow, refMinRow);
                maxCol = Math.min(maxCol, refMaxCol);
                minCol = Math.max(minCol, refMinCol);
            }
        }
        
        if (err) return err;
        
        // check if the ref can be reduced to cell reference
        if (maxRow === minRow && maxCol === minCol) {
            res = {
                ref: {
                    sheet,
                    row: maxRow,
                    col: maxCol
                } as CellRef
            };
        } else {
            res = {
                ref: {
                    sheet,
                    from: { row: minRow, col: minCol },
                    to: { row: maxRow, col: maxCol }
                } as RangeRef
            };
        }

        if (!res.ref.sheet) {
            delete res.ref.sheet;
        }
        return res;
    }

    applyUnion(refs: any[]): Collection {
        const collection = new Collection();
        for (let i = 0; i < refs.length; i++) {
            if (this.isFormulaError(refs[i])) {
                return refs[i];
            }
            collection.add(this.extractRefValue(refs[i]).val, refs[i]);
        }
        return collection;
    }

    /**
     * Apply multiple references, e.g. A1:B3:C8:A:1:.....
     * @param refs - Array of references
     * @return Range reference object
     */
    applyRange(refs: any[]): RefObject {
        let res: RefObject;
        let maxRow = -1, maxCol = -1, minRow = MAX_ROW + 1, minCol = MAX_COLUMN + 1;
        
        for (let ref of refs) {
            if (this.isFormulaError(ref)) {
                continue;
            }
            // row ref is saved as number, parse the number to row ref here
            if (typeof ref === 'number') {
                ref = this.parseRow(ref);
            }
            const currentRef = ref.ref as CellRef;
            // check whole row/col reference
            if (currentRef.row === undefined) {
                minRow = 1;
                maxRow = MAX_ROW;
            }
            if (currentRef.col === undefined) {
                minCol = 1;
                maxCol = MAX_COLUMN;
            }

            if (currentRef.row > maxRow) {
                maxRow = currentRef.row;
            }
            if (currentRef.row < minRow) {
                minRow = currentRef.row;
            }
            if (currentRef.col > maxCol) {
                maxCol = currentRef.col;
            }
            if (currentRef.col < minCol) {
                minCol = currentRef.col;
            }
        }
        
        if (maxRow === minRow && maxCol === minCol) {
            res = {
                ref: {
                    row: maxRow,
                    col: maxCol
                } as CellRef
            };
        } else {
            res = {
                ref: {
                    from: { row: minRow, col: minCol },
                    to: { row: maxRow, col: maxCol }
                } as RangeRef
            };
        }
        return res;
    }

    /**
     * Throw away the refs, and retrieve the value.
     * @return Object with value and array flag
     */
    extractRefValue(obj: any): ExtractedValue {
        let res = obj;
        let isArray = false;
        if (Array.isArray(res)) {
            isArray = true;
        }
        if (obj && obj.ref) {
            // can be number or array
            return { val: this.context.retrieveRef(obj), isArray };
        }
        return { val: res, isArray };
    }

    /**
     * Convert array representation to actual array
     * @param array - Array to convert
     * @return Converted array
     */
    toArray(array: any[][]): any[][] {
        return array;
    }

    /**
     * Convert string to number
     * @param number - Number string
     * @return Converted number
     */
    toNumber(number: string): number {
        return Number(number);
    }

    /**
     * Convert quoted string to actual string
     * @param string - Quoted string
     * @return Unquoted string
     */
    toString(string: string): string {
        return string.substring(1, string.length - 1).replace(/""/g, '"');
    }

    /**
     * Convert boolean string to boolean value
     * @param bool - Boolean string
     * @return Boolean value
     */
    toBoolean(bool: string): boolean {
        return bool === 'TRUE';
    }

    /**
     * Parse an error string to FormulaError
     * @param error - Error string
     * @return FormulaError instance
     */
    toError(error: string): FormulaError {
        return new FormulaError(error.toUpperCase());
    }

    isFormulaError(obj: any): obj is FormulaError {
        return obj instanceof FormulaError;
    }

    static formatChevrotainError(error: any, inputText: string): FormulaError {
        let line: number, column: number;
        let msg = '';
        
        // e.g. SUM(1))
        if (error instanceof NotAllInputParsedException) {
            line = error.token.startLine || 1;
            column = error.token.startColumn || 1;
        } else {
            line = error.previousToken?.startLine || 1;
            column = (error.previousToken?.startColumn || 0) + 1;
        }

        msg += '\n' + inputText.split('\n')[line - 1] + '\n';
        msg += Array(column - 1).fill(' ').join('') + '^\n';
        msg += `Error at position ${line}:${column}\n` + error.message;
        error.errorLocation = { line, column };
        return FormulaError.ERROR(msg, error as ErrorDetails);
    }
}

export default Utils;
export const formatChevrotainError = Utils.formatChevrotainError;