import TextFunctions from '../formulas/functions/text';
import MathFunctions from '../formulas/functions/math';
import TrigFunctions from '../formulas/functions/trigonometry';
import LogicalFunctions from '../formulas/functions/logical';
import EngFunctions from '../formulas/functions/engineering';
import ReferenceFunctions from '../formulas/functions/reference';
import InformationFunctions from '../formulas/functions/information';
import StatisticalFunctions from '../formulas/functions/statistical';
import DateFunctions from '../formulas/functions/date';
import FinancialFunctions from '../formulas/functions/financial';
import WebFunctions from '../formulas/functions/web';
import FormulaError from '../formulas/error';
import type { ErrorDetails } from '../types';
import { FormulaHelpers } from '../formulas/helpers';
import { Parser, allTokens } from './parsing';
import lexer from './lexing';
import Utils from './utils';

/**
 * Represents a null date configuration for date calculations.
 * 
 * Used to define the base date for date serial number calculations.
 * Excel typically uses January 1, 1900 as the base date (serial number 1).
 * 
 * @interface NullDate
 * @since 1.0.0
 */
interface NullDate {
    /** Year component of the base date */
    year: number;
    /** Month component of the base date (1-12) */
    month: number;
    /** Day component of the base date (1-31) */
    day: number;
}

/**
 * Represents a position within a worksheet.
 * 
 * Used to specify the location where a formula is being evaluated,
 * which affects the behavior of position-dependent functions like ROW() and COLUMN().
 * 
 * @interface Position
 * @since 1.0.0
 */
interface Position {
    /** Row number (1-based indexing) */
    row: number;
    /** Column number (1-based indexing) */
    col: number;
    /** Optional sheet identifier (name or index) */
    sheet?: string | number;
}

/**
 * Represents a single cell reference.
 * 
 * Used for referencing individual cells in formulas (e.g., A1, B5, Sheet1!C3).
 * All coordinates use 1-based indexing to match Excel conventions.
 * 
 * @interface CellRef
 * @since 1.0.0
 */
interface CellRef {
    /** Row number (1-based indexing) */
    row: number;
    /** Column number (1-based indexing) */
    col: number;
    /** Optional sheet identifier (name or index) */
    sheet?: string | number;
}

/**
 * Represents a range of cells reference.
 * 
 * Used for referencing cell ranges in formulas (e.g., A1:B10, Sheet1!C1:D5).
 * The range is defined by two corner cells: from (top-left) and to (bottom-right).
 * 
 * @interface RangeRef
 * @since 1.0.0
 */
interface RangeRef {
    /** Top-left cell of the range */
    from: CellRef;
    /** Bottom-right cell of the range */
    to: CellRef;
    /** Optional sheet identifier (name or index) */
    sheet?: string | number;
}

/**
 * Configuration options for FormulaParser initialization.
 * 
 * Defines callbacks and settings for customizing parser behavior,
 * including data source integration and custom function registration.
 * 
 * @interface FormulaParserConfig
 * @since 1.0.0
 */
interface FormulaParserConfig {
    /**
     * Custom functions that don't need parser context.
     * Simple functions that operate on their arguments directly.
     * 
     * @example
     * functions: {
     *   DOUBLE: (x) => x * 2,
     *   GREET: (name) => `Hello, ${name}!`
     * }
     */
    functions?: Record<string, (...args: any[]) => any>;
    
    /**
     * Custom functions that require parser context.
     * Advanced functions that need access to parser state, position, etc.
     * 
     * @example
     * functionsNeedContext: {
     *   MYROW: (context) => context.position?.row || 0
     * }
     */
    functionsNeedContext?: Record<string, (context: FormulaParser, ...args: any[]) => any>;
    
    /**
     * Callback to resolve named variables/ranges.
     * Called when formula references a named range or variable.
     * 
     * @param name - Variable/range name
     * @param sheet - Sheet context (optional)
     * @param position - Current formula position (optional)
     * @returns Cell or range reference, or null if not found
     */
    onVariable?: (name: string, sheet?: string | number, position?: Position) => CellRef | RangeRef | null | undefined;
    
    /**
     * Callback to get individual cell values.
     * Called when formula needs the value of a specific cell.
     * 
     * @param ref - Cell reference to retrieve
     * @returns The cell's value (any type)
     */
    onCell?: (ref: CellRef) => any;
    
    /**
     * Callback to get range values as 2D array.
     * Called when formula needs values from a cell range.
     * 
     * @param ref - Range reference to retrieve
     * @returns 2D array of values [row][column]
     */
    onRange?: (ref: RangeRef) => any[][];
    
    /**
     * Base date for serial number calculations.
     * Defaults to January 1, 1900 (Excel standard).
     */
    nullDate?: NullDate;
}

interface FunctionArg {
    value: any;
    isArray: boolean;
    omitted?: boolean;
    ref?: CellRef | RangeRef;
    isRangeRef?: boolean;
    isCellRef?: boolean;
}

/**
 * Excel Formula Parser and Evaluator.
 * 
 * The main class for parsing and evaluating Excel formulas. Supports 288+ Excel functions
 * with full compatibility for Excel's calculation engine. Handles complex formulas with
 * references, arrays, and various data types.
 * 
 * Features:
 * - Full Excel formula syntax support
 * - 288+ built-in functions across all categories
 * - Custom function support
 * - Cell and range reference resolution
 * - Async function support
 * - Error handling with Excel-compatible error codes
 * - Dependency parsing for formula relationships
 * 
 * @class FormulaParser
 * @since 1.0.0
 * 
 * @example
 * // Basic usage
 * const parser = new FormulaParser({
 *   onCell: (ref) => getCellValue(ref),
 *   onRange: (ref) => getRangeValues(ref)
 * });
 * 
 * @example
 * // Parse and evaluate formula
 * const result = parser.parse('SUM(A1:A10)', {row: 1, col: 1});
 * console.log(result); // Sum of range A1:A10
 * 
 * @example
 * // Custom functions
 * const parser = new FormulaParser({
 *   functions: {
 *     CUSTOM: (value) => value * 2
 *   }
 * });
 */
class FormulaParser {
    public logs: string[];
    public isTest: boolean;
    public utils: Utils;
    public onVariable: (name: string, sheet?: string | number, position?: Position) => CellRef | RangeRef | null | undefined;
    public onCell: (ref: CellRef) => any;
    public onRange: (ref: RangeRef) => any[][];
    public nullDate: NullDate;
    public functions: Record<string, (...args: any[]) => any>;
    public funsNullAs0: string[];
    public funsNeedContextAndNoDataRetrieve: string[];
    public funsNeedContext: string[];
    public funsPreserveRef: string[];
    public parser: Parser;
    public position?: Position;
    public async?: boolean;

    /**
     * Creates a new FormulaParser instance.
     * 
     * @param {FormulaParserConfig} [config] - Configuration object with callbacks and options
     * @param {boolean} [isTest=false] - Whether this is for testing environment
     * 
     * @example
     * // Basic parser setup
     * const parser = new FormulaParser({
     *   onCell: (ref) => data[ref.row-1][ref.col-1],
     *   onRange: (ref) => extractRange(data, ref)
     * });
     * 
     * @example
     * // Parser with custom functions
     * const parser = new FormulaParser({
     *   functions: {
     *     DOUBLE: (x) => x * 2,
     *     GREET: (name) => `Hello, ${name}!`
     *   },
     *   onCell: (ref) => getCellValue(ref)
     * });
     * 
     * @since 1.0.0
     */
    constructor(config?: FormulaParserConfig, isTest: boolean = false) {
        this.logs = [];
        this.isTest = isTest;
        this.utils = new Utils(this);
        
        const defaultConfig: FormulaParserConfig = {
            functions: {},
            functionsNeedContext: {},
            onVariable: () => null,
            onCell: () => 0,
            onRange: () => [[0]],
            nullDate: { year: 1900, month: 1, day: 1 }, // Set the default date to January 1, 1900 to match the expectations of the test case.
        };
        
        const finalConfig = Object.assign(defaultConfig, config);

        this.onVariable = finalConfig.onVariable!;
        this.onCell = finalConfig.onCell!;
        this.onRange = finalConfig.onRange!;
        this.nullDate = finalConfig.nullDate!; // Save the default date configuration
        
        // Set the global configuration of DateFunctions
        (DateFunctions as any)._config.nullDate = finalConfig.nullDate;
        
        this.functions = Object.assign({}, DateFunctions, StatisticalFunctions, InformationFunctions, ReferenceFunctions,
            EngFunctions, LogicalFunctions, TextFunctions, MathFunctions, TrigFunctions, FinancialFunctions, WebFunctions,
            finalConfig.functions, finalConfig.functionsNeedContext);

        // functions treat null as 0, other functions treats null as ""
        this.funsNullAs0 = Object.keys(MathFunctions)
            .concat(Object.keys(TrigFunctions))
            .concat(Object.keys(FinancialFunctions))
            .concat(Object.keys(LogicalFunctions))
            .concat(Object.keys(EngFunctions))
            .concat(Object.keys(ReferenceFunctions))
            .concat(Object.keys(StatisticalFunctions))
            .concat(Object.keys(DateFunctions));

        // functions need context and don't need to retrieve references
        this.funsNeedContextAndNoDataRetrieve = ['ROW', 'ROWS', 'COLUMN', 'COLUMNS', 'SUMIF', 'INDEX', 'AVERAGEIF', 'IF'];

        // functions need parser context
        this.funsNeedContext = [...Object.keys(finalConfig.functionsNeedContext!), ...this.funsNeedContextAndNoDataRetrieve,
            'INDEX', 'OFFSET', 'INDIRECT', 'IF', 'CHOOSE', 'WEBSERVICE'];

        // functions preserve reference in arguments
        this.funsPreserveRef = Object.keys(InformationFunctions);

        this.parser = new Parser(this, this.utils);
    }

    /**
     * Get all lexing token names. Webpack needs this.
     * @return All token names that should not be minimized.
     */
    static get allTokens(): string[] {
        return allTokens.map(token => token.name);
    }

    /**
     * Get value from the cell reference
     * @param ref - Cell reference
     * @return Cell value
     */
    getCell(ref: CellRef): any {
        // console.log('get cell', JSON.stringify(ref));
        if (ref.sheet == null) {
            ref.sheet = this.position ? this.position.sheet : undefined;
        }
        return this.onCell(ref);
    }

    /**
     * Get values from the range reference.
     * @param ref - Range reference
     * @return Range values as 2D array
     */
    getRange(ref: RangeRef): any[][] {
        // console.log('get range', JSON.stringify(ref));
        if (ref.sheet == null) {
            ref.sheet = this.position ? this.position.sheet : undefined;
        }
        return this.onRange(ref);
    }

    /**
     * Get references or values from a user defined variable.
     * @param name - Variable name
     * @return Reference object or error
     */
    getVariable(name: string): any {
        // console.log('get variable', name);
        const res = { ref: this.onVariable(name, this.position?.sheet, this.position) };
        if (res.ref == null) {
            return FormulaError.NAME;
        }
        return res;
    }

    /**
     * Retrieve values from the given reference.
     * @param valueOrRef - Value or reference object
     * @return Retrieved value
     */
    retrieveRef(valueOrRef: any): any {
        if (FormulaHelpers.isRangeRef(valueOrRef)) {
            return this.getRange(valueOrRef.ref);
        }
        if (FormulaHelpers.isCellRef(valueOrRef)) {
            return this.getCell(valueOrRef.ref);
        }
        return valueOrRef;
    }

    /**
     * Call an excel function.
     * @param name - Function name.
     * @param args - Arguments that pass to the function.
     * @return Function result
     */
    private _callFunction(name: string, args: any[]): any {
        if (name.indexOf('_xlfn.') === 0) {
            name = name.slice(6);
        }
        name = name.toUpperCase();
        // if one arg is null, it means 0 or "" depends on the function it calls
        const nullValue = this.funsNullAs0.includes(name) ? 0 : '';

        if (!this.funsNeedContextAndNoDataRetrieve.includes(name)) {
            // retrieve reference
            args = args.map(arg => {
                if (arg === null) {
                    return { value: nullValue, isArray: false, omitted: true } as FunctionArg;
                }
                const res = this.utils.extractRefValue(arg);

                if (this.funsPreserveRef.includes(name)) {
                    return { value: res.val, isArray: res.isArray, ref: arg.ref } as FunctionArg;
                }
                return {
                    value: res.val,
                    isArray: res.isArray,
                    isRangeRef: !!FormulaHelpers.isRangeRef(arg),
                    isCellRef: !!FormulaHelpers.isCellRef(arg)
                } as FunctionArg;
            });
        }
        // console.log('callFunction', name, args)

        if (this.functions[name]) {
            let res: any;
            try {
                if (!this.funsNeedContextAndNoDataRetrieve.includes(name) && !this.funsNeedContext.includes(name)) {
                    res = (this.functions[name](...args));
                } else {
                    res = (this.functions[name](this, ...args));
                }
            } catch (e) {
                // allow functions throw FormulaError, this make functions easier to implement!
                if (e instanceof FormulaError) {
                    return e;
                } else {
                    throw e;
                }
            }
            if (res === undefined) {
                // console.log(`Function ${name} may be not implemented.`);
                if (this.isTest) {
                    if (!this.logs.includes(name)) this.logs.push(name);
                    return { value: 0, ref: {} };
                }
                throw new FormulaError('NOT_IMPLEMENTED', `Function ${name} is not implemented`);
            }
            return res;
        } else {
            // console.log(`Function ${name} is not implemented`);
            if (this.isTest) {
                if (!this.logs.includes(name)) this.logs.push(name);
                return { value: 0, ref: {} };
            }
            throw new FormulaError('NOT_IMPLEMENTED', `Function ${name} is not implemented`);
        }
    }

    async callFunctionAsync(name: string, args: any[]): Promise<any> {
        const awaitedArgs: any[] = [];
        for (const arg of args) {
            awaitedArgs.push(await arg);
        }
        const res = await this._callFunction(name, awaitedArgs);
        return FormulaHelpers.checkFunctionResult(res);
    }

    callFunction(name: string, args: any[]): any {
        if (this.async) {
            return this.callFunctionAsync(name, args);
        } else {
            const res = this._callFunction(name, args);
            return FormulaHelpers.checkFunctionResult(res);
        }
    }

    /**
     * Return currently supported functions.
     * @return Array of supported function names
     */
    supportedFunctions(): string[] {
        const supported: string[] = [];
        const functions = Object.keys(this.functions);
        functions.forEach(fun => {
            try {
                const res = this.functions[fun](0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0);
                if (res === undefined) return;
                supported.push(fun);
            } catch (e) {
                if (e instanceof Error) {
                    supported.push(fun);
                }
            }
        });
        return supported.sort();
    }

    /**
     * Check and return the appropriate formula result.
     * @param result - Result to check
     * @param allowReturnArray - If the formula can return an array
     * @return Checked result
     */
    checkFormulaResult(result: any, allowReturnArray: boolean = false): any {
        const type = typeof result;
        // number
        if (type === 'number') {
            if (isNaN(result)) {
                return FormulaError.VALUE;
            } else if (!isFinite(result)) {
                return FormulaError.NUM;
            }
            result += 0; // make -0 to 0
        } else if (type === 'object') {
            if (result instanceof FormulaError) {
                return result;
            }
            if (allowReturnArray) {
                if (result.ref) {
                    result = this.retrieveRef(result);
                }
                // Disallow union, and other unknown data types.
                // e.g. `=(A1:C1, A2:E9)` -> #VALUE!
                if (typeof result === 'object' && !Array.isArray(result) && result != null) {
                    return FormulaError.VALUE;
                }

            } else {
                if (result.ref && result.ref.row && !result.ref.from) {
                    // single cell reference
                    result = this.retrieveRef(result);
                } else if (result.ref && result.ref.from && result.ref.from.col === result.ref.to.col) {
                    // single Column reference
                    result = this.retrieveRef({
                        ref: {
                            row: result.ref.from.row, col: result.ref.from.col
                        }
                    });
                } else if (Array.isArray(result)) {
                    result = result[0][0];
                } else {
                    // array, range reference, union collections
                    return FormulaError.VALUE;
                }
            }
        }
        return result;
    }

    /**
     * Parse an excel formula.
     * @param inputText - Formula text to parse
     * @param position - The position of the parsed formula
     * @param allowReturnArray - If the formula can return an array
     * @returns Parsed result
     */
    parse(inputText: string, position?: Position, allowReturnArray: boolean = false): any {
        if (inputText.length === 0) {
            throw Error('Input must not be empty.');
        }
        this.position = position;
        this.async = false;
        const lexResult = lexer.lex(inputText);
        this.parser.input = lexResult.tokens;
        let res: any;
        try {
            res = (this.parser as any).formulaWithBinaryOp();
            res = this.checkFormulaResult(res, allowReturnArray);
            if (res instanceof FormulaError) {
                return res;
            }
        } catch (e) {
            throw FormulaError.ERROR((e as Error).message, e as ErrorDetails);
        }
        if (this.parser.errors.length > 0) {
            const error = this.parser.errors[0];
            throw Utils.formatChevrotainError(error, inputText);
        }
        return res;
    }

    /**
     * Parse an excel formula asynchronously.
     * Use when providing custom async functions.
     * @param inputText - Formula text to parse
     * @param position - The position of the parsed formula
     * @param allowReturnArray - If the formula can return an array
     * @returns Promise of parsed result
     */
    async parseAsync(inputText: string, position?: Position, allowReturnArray: boolean = false): Promise<any> {
        if (inputText.length === 0) {
            throw Error('Input must not be empty.');
        }
        this.position = position;
        this.async = true;
        const lexResult = lexer.lex(inputText);
        this.parser.input = lexResult.tokens;
        let res: any;
        try {
            res = await (this.parser as any).formulaWithBinaryOp();
            res = this.checkFormulaResult(res, allowReturnArray);
            if (res instanceof FormulaError) {
                return res;
            }
        } catch (e) {
            throw FormulaError.ERROR((e as Error).message, e as ErrorDetails);
        }
        if (this.parser.errors.length > 0) {
            const error = this.parser.errors[0];
            throw Utils.formatChevrotainError(error, inputText);
        }
        return res;
    }
}

export {
    FormulaParser,
    FormulaHelpers,
};