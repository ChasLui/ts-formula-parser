import FormulaError from '../../formulas/error';
import type { ErrorDetails } from '../../types';
import { FormulaHelpers } from '../../formulas/helpers';
import { Parser } from '../parsing';
import lexer from '../lexing';
import Utils from './utils';
import { formatChevrotainError } from '../utils';

// Define interfaces for position and configuration
interface Position {
    row: number;
    col: number;
    sheet?: string | number;
}

interface CellRef {
    row: number;
    col: number;
    sheet?: string | number;
}

interface RangeRef {
    from: CellRef;
    to: CellRef;
    sheet?: string | number;
}

interface RefObject {
    ref: CellRef | RangeRef;
}

interface DepParserConfig {
    onVariable?: (name: string, sheet?: string | number, position?: Position) => CellRef | RangeRef | null | undefined;
}

class DepParser {
    public data: (CellRef | RangeRef)[];
    public functions: Record<string, any>;
    private utils: Utils;
    private onVariable: (name: string, sheet?: string | number, position?: Position) => CellRef | RangeRef | null | undefined;
    private parser: Parser;
    private position?: Position;

    /**
     * Dependency parser constructor
     * @param config - Configuration object with onVariable function
     */
    constructor(config?: DepParserConfig) {
        this.data = [];
        this.utils = new Utils(this);
        
        const defaultConfig: DepParserConfig = {
            onVariable: () => null,
        };
        
        const finalConfig = Object.assign(defaultConfig, config);
        this.onVariable = finalConfig.onVariable!;
        this.functions = {};

        this.parser = new Parser(this as any, this.utils as any);
    }

    /**
     * Get value from the cell reference
     * @param ref - Cell reference
     * @return Always returns 0 for dependency parser
     */
    getCell(ref: CellRef): number {
        if (ref.row != null) {
            if (ref.sheet == null) {
                ref.sheet = this.position ? this.position.sheet : undefined;
            }
            const idx = this.data.findIndex(element => {
                if ('from' in element) {
                    // Range reference
                    const rangeElement = element as RangeRef;
                    return rangeElement.from.row <= ref.row && rangeElement.to.row >= ref.row
                        && rangeElement.from.col <= ref.col && rangeElement.to.col >= ref.col;
                } else {
                    // Cell reference
                    const cellElement = element as CellRef;
                    return cellElement.row === ref.row && cellElement.col === ref.col && cellElement.sheet === ref.sheet;
                }
            });
            if (idx === -1) {
                this.data.push(ref);
            }
        }
        return 0;
    }

    /**
     * Get values from the range reference.
     * @param ref - Range reference
     * @return Always returns [[0]] for dependency parser
     */
    getRange(ref: RangeRef): number[][] {
        if (ref.from.row != null) {
            if (ref.sheet == null) {
                ref.sheet = this.position ? this.position.sheet : undefined;
            }

            const idx = this.data.findIndex(element => {
                if ('from' in element) {
                    const rangeElement = element as RangeRef;
                    return rangeElement.from.row === ref.from.row && rangeElement.from.col === ref.from.col
                        && rangeElement.to.row === ref.to.row && rangeElement.to.col === ref.to.col;
                }
                return false;
            });
            if (idx === -1) {
                this.data.push(ref);
            }
        }
        return [[0]];
    }

    /**
     * Get references or values from a user defined variable.
     * @param name - Variable name
     * @return Always returns 0 for dependency parser
     */
    getVariable(name: string): number | FormulaError {
        const res: RefObject = { ref: this.onVariable(name, this.position?.sheet) as CellRef | RangeRef };
        if (res.ref == null) {
            return FormulaError.NAME;
        }
        if (FormulaHelpers.isCellRef(res)) {
            this.getCell(res.ref as CellRef);
        } else {
            this.getRange(res.ref as RangeRef);
        }
        return 0;
    }

    /**
     * Retrieve values from the given reference.
     * @param valueOrRef - Value or reference object
     * @return Retrieved value or reference
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
     * @return Always returns {value: 0, ref: {}} for dependency parser
     */
    callFunction(name: string, args: any[]): { value: number; ref: {} } {
        args.forEach(arg => {
            if (arg == null) {
                return;
            }
            this.retrieveRef(arg);
        });
        return { value: 0, ref: {} };
    }

    /**
     * Check and return the appropriate formula result.
     * @param result - Result to check
     * @return Undefined (dependency parser doesn't return results)
     */
    checkFormulaResult(result: any): void {
        this.retrieveRef(result);
    }

    /**
     * Parse an excel formula and return the dependencies
     * @param inputText - Formula text to parse
     * @param position - Position context for parsing
     * @param ignoreError - If true, throw FormulaError when error occurred. If false, return partial dependencies.
     * @returns Array of dependencies
     */
    parse(inputText: string, position: Position, ignoreError: boolean = false): (CellRef | RangeRef)[] {
        if (inputText.length === 0) {
            throw Error('Input must not be empty.');
        }
        this.data = [];
        this.position = position;
        
        const lexResult = lexer.lex(inputText);
        this.parser.input = lexResult.tokens;
        
        try {
            const res = (this.parser as any).formulaWithBinaryOp();
            this.checkFormulaResult(res);
        } catch (e) {
            if (!ignoreError) {
                throw FormulaError.ERROR((e as Error).message, e as ErrorDetails);
            }
        }
        
        if (this.parser.errors.length > 0 && !ignoreError) {
            const error = this.parser.errors[0];
            throw formatChevrotainError(error, inputText);
        }

        return this.data;
    }
}

export { DepParser };