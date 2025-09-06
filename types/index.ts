// Global type definitions

/**
 * Cell reference type
 */
export interface CellReference {
    row: number;
    col: number;
    sheet?: string;
}

/**
 * Range reference type
 */
export interface RangeReference {
    from: CellReference;
    to: CellReference;
    sheet?: string;
}

/**
 * Date configuration type
 */
export interface DateConfig {
    year: number;
    month: number;
    day: number;
}

/**
 * Formula parser position type
 */
export interface FormulaPosition {
    row: number;
    col: number;
    sheet?: string;
}

/**
 * Formula function parameter type
 */
export interface FunctionArg {
    value: unknown;
    isArray?: boolean;
    isCellRef?: boolean;
    isRangeRef?: boolean;
    omitted?: boolean;
    ref?: CellReference | RangeReference;
}

/**
 * Generic value type
 */
export type FormulaValue = string | number | boolean | Date | null | undefined;

/**
 * Array value type
 */
export type ArrayValue = FormulaValue[][];

/**
 * Parser context type
 */
export interface ParserContext {
    position?: FormulaPosition;
    async?: boolean;
    utils: {
        extractRefValue(arg: unknown): { val: unknown; isArray: boolean };
    };
    getCell(ref: CellReference): FormulaValue;
    getRange(ref: RangeReference): ArrayValue;
    getVariable(name: string): unknown;
    retrieveRef(valueOrRef: unknown): FormulaValue | ArrayValue;
    callFunction(name: string, args: unknown[]): unknown;
    checkFormulaResult(result: unknown, allowReturnArray?: boolean): FormulaValue | ArrayValue;
}

/**
 * Formula function type
 */
export type FormulaFunction = (...args: FunctionArg[]) => FormulaValue | ArrayValue;

/**
 * Formula function type that requires context
 */
export type ContextualFormulaFunction = (context: ParserContext, ...args: FunctionArg[]) => FormulaValue | ArrayValue;

/**
 * Formula configuration type
 */
export interface FormulaConfig {
    functions?: Record<string, FormulaFunction>;
    functionsNeedContext?: Record<string, ContextualFormulaFunction>;
    onVariable?: (name: string, sheet?: string, position?: FormulaPosition) => unknown;
    onCell?: (ref: CellReference) => FormulaValue;
    onRange?: (ref: RangeReference) => ArrayValue;
    nullDate?: DateConfig;
}

/**
 * Error details type
 */
export type ErrorDetails = Error | object | string | null | undefined;

/**
 * Lexer Token type
 */
export interface LexerToken {
    image: string;
    startOffset: number;
    endOffset: number;
    tokenType: {
        name: string;
        PATTERN?: RegExp;
    };
}

/**
 * Lexical analysis result type
 */
export interface LexResult {
    tokens: LexerToken[];
    errors: unknown[];
}

/**
 * Parse error type
 */
export interface ParseError {
    message: string;
    token: LexerToken;
    resyncedTokens: LexerToken[];
    context?: {
        ruleStack: string[];
        ruleOccurrenceStack: number[];
    };
}