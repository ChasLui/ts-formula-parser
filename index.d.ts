/**
 * TypeScript definitions for ts-formula-parser
 * A fast and reliable Excel formula parser with full ESM support
 */

// ======================
// Core Types
// ======================

/** Date configuration for null date handling */
export interface NullDate {
  year: number;
  month: number; 
  day: number;
}

/** Cell reference (1-based indexing) */
export interface CellRef {
  row: number;
  col: number;
  sheet?: string | number;
}

/** Range reference (1-based indexing) */
export interface RangeRef {
  from: CellRef;
  to: CellRef;
  sheet?: string | number;
}

/** Union reference - combination of multiple references */
export interface UnionRef {
  refs: (CellRef | RangeRef)[];
}

/** Parser position context */
export interface ParserPosition {
  row: number;
  col: number;
  sheet?: string | number;
}

// ======================
// Formula Types
// ======================

/** Valid function argument types */
export type FunctionArg =
  | null
  | undefined
  | number
  | string
  | boolean
  | CellRef
  | RangeRef
  | UnionRef
  | Array<Array<number | string | boolean | null>>;

/** Valid function result types */
export type FunctionResult =
  | number
  | string
  | boolean
  | CellRef
  | RangeRef
  | UnionRef
  | Array<Array<number | string | boolean | null>>;

// ======================
// Formula Error Class
// ======================

/** Formula error class */
declare class FormulaErrorClass extends Error {
  readonly _error: string;
  readonly details?: any;
  
  constructor(error: string, msg?: string, details?: any);
  
  get error(): string;
  get name(): string;
  
  toString(): string;
  
  // Static error instances
  static readonly NA: FormulaErrorClass;
  static readonly NAME: FormulaErrorClass;
  static readonly NULL: FormulaErrorClass;
  static readonly NUM: FormulaErrorClass;
  static readonly REF: FormulaErrorClass;
  static readonly VALUE: FormulaErrorClass;
  static readonly DIV0: FormulaErrorClass;
  
  // Static error factory methods
  static ARG_MISSING(args: string[]): FormulaErrorClass;
  static ERROR(msg: string, details?: any): FormulaErrorClass;
  static NOT_IMPLEMENTED(functionName: string): FormulaErrorClass;
  
  // Error map
  static readonly errorMap: Map<string, FormulaErrorClass>;
}

// ======================
// Formula Helpers Types
// ======================

/** Helper types enumeration */
declare const TypesEnum: {
  readonly NUMBER: 0;
  readonly ARRAY: 1;
  readonly BOOLEAN: 2;
  readonly STRING: 3;
  readonly RANGE_REF: 4;
  readonly CELL_REF: 5;
  readonly COLLECTIONS: 6;
  readonly NUMBER_NO_BOOLEAN: 10;
};

export type TypeValue = typeof TypesEnum[keyof typeof TypesEnum];

/** Formula helpers utility class */
declare class FormulaHelpersClass {
  readonly Types: typeof TypesEnum;
  
  constructor();
  
  // Type checking and conversion
  accept(arg: any, type: TypeValue | TypeValue[]): any;
  checkFunctionResult(result: any): any;
  
  // Array utilities
  flattenParams(params: any[]): any[];
  
  // Type detection
  getType(value: any): TypeValue;
  isNumber(value: any): boolean;
  isBoolean(value: any): boolean;
  isString(value: any): string;
  isArray(value: any): boolean;
  isCellRef(value: any): boolean;
  isRangeRef(value: any): boolean;
  
  // Value conversion
  toNumber(value: any): number;
  toString(value: any): string;
  toBoolean(value: any): boolean;
  
  // Array operations
  transpose(array: any[][]): any[][];
  
  // Static utilities
  static accept(arg: any, type: TypeValue | TypeValue[]): any;
  static flattenParams(params: any[]): any[];
}

// ======================
// Parser Configuration
// ======================

export interface FormulaParserConfig {
  /** Custom function definitions */
  functions?: Record<string, (...args: any[]) => any>;
  
  /** Custom functions that need parser context */
  functionsNeedContext?: Record<string, (context: ParserContext, ...args: any[]) => any>;
  
  /** Variable resolver for defined names */
  onVariable?: (name: string, sheet?: string | number, position?: ParserPosition) => CellRef | RangeRef | null | undefined;
  
  /** Cell value provider */
  onCell?: (ref: CellRef) => any;
  
  /** Range value provider */
  onRange?: (ref: RangeRef) => any[][];
  
  /** Null date configuration */
  nullDate?: NullDate;
}

/** Parser context provided to context-aware functions */
export interface ParserContext {
  position: ParserPosition;
  parser: FormulaParserClass;
}

// ======================
// Main Parser Class
// ======================

declare class FormulaParserClass {
  // Static constants
  static readonly MAX_ROW: number;
  static readonly MAX_COLUMN: number;
  static readonly SSF: any;
  static readonly DepParser: typeof DepParserClass;
  static readonly FormulaError: typeof FormulaErrorClass;
  static readonly FormulaHelpers: typeof FormulaHelpersClass;
  
  // Instance properties
  readonly config: FormulaParserConfig;
  readonly logs: string[];
  
  constructor(config?: FormulaParserConfig, isTest?: boolean);
  
  /** Get all available tokens */
  static get allTokens(): string[];
  
  /** Parse formula synchronously */
  parse(
    inputText: string,
    position?: ParserPosition,
    allowReturnArray?: boolean
  ): FunctionResult | any;
  
  /** Parse formula asynchronously */
  parseAsync(
    inputText: string,
    position?: ParserPosition,
    allowReturnArray?: boolean
  ): Promise<FunctionResult | any>;
  
  /** Call a function by name */
  callFunction(name: string, args: FunctionArg[]): FunctionResult | any;
  
  /** Call a function asynchronously by name */
  callFunctionAsync(
    name: string,
    args: FunctionArg[]
  ): Promise<FunctionResult | any>;
  
  /** Get list of supported function names */
  supportedFunctions(): string[];
  
  /** Check if a function is supported */
  isFunctionSupported(name: string): boolean;
}

// ======================
// Dependency Parser
// ======================

export interface DepParserConfig {
  /** Variable resolver for dependency analysis */
  onVariable?: (name: string, sheet?: string | number, position?: ParserPosition) => CellRef | RangeRef | null | undefined;
}

declare class DepParserClass {
  readonly data: any[];
  readonly functions: Record<string, any>;
  
  constructor(config?: DepParserConfig);
  
  /** Parse formula and extract dependencies */
  parse(
    inputText: string,
    position?: ParserPosition
  ): (CellRef | RangeRef)[];
}

// ======================
// SSF (Spreadsheet Formatting)
// ======================

declare const SSFObject: {
  format(fmt: string, val: number | Date, opts?: any): string;
  parse_date_code(v: number, opts?: any, b2?: any): Date;
  // Additional SSF methods would be defined here
  [key: string]: any;
};

// ======================
// Constants
// ======================

declare const MAX_ROW_CONST: number;
declare const MAX_COLUMN_CONST: number;

// ======================
// Default Export
// ======================

declare const DefaultFormulaParser: typeof FormulaParserClass;
export default DefaultFormulaParser;

// ======================
// Named Exports
// ======================

export {
  FormulaParserClass as FormulaParser,
  DepParserClass as DepParser,
  FormulaErrorClass as FormulaError,
  FormulaHelpersClass as FormulaHelpers,
  TypesEnum as Types,
  SSFObject as SSF,
  MAX_ROW_CONST as MAX_ROW,
  MAX_COLUMN_CONST as MAX_COLUMN
};