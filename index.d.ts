export type NullDate = { year: number; month: number; day: number };

export type CellRef = { row: number; col: number; sheet?: string | number };
export type RangeRef = { from: CellRef; to: CellRef };

export type Ref = { ref: CellRef } | { ref: RangeRef };

export type ParserPosition = {
  row: number;
  col: number;
  sheet?: string | number;
};

export type FunctionArg =
  | null
  | number
  | string
  | boolean
  | Ref
  | Array<Array<number | string | boolean | Ref>>;

export type FunctionResult =
  | number
  | string
  | boolean
  | Ref
  | Array<Array<number | string | boolean | Ref>>;

export interface FormulaParserConfig {
  functions?: Record<string, (...args: any[]) => any>;
  functionsNeedContext?: Record<
    string,
    (parser: FormulaParser, ...args: any[]) => any
  >;
  onVariable?: (
    name: string,
    sheet?: string | number,
    position?: ParserPosition
  ) => Ref | null | undefined;
  onCell?: (ref: CellRef) => any;
  onRange?: (ref: RangeRef) => any;
  nullDate?: NullDate;
}

declare class FormulaParser {
  static MAX_ROW: number;
  static MAX_COLUMN: number;
  static SSF: any;
  static DepParser: any;
  static FormulaError: any;

  constructor(config?: FormulaParserConfig, isTest?: boolean);

  static get allTokens(): string[];

  parse(
    inputText: string,
    position?: ParserPosition,
    allowReturnArray?: boolean
  ): FunctionResult | any;
  parseAsync(
    inputText: string,
    position?: ParserPosition,
    allowReturnArray?: boolean
  ): Promise<FunctionResult | any>;

  callFunction(name: string, args: FunctionArg[]): FunctionResult | any;
  callFunctionAsync(
    name: string,
    args: FunctionArg[]
  ): Promise<FunctionResult | any>;

  supportedFunctions(): string[];
}

export = FormulaParser;
