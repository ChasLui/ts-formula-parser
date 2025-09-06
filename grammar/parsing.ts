import lexer, { tokenVocabulary } from './lexing';
import { EmbeddedActionsParser, TokenType, IToken } from 'chevrotain';
import { FormulaParser } from './hooks';
import { DepParser } from './dependency/hooks';
import Utils from './utils';

const {
    String,
    SheetQuoted,
    ExcelRefFunction,
    ExcelConditionalRefFunction,
    Function,
    FormulaErrorT,
    RefError,
    Cell,
    Sheet,
    Name,
    Number,
    Boolean,
    Column,

    // At,
    Comma,
    Colon,
    Semicolon,
    OpenParen,
    CloseParen,
    // OpenSquareParen,
    // CloseSquareParen,
    // ExclamationMark,
    OpenCurlyParen,
    CloseCurlyParen,
    MulOp,
    PlusOp,
    DivOp,
    MinOp,
    ConcatOp,
    ExOp,
    PercentOp,
    NeqOp,
    GteOp,
    LteOp,
    GtOp,
    EqOp,
    LtOp
} = lexer.tokenVocabulary;

// Define interface for parser context (FormulaParser or DepParser)
interface ParserContext {
    callFunction(name: string, args: any[]): any;
    getVariable(name: string): any;
}

class Parsing extends EmbeddedActionsParser {
    private utils: Utils;
    private binaryOperatorsPrecedence: string[][];
    private c1?: { ALT: () => string }[]; // Cache for alternatives

    /**
     * Parser constructor
     * @param context - FormulaParser or DepParser instance
     * @param utils - Utils instance
     */
    constructor(context: ParserContext, utils: Utils) {
        super(tokenVocabulary, {
            maxLookahead: 1,
            skipValidations: true,
            // traceInitPerf: true,
        });
        this.utils = utils;
        this.binaryOperatorsPrecedence = [
            ['^'],
            ['*', '/'],
            ['+', '-'],
            ['&'],
            ['<', '>', '=', '<>', '<=', '>='],
        ];
        const $ = this;

        // Adopted from https://github.com/spreadsheetlab/XLParser/blob/master/src/XLParser/ExcelFormulaGrammar.cs

        ($ as any).RULE('formulaWithBinaryOp', () => {
            const infixes: string[] = [];
            const values = [$.SUBRULE(($ as any).formulaWithPercentOp)];
            $.MANY(() => {
                // Caching Arrays of Alternatives
                // https://sap.github.io/chevrotain/docs/guide/performance.html#caching-arrays-of-alternatives
                infixes.push($.OR($.c1 ||
                    (
                        $.c1 = [
                            {ALT: () => $.CONSUME(GtOp).image},
                            {ALT: () => $.CONSUME(EqOp).image},
                            {ALT: () => $.CONSUME(LtOp).image},
                            {ALT: () => $.CONSUME(NeqOp).image},
                            {ALT: () => $.CONSUME(GteOp).image},
                            {ALT: () => $.CONSUME(LteOp).image},
                            {ALT: () => $.CONSUME(ConcatOp).image},
                            {ALT: () => $.CONSUME(PlusOp).image},
                            {ALT: () => $.CONSUME(MinOp).image},
                            {ALT: () => $.CONSUME(MulOp).image},
                            {ALT: () => $.CONSUME(DivOp).image},
                            {ALT: () => $.CONSUME(ExOp).image}
                        ]
                    )));
                values.push($.SUBRULE2(($ as any).formulaWithPercentOp));
            });
            $.ACTION(() => {
                // evaluate
                for (const ops of this.binaryOperatorsPrecedence) {
                    for (let index = 0, length = infixes.length; index < length; index++) {
                        const infix = infixes[index];
                        if (!ops.includes(infix)) continue;
                        infixes.splice(index, 1);
                        values.splice(index, 2, this.utils.applyInfix(values[index], infix, values[index + 1]));
                        index--;
                        length--;
                    }
                }
            });

            return values[0];
        });

        ($ as any).RULE('plusMinusOp', () => $.OR([
            {ALT: () => $.CONSUME(PlusOp).image},
            {ALT: () => $.CONSUME(MinOp).image}
        ]));

        ($ as any).RULE('formulaWithPercentOp', () => {
            let value = $.SUBRULE(($ as any).formulaWithUnaryOp);
            $.OPTION(() => {
                const postfix = $.CONSUME(PercentOp).image;
                value = $.ACTION(() => this.utils.applyPostfix(value, postfix));
            });
            return value;
        });

        ($ as any).RULE('formulaWithUnaryOp', () => {
            // support ++---3 => -3
            const prefixes: string[] = [];
            $.MANY(() => {
                const op = $.OR([
                    {ALT: () => $.CONSUME(PlusOp).image},
                    {ALT: () => $.CONSUME(MinOp).image}
                ]);
                prefixes.push(op);
            });
            const formula = $.SUBRULE(($ as any).formulaWithIntersect);
            if (prefixes.length > 0) return $.ACTION(() => this.utils.applyPrefix(prefixes, formula));
            return formula;
        });

        ($ as any).RULE('formulaWithIntersect', () => {
            // e.g.  'A1 A2 A3'
            let ref1 = $.SUBRULE(($ as any).formulaWithRange);
            const refs = [ref1];
            // console.log('check intersect')
            $.MANY({
                GATE: () => {
                    // see https://github.com/SAP/chevrotain/blob/master/examples/grammars/css/css.js#L436-L441
                    const prevToken = $.LA(0);
                    const nextToken = $.LA(1);
                    //  This is the only place where the grammar is whitespace sensitive.
                    return nextToken.startOffset > prevToken.endOffset! + 1;
                },
                DEF: () => {
                    refs.push($.SUBRULE3(($ as any).formulaWithRange));
                }
            });
            if (refs.length > 1) {
                return $.ACTION(() => $.ACTION(() => this.utils.applyIntersect(refs)))
            }
            return ref1;
        });

        ($ as any).RULE('formulaWithRange', () => {
            // e.g. 'A1:C3' or 'A1:A3:C4', can be any number of references, at lease 2
            const ref1 = $.SUBRULE(($ as any).formula);
            const refs = [ref1];
            $.MANY(() => {
                $.CONSUME(Colon);
                refs.push($.SUBRULE2(($ as any).formula));
            });
            if (refs.length > 1)
                return $.ACTION(() => $.ACTION(() => this.utils.applyRange(refs)));
            return ref1;
        });

        ($ as any).RULE('formula', () => $.OR9([
            {ALT: () => $.SUBRULE(($ as any).referenceWithoutInfix)},
            {ALT: () => $.SUBRULE(($ as any).paren)},
            {ALT: () => $.SUBRULE(($ as any).constant)},
            {ALT: () => $.SUBRULE(($ as any).functionCall)},
            {ALT: () => $.SUBRULE(($ as any).constantArray)},
        ]));

        ($ as any).RULE('paren', () => {
            // formula paren or union paren
            $.CONSUME(OpenParen);
            let result: any;
            const refs: any[] = [];
            refs.push($.SUBRULE(($ as any).formulaWithBinaryOp));
            $.MANY(() => {
                $.CONSUME(Comma);
                refs.push($.SUBRULE2(($ as any).formulaWithBinaryOp));
            });
            if (refs.length > 1)
                result = $.ACTION(() => this.utils.applyUnion(refs));
            else
                result = refs[0];

            $.CONSUME(CloseParen);
            return result;
        });

        ($ as any).RULE('constantArray', () => {
            // console.log('constantArray');
            const arr: any[][] = [[]];
            let currentRow = 0;
            $.CONSUME(OpenCurlyParen);

            // array must contain at least one item
            arr[currentRow].push($.SUBRULE(($ as any).constantForArray));
            $.MANY(() => {
                const sep = $.OR([
                    {ALT: () => $.CONSUME(Comma).image},
                    {ALT: () => $.CONSUME(Semicolon).image}
                ]);
                const constant = $.SUBRULE2(($ as any).constantForArray);
                if (sep === ',') {
                    arr[currentRow].push(constant)
                } else {
                    currentRow++;
                    arr[currentRow] = [];
                    arr[currentRow].push(constant)
                }
            });

            $.CONSUME(CloseCurlyParen);

            return $.ACTION(() => this.utils.toArray(arr));
        });

        /**
         * Used in array
         */
        ($ as any).RULE('constantForArray', () => $.OR([
            {
                ALT: () => {
                    const prefix = $.OPTION(() => $.SUBRULE(($ as any).plusMinusOp));
                    const image = $.CONSUME(Number).image;
                    const number = $.ACTION(() => this.utils.toNumber(image));
                    if (prefix)
                        return $.ACTION(() => this.utils.applyPrefix([prefix as string], number));
                    return number;
                }
            }, {
                ALT: () => {
                    const str = $.CONSUME(String).image;
                    return $.ACTION(() => this.utils.toString(str));
                }
            }, {
                ALT: () => {
                    const bool = $.CONSUME(Boolean).image;
                    return $.ACTION(() => this.utils.toBoolean(bool));
                }
            }, {
                ALT: () => {
                    const err = $.CONSUME(FormulaErrorT).image;
                    return $.ACTION(() => this.utils.toError(err));
                }
            }, {
                ALT: () => {
                    const err = $.CONSUME(RefError).image;
                    return $.ACTION(() => this.utils.toError(err));
                }
            },
        ]));

        ($ as any).RULE('constant', () => $.OR([
            {
                ALT: () => {
                    const number = $.CONSUME(Number).image;
                    return $.ACTION(() => this.utils.toNumber(number));
                }
            }, {
                ALT: () => {
                    const str = $.CONSUME(String).image;
                    return $.ACTION(() => this.utils.toString(str));
                }
            }, {
                ALT: () => {
                    const bool = $.CONSUME(Boolean).image;
                    return $.ACTION(() => this.utils.toBoolean(bool));
                }
            }, {
                ALT: () => {
                    const err = $.CONSUME(FormulaErrorT).image;
                    return $.ACTION(() => this.utils.toError(err));
                }
            },
        ]));

        ($ as any).RULE('functionCall', () => {
            const functionName = $.CONSUME(Function).image.slice(0, -1);
            // console.log('functionName', functionName);
            const args = $.SUBRULE(($ as any).arguments);
            $.CONSUME(CloseParen);
            // dependency parser won't call function.
            return $.ACTION(() => context.callFunction(functionName, args as any[]));

        });

        ($ as any).RULE('arguments', () => {
            // console.log('try arguments')

            // allows ',' in the front
            $.MANY2(() => {
                $.CONSUME2(Comma);
            });
            const args: any[] = [];
            // allows empty arguments
            $.OPTION(() => {
                args.push($.SUBRULE(($ as any).formulaWithBinaryOp));
                $.MANY(() => {
                    $.CONSUME1(Comma);
                    args.push(null); // e.g. ROUND(1.5,)
                    $.OPTION3(() => {
                        args.pop();
                        args.push($.SUBRULE2(($ as any).formulaWithBinaryOp))
                    });
                });
            });
            return args;
        });

        ($ as any).RULE('referenceWithoutInfix', () => $.OR([

            {ALT: () => $.SUBRULE(($ as any).referenceItem)},

            {
                // sheet name prefix
                ALT: () => {
                    // console.log('try sheetName');
                    const sheetName = $.SUBRULE(($ as any).prefixName);
                    // console.log('sheetName', sheetName);
                    const referenceItem = $.SUBRULE2(($ as any).formulaWithRange);

                    return $.ACTION(() => {
                        if (this.utils.isFormulaError(referenceItem))
                            return referenceItem;
                        if (referenceItem && (referenceItem as any).ref) {
                            (referenceItem as any).ref.sheet = sheetName;
                        }
                        return referenceItem;
                    });
                }
            },

            // {ALT: () => $.SUBRULE('dynamicDataExchange')},
        ]));

        ($ as any).RULE('referenceItem', () => $.OR([
            {
                ALT: () => {
                    const address = $.CONSUME(Cell).image;
                    return $.ACTION(() => this.utils.parseCellAddress(address));
                }
            },
            {
                ALT: () => {
                    const name = $.CONSUME(Name).image;
                    return $.ACTION(() => context.getVariable(name))
                }
            },
            {
                ALT: () => {
                    const column = $.CONSUME(Column).image;
                    return $.ACTION(() => this.utils.parseCol(column))
                }
            },
            // A row check should be here, but the token is same with Number,
            // In other to resolve ambiguities, I leave this empty, and
            // parse the number to row number when needed.
            {
                ALT: () => {
                    const err = $.CONSUME(RefError).image;
                    return $.ACTION(() => this.utils.toError(err))
                }
            },
            // {ALT: () => $.SUBRULE($.udfFunctionCall)},
            // {ALT: () => $.SUBRULE($.structuredReference)},
        ]));

        ($ as any).RULE('prefixName', () => $.OR([
            {ALT: () => $.CONSUME(Sheet).image.slice(0, -1)},
            {ALT: () => $.CONSUME(SheetQuoted).image.slice(1, -2).replace(/''/g, "'")},
        ]));

        this.performSelfAnalysis();
    }
}

export {
    Parsing as Parser,
};
export const allTokens: TokenType[] = [...Object.values(tokenVocabulary)];