import { expect, describe, it } from 'vitest';
import FormulaError from '../../formulas/error';
import { FormulaParser } from '../../grammar/hooks';
import { DepParser } from '../../grammar/dependency/hooks';
import { MAX_ROW, MAX_COLUMN } from '../../index';
import type { CellRef, RangeRef, ParserPosition } from '../../index.d';

const parser = new FormulaParser({
    onCell: (ref: CellRef): number | null => {
        if (ref.row === 5 && ref.col === 5)
            return null
        return 1;
    },
    onRange: (ref: RangeRef): number[][] => {
        if (ref.to.row === MAX_ROW) {
            return [[1, 2, 3]];
        } else if (ref.to.col === MAX_COLUMN) {
            return [[1], [0]]
        }
        return [[1, 2, 3], [0, 0, 0]]
    },
    functions: {
        BAD_FN: (): never => {
            throw new SyntaxError();
        }
    }
});

const depParser = new DepParser({
    onVariable: (variable: string): CellRef | RangeRef => {
        return 'aaaa' === variable ? {from: {row: 1, col: 1}, to: {row: 2, col: 2}} : {row: 1, col: 1};
    }
});

const parsers: Array<FormulaParser | DepParser> = [parser, depParser];
const names: string[] = ['', ' (DepParser)']

const position: ParserPosition = {row: 1, col: 1, sheet: 'Sheet1'};

describe('#ERROR! Error handling', () => {
    parsers.forEach((parser, idx) => {
        it('should handle NotAllInputParsedException' + names[idx], function () {
            try {
                parser.parse('SUM(1))', position);
            } catch (e: any) {
                expect(e).to.be.instanceof(FormulaError);
                expect(e.details.errorLocation.line).to.eq(1);
                expect(e.details.errorLocation.column).to.eq(7);
                expect(e.name).to.eq('#ERROR!');
                expect(e.details.name).to.eq('NotAllInputParsedException');
                return;
            }
            throw Error('Should not reach here.');
        });

        it('should handle lexing error' + names[idx], function () {
            try {
                parser.parse('SUM(1)$', position);
            } catch (e: any) {
                expect(e).to.be.instanceof(FormulaError);
                expect(e.details.errorLocation.line).to.eq(1);
                expect(e.details.errorLocation.column).to.eq(7);
                expect(e.name).to.eq('#ERROR!');
                return;
            }
            throw Error('Should not reach here.');

        });

        it('should handle Parser error []' + names[idx], function () {
            try {
                parser.parse('SUM([Sales.xlsx]Jan!B2:B5)', position);
            } catch (e: any) {
                expect(e).to.be.instanceof(FormulaError);
                expect(e.name).to.eq('#ERROR!');
                return;
            }
            throw Error('Should not reach here.');
        });

        it('should handle Parser error' + names[idx], function () {
            try {
                parser.parse('SUM(B2:B5, "123"+)', position);
            } catch (e: any) {
                expect(e).to.be.instanceof(FormulaError);
                expect(e.name).to.eq('#ERROR!');
                return;
            }
            throw Error('Should not reach here.');

        });
    });

    it('should handle error from functions', function () {
        try {
            parser.parse('BAD_FN()', position);
        } catch (e: any) {
            expect(e).to.be.instanceof(FormulaError);
            expect(e.name).to.eq('#ERROR!');
            expect(e.details.name).to.eq('SyntaxError');
            return;
        }
        throw Error('Should not reach here.');

    });

    it('should handle errors in async', async function () {
        try {
            await parser.parseAsync('SUM(*()', position, true);
        } catch (e: any) {
            expect(e).to.be.instanceof(FormulaError);
            expect(e.name).to.eq('#ERROR!');
            return;
        }
        throw Error('Should not reach here.');
    });

    it('should not throw error when ignoreError = true (DepParser)', function () {
        try {
            depParser.parse('SUM(*()', position, true);
        } catch (e: any) {
            throw Error('Should not reach here.');
        }
    });
});