import { expect, it } from 'vitest';
import type { FormulaParser } from '../grammar/hooks';
import type { FormulaError } from '../formulas/error';

export interface TestCase {
    [functionName: string]: {
        [formula: string]: any;
    };
}

export const TestUtils = {
    generateTests: (parser: FormulaParser, TestCase: TestCase): void => {
        const funs = Object.keys(TestCase);

        funs.forEach(fun => {
            it(fun, () => {
                const formulas = Object.keys(TestCase[fun]);
                formulas.forEach(formula => {
                    const expected = TestCase[fun][formula];
                    let result = parser.parse(formula, {row: 1, col: 1});
                    if (result.result) result = result.result;
                    if (typeof result === "number" && typeof expected === "number") {
                        expect(result, `${formula} should equal ${expected}\n`).to.closeTo(expected, 0.00000001);
                    } else {
                        // For FormulaError
                        if (expected && typeof expected === 'object' && 'equals' in expected && typeof expected.equals === 'function') {
                            expect((expected as FormulaError).equals(result), `${formula} should equal ${(expected as FormulaError).error}\n`).to.equal(true);
                        } else {
                            expect(result, `${formula} should equal ${expected}\n`).to.equal(expected);
                        }
                    }
                })
            });
        });
    }
};

export const generateTests = TestUtils.generateTests;