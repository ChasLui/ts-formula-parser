import { describe } from 'vitest';
import { FormulaParser } from '../../grammar/hooks';
import TestCase from './testcase';
import { generateTests } from '../utils';
import FormulaError from '../../formulas/error';
import type { CellRef, RangeRef } from '../../index';

const data: any[][] = [
    [1, 2, 3, 4, 5],
    [100000, 7000, 250000, 5, 6],
    [200000, 14000, 4, 5, 6],
    [300000, 21000, 4, 5, 6],
    [400000, 28000, 4, 5, 6],
    ['string', 3, 4, 5, 6],
    // for SUMIF ex2
    ['Vegetables', 'Tomatoes', 2300, 5, 6], // row 7
    ['Vegetables', 'Celery', 5500, 5, 6], // row 8
    ['Fruits', 'Oranges', 800, 5, 6], // row 9
    ['', 'Butter', 400, 5, 6], // row 10
    ['Vegetables', 'Carrots', 4200, 5, 6], // row 11
    ['Fruits', 'Apples', 1200, 5, 6], // row 12
    [undefined, true, false, FormulaError.DIV0, 0] // row 13
];

const parser = new FormulaParser({
    onCell: (ref: CellRef) => {
        return data[ref.row - 1]?.[ref.col - 1];
    },
    onRange: (ref: RangeRef) => {
        const arr: any[][] = [];
        for (let row = ref.from.row - 1; row < ref.to.row; row++) {
            const innerArr: any[] = [];
            for (let col = ref.from.col - 1; col < ref.to.col; col++) {
                innerArr.push(data[row]?.[col]);
            }
            arr.push(innerArr);
        }
        return arr;
    },
    onVariable: (name: string) => {
        if (name === 'hello')
            return {row: 2, col: 2};
        else if (name === '_xlnm.print_titles')
            return {row: 7, col: 1};
    },
});

describe('Operators', function () {
    generateTests(parser, TestCase);
});