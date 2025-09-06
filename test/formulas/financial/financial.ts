import { describe } from 'vitest';
import { FormulaParser } from '../../../grammar/hooks';
import TestCase from './testcase';
import { generateTests } from '../../utils';
import type { CellRef, RangeRef } from '../../../index';

// Provide test data for complex test scenarios
const data: any[][] = [
    ['Date', 'Rate', 'Value', 'Frequency'], // Header row
    [41640, 0.1, 1000, 1],  // 2014-01-01
    [41731, 0.05, 500, 2],  // 2014-04-01  
    [41823, 0.08, 2000, 4], // 2014-07-01
    [41914, 0.12, 1500, 1], // 2014-10-01
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
    }
});

describe('Financial Functions', function () {
    generateTests(parser, TestCase);
});