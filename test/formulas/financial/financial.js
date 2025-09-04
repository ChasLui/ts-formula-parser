import { describe } from 'vitest';
import {FormulaParser} from '../../../grammar/hooks.js';
import TestCase from './testcase.js';
import {generateTests} from '../../utils.js';

// 提供一些测试数据，用于复杂的测试场景
const data = [
    ['Date', 'Rate', 'Value', 'Frequency'], // Header row
    [41640, 0.1, 1000, 1],  // 2014-01-01
    [41731, 0.05, 500, 2],  // 2014-04-01  
    [41823, 0.08, 2000, 4], // 2014-07-01
    [41914, 0.12, 1500, 1], // 2014-10-01
];

const parser = new FormulaParser({
    onCell: ref => {
        return data[ref.row - 1][ref.col - 1];
    },
    onRange: ref => {
        const arr = [];
        for (let row = ref.from.row - 1; row < ref.to.row; row++) {
            const innerArr = [];
            for (let col = ref.from.col - 1; col < ref.to.col; col++) {
                innerArr.push(data[row][col])
            }
            arr.push(innerArr);
        }
        return arr;
    }
});

describe('Financial Functions', function () {
    generateTests(parser, TestCase);
});