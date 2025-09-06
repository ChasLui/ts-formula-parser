import { describe } from 'vitest';
import { FormulaParser } from '../../../grammar/hooks';
import TestCase from './testcase';
import { generateTests } from '../../utils';
import type { CellRef } from '../../../index';

const data: any[][] = [
    ['', 1,2,3,4],
    ['string', 3,4,5,6],
    [null, undefined]

];

const parser = new FormulaParser({
    onCell: (ref: CellRef) => {
        return data[ref.row - 1]?.[ref.col - 1];
    }
});

describe('Information Functions', function () {
    generateTests(parser, TestCase);
});