import {FormulaParser} from '../../../grammar/hooks.js';
import TestCase from './testcase.js';
import {generateTests} from '../../utils.js';
const data = [
    ['', 1,2,3,4],
    ['string', 3,4,5,6],

];
const parser = new FormulaParser({
    onCell: ref => {
        return data[ref.row - 1][ref.col - 1];
    }
});

describe('Trigonometry Functions', function () {
    generateTests(parser, TestCase);
});
