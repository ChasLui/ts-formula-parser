import {FormulaParser} from '../../../grammar/hooks.js';
import TestCase from './testcase.js';
import {generateTests} from '../../utils.js';

const parser = new FormulaParser();

describe('Text Functions', function () {
    generateTests(parser, TestCase);
});
