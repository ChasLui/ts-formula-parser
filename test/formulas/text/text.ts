import { describe } from 'vitest';
import { FormulaParser } from '../../../grammar/hooks';
import TestCase from './testcase';
import { generateTests } from '../../utils';

const parser = new FormulaParser();

describe('Text Functions', function () {
    generateTests(parser, TestCase);
});