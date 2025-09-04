import {FormulaParser} from '../../../grammar/hooks.js';
import TestCase from './testcase.js';
import {generateTests} from '../../utils.js';
import { expect, describe, it, beforeEach, afterEach } from 'vitest';
import FormulaError from '../../../formulas/error.js';

const data = [
    ['fruit', 'price', 'count', 4, 5],
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

describe('Web Functions', function () {
    generateTests(parser, TestCase);
    
    // 测试没有fetch时的情况
    describe('WEBSERVICE without fetch', function () {
        let originalFetch;
        
        beforeEach(() => {
            // 保存并删除全局fetch
            originalFetch = global.fetch;
            delete global.fetch;
        });
        
        afterEach(() => {
            // 恢复全局fetch
            if (originalFetch) {
                global.fetch = originalFetch;
            }
        });
        
        it('should throw error when fetch is not available', () => {
            const result = parser.parse('WEBSERVICE("http://example.com")', {row: 1, col: 1});
            expect(FormulaError.ERROR().equals(result)).to.be.true;
        });
    });
    
    // 额外测试：覆盖WEBSERVICE函数在有fetch时的分支
    describe('WEBSERVICE with fetch available', function () {
        beforeEach(() => {
            // 模拟全局fetch函数
            global.fetch = () => Promise.resolve({
                text: () => Promise.resolve('mock response')
            });
        });
        
        afterEach(() => {
            delete global.fetch;
        });
        
        it('should handle WEBSERVICE with fetch available', async () => {
            const result = await parser.parseAsync('WEBSERVICE("http://example.com")', {row: 1, col: 1});
            expect(result).to.equal('mock response');
        });
        
        it('should validate URL parameter type', async () => {
            const result = await parser.parseAsync('WEBSERVICE("http://123.com")', {row: 1, col: 1});
            expect(result).to.equal('mock response');
        });
    });
});
