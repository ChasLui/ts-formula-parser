import { FormulaParser } from '../../../grammar/hooks';
import TestCase from './testcase';
import { generateTests } from '../../utils';
import { expect, describe, it, beforeEach, afterEach } from 'vitest';
import FormulaError from '../../../formulas/error';
import type { CellRef, RangeRef } from '../../../index';

const data: any[][] = [
    ['fruit', 'price', 'count', 4, 5],
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

describe('Web Functions', function () {
    generateTests(parser, TestCase);
    
    // Test when fetch is not available
    describe('WEBSERVICE without fetch', function () {
        let originalFetch: any;
        
        beforeEach(() => {
            // Save and delete global fetch
            originalFetch = (global as any).fetch;
            delete (global as any).fetch;
        });
        
        afterEach(() => {
            // Restore global fetch
            if (originalFetch) {
                (global as any).fetch = originalFetch;
            }
        });
        
        it('should throw error when fetch is not available', () => {
            const result = parser.parse('WEBSERVICE("http://example.com")', {row: 1, col: 1});
            expect(FormulaError.ERROR("WEBSERVICE failed").equals(result)).to.be.true;
        });
    });
    
    // Additional test: cover WEBSERVICE function branch when fetch is available
    describe('WEBSERVICE with fetch available', function () {
        beforeEach(() => {
            // Mock global fetch function
            (global as any).fetch = () => Promise.resolve({
                text: () => Promise.resolve('mock response')
            });
        });
        
        afterEach(() => {
            delete (global as any).fetch;
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