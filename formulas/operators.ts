import FormulaError from './error';
import {FormulaHelpers} from './helpers';

/**
 * Type mapping for comparison operations
 */
const type2Number: Record<string, number> = {'boolean': 3, 'string': 2, 'number': 1};

/**
 * Prefix operators (unary operators like +, -)
 */
const Prefix = {
    /**
     * Process unary operators
     * @param prefixes - Array of prefix operators
     * @param value - Value to apply operators to
     * @param isArray - Whether the value is from an array
     * @returns Processed value
     */
    unaryOp: (prefixes: string[], value: any, isArray: boolean): any => {
        let sign = 1;
        prefixes.forEach(prefix => {
            if (prefix === '+') {
                // No change for positive
            } else if (prefix === '-') {
                sign = -sign;
            } else {
                throw new Error(`Unrecognized prefix: ${prefix}`);
            }
        });

        if (value == null) {
            value = 0;
        }
        // positive means no changes
        if (sign === 1) {
            return value;
        }
        // negative
        try {
            value = FormulaHelpers.acceptNumber(value, isArray);
        } catch (e) {
            if (e instanceof FormulaError) {
                // parse number fails
                if (Array.isArray(value))
                    value = value[0][0]
            } else
                throw e;
        }

        if (typeof value === "number" && isNaN(value)) return FormulaError.VALUE;
        return -value;
    }
};

/**
 * Postfix operators (like %)
 */
const Postfix = {
    /**
     * Process percent operator
     * @param value - Value to apply percent to
     * @param postfix - Postfix operator
     * @param isArray - Whether the value is from an array
     * @returns Processed value or error
     */
    percentOp: (value: any, postfix: string, isArray: boolean): number | FormulaError => {
        try {
            value = FormulaHelpers.acceptNumber(value, isArray);
        } catch (e) {
            if (e instanceof FormulaError)
                return e;
            throw e;
        }
        if (postfix === '%') {
            return value / 100;
        }
        throw new Error(`Unrecognized postfix: ${postfix}`);
    }
};

/**
 * Infix operators (binary operators like +, -, *, /, =, <, >, etc.)
 */
const Infix = {
    /**
     * Process comparison operators
     * @param value1 - First value
     * @param infix - Comparison operator
     * @param value2 - Second value
     * @param isArray1 - Whether first value is from array
     * @param isArray2 - Whether second value is from array
     * @returns Comparison result
     */
    compareOp: (value1: any, infix: string, value2: any, isArray1: boolean, isArray2: boolean): boolean => {
        if (value1 == null) value1 = 0;
        if (value2 == null) value2 = 0;
        // for array: {1,2,3}, get the first element to compare
        if (isArray1) {
            value1 = value1[0][0];
        }
        if (isArray2) {
            value2 = value2[0][0];
        }

        const type1 = typeof value1, type2 = typeof value2;

        if (type1 === type2) {
            // same type comparison
            switch (infix) {
                case '=':
                    return value1 === value2;
                case '>':
                    return value1 > value2;
                case '<':
                    return value1 < value2;
                case '<>':
                    return value1 !== value2;
                case '<=':
                    return value1 <= value2;
                case '>=':
                    return value1 >= value2;
            }
        } else {
            switch (infix) {
                case '=':
                    return false;
                case '>':
                    return type2Number[type1] > type2Number[type2];
                case '<':
                    return type2Number[type1] < type2Number[type2];
                case '<>':
                    return true;
                case '<=':
                    return type2Number[type1] <= type2Number[type2];
                case '>=':
                    return type2Number[type1] >= type2Number[type2];
            }
        }
        throw new Error('Infix.compareOp: Should not reach here.');
    },

    /**
     * Process concatenation operator
     * @param value1 - First value
     * @param infix - Concatenation operator
     * @param value2 - Second value
     * @param isArray1 - Whether first value is from array
     * @param isArray2 - Whether second value is from array
     * @returns Concatenated string
     */
    concatOp: (value1: any, infix: string, value2: any, isArray1: boolean, isArray2: boolean): string => {
        if (value1 == null) value1 = '';
        if (value2 == null) value2 = '';
        // for array: {1,2,3}, get the first element to concat
        if (isArray1) {
            value1 = value1[0][0];
        }
        if (isArray2) {
            value2 = value2[0][0];
        }

        const type1 = typeof value1, type2 = typeof value2;
        // convert boolean to string
        if (type1 === 'boolean')
            value1 = value1 ? 'TRUE' : 'FALSE';
        if (type2 === 'boolean')
            value2 = value2 ? 'TRUE' : 'FALSE';
        return '' + value1 + value2;
    },

    /**
     * Process mathematical operators
     * @param value1 - First value
     * @param infix - Mathematical operator
     * @param value2 - Second value
     * @param isArray1 - Whether first value is from array
     * @param isArray2 - Whether second value is from array
     * @returns Mathematical result or error
     */
    mathOp: (value1: any, infix: string, value2: any, isArray1: boolean, isArray2: boolean): number | FormulaError => {
        if (value1 == null) value1 = 0;
        if (value2 == null) value2 = 0;

        try {
            value1 = FormulaHelpers.acceptNumber(value1, isArray1);
            value2 = FormulaHelpers.acceptNumber(value2, isArray2);
        } catch (e) {
            if (e instanceof FormulaError)
                return e;
            throw e;
        }

        switch (infix) {
            case '+':
                return value1 + value2;
            case '-':
                return value1 - value2;
            case '*':
                return value1 * value2;
            case '/':
                if (value2 === 0)
                    return FormulaError.DIV0;
                return value1 / value2;
            case '^':
                return Math.pow(value1, value2);
        }

        throw new Error('Infix.mathOp: Should not reach here.');
    },
};

/**
 * Operator categories for easy reference
 */
export const Operators = {
    compareOp: ['<', '>', '=', '<>', '<=', '>='] as const,
    concatOp: ['&'] as const,
    mathOp: ['+', '-', '*', '/', '^'] as const,
};

/**
 * Type definitions for operator categories
 */
export type CompareOperator = typeof Operators.compareOp[number];
export type ConcatOperator = typeof Operators.concatOp[number];
export type MathOperator = typeof Operators.mathOp[number];

export {
    Prefix,
    Postfix,
    Infix
};