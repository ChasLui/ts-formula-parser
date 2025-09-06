/**
 * Common types for Excel formula functions
 */

/**
 * Common type for Excel function arguments
 * Represents all possible values that can be passed to Excel functions
 */
export type ExcelValue = any;

/**
 * Function result type
 * Represents all possible return values from Excel functions
 */
export type FunctionResult = any;

/**
 * Excel array/range type
 */
export type ExcelArray = any[][];

/**
 * Excel criteria type for filtering functions
 */
export type ExcelCriteria = string | number | boolean;

/**
 * Excel date/time value type
 */
export type ExcelDateTime = number | Date | string;