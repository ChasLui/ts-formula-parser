import type { ErrorDetails } from '../types/index';

/**
 * Represents Excel formula errors with proper error codes and messaging.
 * 
 * This class extends the standard Error class to provide Excel-compatible
 * error handling with specific error codes like #VALUE!, #REF!, #NUM!, etc.
 * Implements singleton pattern for common errors to improve performance.
 * 
 * @class FormulaError
 * @extends Error
 * @since 1.0.0
 * 
 * @example
 * // Create a VALUE error
 * throw FormulaError.VALUE;
 * 
 * @example
 * // Create custom error with details
 * throw new FormulaError('#CUSTOM!', 'Custom error message', {code: 'CUSTOM'});
 */
class FormulaError extends Error {
    private _error!: string;
    readonly details?: ErrorDetails;

    /**
     * Static error map for singleton error instances.
     * 
     * Maintains a cache of common error instances to avoid creating
     * multiple objects for the same error type, improving performance
     * and memory usage.
     * 
     * @static
     * @readonly
     * @type {Map<string, FormulaError>}
     * @private
     */
    static readonly errorMap = new Map<string, FormulaError>();

    /**
     * Creates a new FormulaError instance.
     * 
     * If creating a basic error without message or details, returns a cached
     * singleton instance for better performance. Otherwise creates a new instance.
     * 
     * @param {string} error - Excel error code (e.g., '#VALUE!', '#REF!', '#NUM!')
     * @param {string} [msg] - Optional detailed error message
     * @param {ErrorDetails} [details] - Optional additional error context
     * 
     * @example
     * // Create basic error (returns singleton)
     * const error = new FormulaError('#VALUE!');
     * 
     * @example
     * // Create detailed error (new instance)
     * const error = new FormulaError('#NUM!', 'Invalid number range', {range: 'A1:B2'});
     */
    constructor(error: string, msg?: string, details?: ErrorDetails) {
        super(msg);
        if (msg == null && details == null && FormulaError.errorMap.has(error))
            return FormulaError.errorMap.get(error)!;
        else if (msg == null && details == null) {
            this._error = error;
            FormulaError.errorMap.set(error, this);
        } else {
            this._error = error;
        }
        this.details = details;
    }

    /**
     * Gets the Excel error code string.
     * 
     * @returns {string} The Excel error code (e.g., '#VALUE!', '#REF!')
     * 
     * @example
     * const error = FormulaError.VALUE;
     * console.log(error.error); // "#VALUE!"
     * 
     * @since 1.0.0
     */
    get error(): string {
        return this._error;
    }
    
    /**
     * Gets the error name (alias for error code).
     * 
     * @returns {string} The Excel error code
     * @since 1.0.0
     */
    get name(): string {
        return this._error;
    }

    /**
     * Compares two FormulaError instances for equality.
     * 
     * Two errors are considered equal if they have the same error code,
     * regardless of their message or details.
     * 
     * @param {FormulaError} err - The error to compare with
     * @returns {boolean} True if both errors have the same error code
     * 
     * @example
     * const error1 = FormulaError.VALUE;
     * const error2 = new FormulaError('#VALUE!');
     * console.log(error1.equals(error2)); // true
     * 
     * @since 1.0.0
     */
    equals(err: FormulaError): boolean {
        return err instanceof FormulaError && err._error === this._error;
    }

    /**
     * Returns the string representation of the error.
     * 
     * @returns {string} The Excel error code as a string
     * 
     * @example
     * const error = FormulaError.VALUE;
     * console.log(error.toString()); // "#VALUE!"
     * console.log(String(error)); // "#VALUE!"
     * 
     * @since 1.0.0
     */
    toString(): string {
        return this._error;
    }

    /**
     * Division by zero error (#DIV/0!).
     * 
     * Thrown when a formula attempts to divide by zero or when a formula
     * contains a reference that causes division by zero.
     * 
     * @static
     * @readonly
     * @type {FormulaError}
     * @since 1.0.0
     */
    static readonly DIV0 = new FormulaError("#DIV/0!");

    /**
     * Not available error (#N/A).
     * 
     * Thrown when a value is not available to a function or formula,
     * such as when a lookup function cannot find a match.
     * 
     * @static
     * @readonly
     * @type {FormulaError}
     * @since 1.0.0
     */
    static readonly NA = new FormulaError("#N/A");

    /**
     * NAME error
     */
    static readonly NAME = new FormulaError("#NAME?");

    /**
     * NULL error
     */
    static readonly NULL = new FormulaError("#NULL!");

    /**
     * NUM error
     */
    static readonly NUM = new FormulaError("#NUM!");

    /**
     * REF error
     */
    static readonly REF = new FormulaError("#REF!");

    /**
     * VALUE error
     */
    static readonly VALUE = new FormulaError("#VALUE!");

    /**
     * NOT_IMPLEMENTED error
     * @param functionName - the name of the not implemented function
     * @returns FormulaError instance
     */
    static NOT_IMPLEMENTED = (functionName: string): FormulaError => {
        return new FormulaError("#NAME?", `Function ${functionName} is not implemented.`);
    };

    /**
     * TOO_MANY_ARGS error
     * @param functionName - the name of the errored function
     * @returns FormulaError instance
     */
    static TOO_MANY_ARGS = (functionName: string): FormulaError => {
        return new FormulaError("#N/A", `Function ${functionName} has too many arguments.`);
    };

    /**
     * ARG_MISSING error
     * @param args - the missing argument types
     * @returns FormulaError instance
     */
    static ARG_MISSING = (args: string[]): FormulaError => {
        // Import Types dynamically to avoid circular dependency
        return new FormulaError("#N/A", `Argument type ${args.join(', ')} is missing.`);
    };

    /**
     * #ERROR!
     * Parse/Lex error or other unexpected errors
     * @param msg - error message
     * @param details - additional error details
     * @return FormulaError instance
     */
    static ERROR = (msg: string, details?: ErrorDetails): FormulaError => {
        return new FormulaError('#ERROR!', msg, details);
    };
}

export default FormulaError;