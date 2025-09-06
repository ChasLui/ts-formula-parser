import type { ExcelValue, FunctionResult } from "./types";
import FormulaError from '../error';
import {FormulaHelpers, Types, WildCard} from '../helpers';
const H = FormulaHelpers;

/**
 * Collection of Excel text manipulation functions.
 * 
 * This interface defines all text-related functions available in the formula parser,
 * following Excel's text function specifications for string manipulation, formatting,
 * and character encoding operations.
 * 
 * @interface TextFunctionCollection
 * @since 1.0.0
 */
interface TextFunctionCollection {
    FIND(findText: ExcelValue, withinText: ExcelValue, startNum?: ExcelValue): FunctionResult;
    FINDB(findText: ExcelValue, withinText: ExcelValue, startNum?: ExcelValue): FunctionResult;
    LEFTB(text: ExcelValue, numBytes?: ExcelValue): FunctionResult;
    LENB(text: ExcelValue): FunctionResult;
    MIDB(text: ExcelValue, startNum: ExcelValue, numBytes: ExcelValue): FunctionResult;
    PHONETIC(text: ExcelValue): FunctionResult;
    REPLACEB(oldText: ExcelValue, startNum: ExcelValue, numBytes: ExcelValue, newText: ExcelValue): FunctionResult;
    RIGHTB(text: ExcelValue, numBytes?: ExcelValue): FunctionResult;
    SEARCHB(findText: ExcelValue, withinText: ExcelValue, startNum?: ExcelValue): FunctionResult;
    [key: string]: any;
}

/** Spreadsheet number formatting utilities */
import ssf from '../../ssf/ssf';

/** Thai baht text conversion utility */
import { bahttext } from 'bahttext';

/**
 * Character set definitions for full-width and half-width character conversions.
 * Supports Latin, Hangul, Kana, and other extended character sets.
 * 
 * @constant {Object} charsets
 * @private
 */
const charsets = {
    latin: {halfRE: /[!-~]/g, fullRE: /[！-～]/g, delta: 0xFEE0},
    hangul1: {halfRE: /[ﾡ-ﾾ]/g, fullRE: /[ᆨ-ᇂ]/g, delta: -0xEDF9},
    hangul2: {halfRE: /[ￂ-ￜ]/g, fullRE: /[ᅡ-ᅵ]/g, delta: -0xEE61},
    kana: {
        delta: 0,
        half: "｡｢｣､･ｦｧｨｩｪｫｬｭｮｯｰｱｲｳｴｵｶｷｸｹｺｻｼｽｾｿﾀﾁﾂﾃﾄﾅﾆﾇﾈﾉﾊﾋﾌﾍﾎﾏﾐﾑﾒﾓﾔﾕﾖﾗﾘﾙﾚﾛﾜﾝﾞﾟ",
        full: "。「」、・ヲァィゥェォャュョッーアイウエオカキクケコサシ" +
            "スセソタチツテトナニヌネノハヒフヘホマミムメモヤユヨラリルレロワン゛゜"
    },
    extras: {
        delta: 0,
        half: "¢£¬¯¦¥₩\u0020|←↑→↓■°",
        full: "￠￡￢￣￤￥￦\u3000￨￩￪￫￬￭￮"
    }
};
/**
 * Converts a character to its full-width equivalent.
 * 
 * @param {Object} set - Character set configuration object
 * @param {string} c - Character to convert
 * @returns {string} Full-width character
 * @private
 */
const toFull = set => c => set.delta ?
    String.fromCharCode(c.charCodeAt(0) + set.delta) :
    [...set.full][[...set.half].indexOf(c)];
/**
 * Converts a character to its half-width equivalent.
 * 
 * @param {Object} set - Character set configuration object
 * @param {string} c - Character to convert
 * @returns {string} Half-width character
 * @private
 */
const toHalf = set => c => set.delta ?
    String.fromCharCode(c.charCodeAt(0) - set.delta) :
    [...set.half][[...set.full].indexOf(c)];
/**
 * Creates or retrieves a regular expression for character matching.
 * 
 * @param {Object} set - Character set configuration
 * @param {string} way - Direction: 'half' or 'full'
 * @returns {RegExp} Regular expression for matching characters
 * @private
 */
const re = (set, way) => set[way + "RE"] || new RegExp("[" + set[way] + "]", "g");
/** Array of all character set configurations */
const sets = Object.keys(charsets).map(i => charsets[i]);
/**
 * Converts all half-width characters in a string to full-width equivalents.
 * 
 * @param {string} str0 - Input string
 * @returns {string} String with full-width characters
 * @private
 */
const toFullWidth = str0 =>
    sets.reduce((str, set) => str.replace(re(set, "half"), toFull(set)), str0);
/**
 * Converts all full-width characters in a string to half-width equivalents.
 * 
 * @param {string} str0 - Input string
 * @returns {string} String with half-width characters
 * @private
 */
const toHalfWidth = str0 =>
    sets.reduce((str, set) => str.replace(re(set, "full"), toHalf(set)), str0);

/**
 * Unicode character segmentation constants.
 * Used for Excel-style character segmentation. Prefers Intl.Segmenter when available,
 * falls back to custom algorithm for environments without native support.
 * 
 * @namespace UnicodeConstants
 * @since 1.0.0
 */

/** Start of high surrogate pairs (U+D800) */
const HIGH_SURROGATE_START = 0xd800;
/** End of high surrogate pairs (U+DBFF) */
const HIGH_SURROGATE_END = 0xdbff;
/** Start of low surrogate pairs (U+DC00) */
const LOW_SURROGATE_START = 0xdc00;
/** Start of regional indicator symbols (U+1F1E6) */
const REGIONAL_INDICATOR_START = 0x1f1e6;
/** End of regional indicator symbols (U+1F1FF) */
const REGIONAL_INDICATOR_END = 0x1f1ff;
/** Start of Fitzpatrick skin tone modifiers (U+1F3FB) */
const FITZPATRICK_MODIFIER_START = 0x1f3fb;
/** End of Fitzpatrick skin tone modifiers (U+1F3FF) */
const FITZPATRICK_MODIFIER_END = 0x1f3ff;
/** Start of variation selectors (U+FE00) */
const VARIATION_MODIFIER_START = 0xfe00;
/** End of variation selectors (U+FE0F) */
const VARIATION_MODIFIER_END = 0xfe0f;
/** Start of combining diacritical marks (U+20D0) */
const DIACRITICAL_MARKS_START = 0x20d0;
/** End of combining diacritical marks (U+20FF) */
const DIACRITICAL_MARKS_END = 0x20ff;
/** Zero-width joiner character (U+200D) */
const ZWJ = 0x200d;

/**
 * Special grapheme characters that form complex clusters.
 * These characters combine with adjacent characters to form single visual units.
 * Includes combining marks from various scripts: Devanagari, Tamil, Thai, and Hangul.
 * 
 * @constant {number[]} GRAPHEMS
 * @private
 */
const GRAPHEMS = [
    0x0308, // ( ◌̈ ) COMBINING DIAERESIS
    0x0937, // ( ष ) DEVANAGARI LETTER SSA
    0x0937, // ( ष ) DEVANAGARI LETTER SSA
    0x093F, // ( ि ) DEVANAGARI VOWEL SIGN I
    0x093F, // ( ि ) DEVANAGARI VOWEL SIGN I
    0x0BA8, // ( ந ) TAMIL LETTER NA
    0x0BBF, // ( ி ) TAMIL VOWEL SIGN I
    0x0BCD, // ( ◌்) TAMIL SIGN VIRAMA
    0x0E31, // ( ◌ั ) THAI CHARACTER MAI HAN-AKAT
    0x0E33, // ( ำ ) THAI CHARACTER SARA AM
    0x0E40, // ( เ ) THAI CHARACTER SARA E
    0x0E49, // ( เ ) THAI CHARACTER MAI THO
    0x1100, // ( ᄀ ) HANGUL CHOSEONG KIYEOK
    0x1161, // ( ᅡ ) HANGUL JUNGSEONG A
    0x11A8  // ( ᆨ ) HANGUL JONGSEONG KIYEOK
];

/**
 * Checks if a value is within an inclusive range.
 * 
 * @param {number} value - Value to check
 * @param {number} lower - Lower bound (inclusive)
 * @param {number} upper - Upper bound (inclusive)
 * @returns {boolean} True if value is within range
 * @private
 */
const betweenInclusive = (value, lower, upper) => value >= lower && value <= upper;

/**
 * Determines if the first character of a string is a high surrogate.
 * High surrogates are the first part of UTF-16 surrogate pairs for characters outside BMP.
 * 
 * @param {string} string - String to check
 * @returns {boolean} True if first character is a high surrogate
 * @private
 */
const isFirstOfSurrogatePair = (string) => {
    return string && betweenInclusive(string[0].charCodeAt(0), HIGH_SURROGATE_START, HIGH_SURROGATE_END);
};

/**
 * Calculates the Unicode code point from a UTF-16 surrogate pair.
 * 
 * @param {string} pair - String containing surrogate pair (2 characters)
 * @returns {number} Unicode code point value
 * @private
 */
const codePointFromSurrogatePair = (pair) => {
    const highOffset = pair.charCodeAt(0) - HIGH_SURROGATE_START;
    const lowOffset = pair.charCodeAt(1) - LOW_SURROGATE_START;
    return (highOffset << 10) + lowOffset + 0x10000;
};

/**
 * Checks if a string starts with a regional indicator symbol (flag emojis).
 * 
 * @param {string} string - String to check
 * @returns {boolean} True if string starts with regional indicator
 * @private
 */
const isRegionalIndicator = (string) => {
    return betweenInclusive(codePointFromSurrogatePair(string), REGIONAL_INDICATOR_START, REGIONAL_INDICATOR_END);
};

/**
 * Checks if a string starts with a Fitzpatrick skin tone modifier.
 * 
 * @param {string} string - String to check
 * @returns {boolean} True if string starts with Fitzpatrick modifier
 * @private
 */
const isFitzpatrickModifier = (string) => {
    return betweenInclusive(codePointFromSurrogatePair(string), FITZPATRICK_MODIFIER_START, FITZPATRICK_MODIFIER_END);
};

/**
 * Checks if a string starts with a variation selector.
 * Variation selectors modify the appearance of preceding characters.
 * 
 * @param {string} string - String to check
 * @returns {boolean} True if string starts with variation selector
 * @private
 */
const isVariationSelector = (string) => {
    return typeof string === 'string' && betweenInclusive(string.charCodeAt(0), VARIATION_MODIFIER_START, VARIATION_MODIFIER_END);
};

/**
 * Checks if a string starts with a combining diacritical mark.
 * Diacritical marks combine with base characters to modify pronunciation or meaning.
 * 
 * @param {string} string - String to check
 * @returns {boolean} True if string starts with diacritical mark
 * @private
 */
const isDiacriticalMark = (string) => {
    return typeof string === 'string' && betweenInclusive(string.charCodeAt(0), DIACRITICAL_MARKS_START, DIACRITICAL_MARKS_END);
};

/**
 * Checks if a string starts with a special grapheme character.
 * These are script-specific combining characters from the GRAPHEMS array.
 * 
 * @param {string} string - String to check
 * @returns {boolean} True if string starts with special grapheme
 * @private
 */
const isGraphem = (string) => {
    return typeof string === 'string' && GRAPHEMS.indexOf(string.charCodeAt(0)) !== -1;
};

/**
 * Checks if a string starts with a zero-width joiner character.
 * ZWJ is used to create complex emojis and ligatures.
 * 
 * @param {string} string - String to check
 * @returns {boolean} True if string starts with ZWJ
 * @private
 */
const isZeroWidthJoiner = (string) => {
    return typeof string === 'string' && string.charCodeAt(0) === ZWJ;
};

const nextUnits = (i, string) => {
    const current = string[i];
    if (!isFirstOfSurrogatePair(current) || i === string.length - 1) {
        return 1;
    }

    const currentPair = current + string[i + 1];
    let nextPair = string.substring(i + 2, i + 5);

    if (isRegionalIndicator(currentPair) && isRegionalIndicator(nextPair)) {
        return 4;
    }

    if (isFitzpatrickModifier(nextPair)) {
        return 4;
    }
    return 2;
};

// Fallback character segmentation function (based on the provided algorithm)
const runesLegacy = (string) => {
    if (typeof string !== 'string') {
        throw new Error('string cannot be undefined or null');
    }
    const result: string[] = [];
    let i = 0;
    let increment = 0;
    while (i < string.length) {
        increment += nextUnits(i + increment, string);
        if (isGraphem(string[i + increment])) {
            increment++;
        }
        if (isVariationSelector(string[i + increment])) {
            increment++;
        }
        if (isDiacriticalMark(string[i + increment])) {
            increment++;
        }
        if (isZeroWidthJoiner(string[i + increment])) {
            increment++;
            continue;
        }
        result.push(string.substring(i, i + increment));
        i += increment;
        increment = 0;
    }
    return result;
};

// Excel-style character segmentation function - prefer to use Intl.Segmenter
const getExcelGraphemeClusters = (string) => {
    if (typeof string !== 'string') {
        throw new Error('string cannot be undefined or null');
    }
    
    // Check if Intl.Segmenter is supported
    if (typeof Intl !== 'undefined' && Intl.Segmenter) {
        try {
            const segmenter = new Intl.Segmenter('en', { granularity: 'grapheme' });
            return Array.from(segmenter.segment(string), s => s.segment);
        } catch (e) {
            // If Intl.Segmenter is unavailable or an error occurs, fall back to the traditional method.
            return runesLegacy(string);
        }
    }
    
    // Fallback to the traditional method
    return runesLegacy(string);
};

/**
 * Collection of Excel text functions.
 * 
 * Implements all text manipulation functions available in Excel, including string operations,
 * character encoding conversions, search and replace functionality, and formatting utilities.
 * All functions follow Excel's behavior and error handling patterns.
 * 
 * @constant {TextFunctionCollection} TextFunctions
 * @since 1.0.0
 */
const TextFunctions: TextFunctionCollection = {
    /**
     * Converts full-width characters in text to half-width equivalents.
     * 
     * This function converts double-byte (full-width) characters to their
     * single-byte (half-width) equivalents. Primarily used for Asian text processing.
     * 
     * @param {ExcelValue} text - Text containing full-width characters to convert
     * @returns {string} Text with half-width characters
     * 
     * @example
     * // Convert full-width numbers to half-width
     * ASC("０１２３") // Returns "0123"
     * 
     * @example
     * // Convert full-width letters
     * ASC("ＡＢＣ") // Returns "ABC"
     * 
     * @throws {FormulaError.VALUE} When text cannot be converted to string
     * @since 1.0.0
     */
    ASC: (text) => {
        text = H.accept(text, Types.STRING);
        return toHalfWidth(text);
    },

    /**
     * Converts a number to Thai baht text representation.
     * 
     * Converts numeric values to Thai currency text format, spelling out
     * the number in Thai language with appropriate currency designations.
     * 
     * @param {ExcelValue} number - Numeric value to convert to Thai baht text
     * @returns {string} Thai text representation of the number as currency
     * 
     * @example
     * // Convert number to Thai baht text
     * BAHTTEXT(1234) // Returns Thai text for "one thousand two hundred thirty-four baht"
     * 
     * @throws {FormulaError.VALUE} When number cannot be converted to numeric value
     * @throws {Error} When bahttext conversion fails
     * @since 1.0.0
     */
    BAHTTEXT: (number) => {
        number = H.accept(number, Types.NUMBER);
        try {
            return bahttext(number);
        } catch (e: unknown) {
            throw Error(`Error in https://github.com/jojoee/bahttext \n${String(e)}`)
        }
    },

    /**
     * Returns the character corresponding to an ASCII code number.
     * 
     * Converts a numeric ASCII code (1-255) to its corresponding character.
     * This function only works with single-byte character codes.
     * 
     * @param {ExcelValue} number - ASCII code number (1-255)
     * @returns {string} Character corresponding to the ASCII code
     * 
     * @example
     * // Get character for ASCII code 65
     * CHAR(65) // Returns "A"
     * 
     * @example
     * // Get character for ASCII code 97
     * CHAR(97) // Returns "a"
     * 
     * @throws {FormulaError.VALUE} When number is not between 1 and 255
     * @throws {FormulaError.VALUE} When number cannot be converted to numeric value
     * @since 1.0.0
     */
    CHAR: (number) => {
        number = H.accept(number, Types.NUMBER);
        if (number > 255 || number < 1)
            throw FormulaError.VALUE;
        return String.fromCharCode(number);
    },

    /**
     * Removes non-printable control characters from text.
     * 
     * Removes all characters with ASCII values 0-31 (control characters)
     * from the text string, leaving only printable characters.
     * 
     * @param {ExcelValue} text - Text to clean of control characters
     * @returns {string} Text with control characters removed
     * 
     * @example
     * // Remove control characters from text
     * CLEAN("Hello\x07World\x1F") // Returns "HelloWorld"
     * 
     * @throws {FormulaError.VALUE} When text cannot be converted to string
     * @since 1.0.0
     */
    CLEAN: (text) => {
        text = H.accept(text, Types.STRING);
        return text.replace(/[\x00-\x1F]/g, '');
    },

    /**
     * Returns the Unicode code point of the first character in text.
     * 
     * Gets the Unicode code point (not just ASCII) of the first character
     * in the text string. Supports full Unicode range including emojis.
     * 
     * @param {ExcelValue} text - Text string (must not be empty)
     * @returns {number} Unicode code point of first character
     * 
     * @example
     * // Get code point of ASCII character
     * CODE("A") // Returns 65
     * 
     * @example
     * // Get code point of Unicode character
     * CODE("中") // Returns 20013 (Chinese character)
     * 
     * @throws {FormulaError.VALUE} When text is empty string
     * @throws {FormulaError.VALUE} When text cannot be converted to string
     * @since 1.0.0
     */
    CODE: (text) => {
        text = H.accept(text, Types.STRING);
        if (text.length === 0)
            throw FormulaError.VALUE;
        return text.codePointAt(0);
    },

    /**
     * Concatenates text from multiple ranges, arrays, and individual strings.
     * 
     * Combines text from multiple sources including ranges and arrays.
     * Unlike CONCATENATE, this function can handle ranges and arrays,
     * flattening them and joining all text values.
     * 
     * @param {...ExcelValue} params - Variable number of text values, ranges, or arrays
     * @returns {string} Concatenated text string
     * 
     * @example
     * // Concatenate multiple strings
     * CONCAT("Hello", " ", "World") // Returns "Hello World"
     * 
     * @example
     * // Concatenate range values (conceptual)
     * CONCAT(A1:A3) // Joins values from range A1:A3
     * 
     * @since 1.0.0
     */
    CONCAT: (...params) => {
        let text = '';
        // does not allow union
        H.flattenParams(params, Types.STRING, false, item => {
            item = H.accept(item, Types.STRING);
            text += item;
        });
        return text
    },

    /**
     * Joins individual text strings into one text string.
     * 
     * Classic Excel function for concatenating individual text arguments.
     * Does not support ranges or arrays - each argument must be a single value.
     * Requires at least one argument.
     * 
     * @param {...ExcelValue} params - Individual text values to concatenate
     * @returns {string} Concatenated text string
     * 
     * @example
     * // Concatenate individual strings
     * CONCATENATE("Hello", " ", "World") // Returns "Hello World"
     * 
     * @example
     * // Concatenate mixed types (converted to strings)
     * CONCATENATE("Value: ", 123) // Returns "Value: 123"
     * 
     * @throws {Error} When no arguments are provided
     * @since 1.0.0
     */
    CONCATENATE: (...params) => {
        let text = '';
        if (params.length === 0)
            throw Error('CONCATENATE need at least one argument.');
        params.forEach(param => {
            // does not allow range reference, array, union
            param = H.accept(param, Types.STRING);
            text += param;
        });

        return text;
    },

    /**
     * Converts half-width characters to full-width equivalents.
     * 
     * Converts single-byte (half-width) characters to their double-byte
     * (full-width) equivalents. This is the opposite of the ASC function.
     * Primarily used for Asian text processing.
     * 
     * @param {ExcelValue} text - Text containing half-width characters to convert
     * @returns {string} Text with full-width characters
     * 
     * @example
     * // Convert half-width numbers to full-width
     * DBCS("0123") // Returns "０１２３"
     * 
     * @example
     * // Convert half-width letters
     * DBCS("ABC") // Returns "ＡＢＣ"
     * 
     * @throws {FormulaError.VALUE} When text cannot be converted to string
     * @since 1.0.0
     */
    DBCS: (text) => {
        text = H.accept(text, Types.STRING);
        return toFullWidth(text);
    },

    DOLLAR: (number, decimals) => {
        number = H.accept(number, Types.NUMBER);
        decimals = H.accept(decimals, Types.NUMBER, 2);
        const decimalString = Array(decimals).fill('0').join('');
        // Note: does not support locales
        // TODO: change currency based on user locale or settings from this library
        return ssf.format(`$#,##0.${decimalString}_);($#,##0.${decimalString})`, number).trim();
    },

    EXACT: (text1, text2) => {
        text1 = H.accept(text1, [Types.STRING]);
        text2 = H.accept(text2, [Types.STRING]);

        return text1 === text2;
    },

    FIND: (findText, withinText, startNum) => {
        findText = H.accept(findText, Types.STRING);
        withinText = H.accept(withinText, Types.STRING);
        startNum = H.accept(startNum, Types.NUMBER, 1);
        const withinTextArray = getExcelGraphemeClusters(withinText);
        if (startNum < 1 || startNum > withinTextArray.length)
            throw FormulaError.VALUE;
        // Convert character position to string index for searching
        let searchStartIndex = 0;
        if (startNum > 1) {
            searchStartIndex = withinTextArray.slice(0, startNum - 1).join('').length;
        }
        const res = withinText.indexOf(findText, searchStartIndex);
        if (res === -1)
            throw FormulaError.VALUE;
        // Convert string index back to character position
        const charPosition = getExcelGraphemeClusters(withinText.substring(0, res)).length + 1;
        return charPosition;
    },

    FINDB: (findText, withinText, startNum) => {
        return TextFunctions.FIND(findText, withinText, startNum);
    },

    FIXED: (number, decimals, noCommas) => {
        number = H.accept(number, Types.NUMBER);
        decimals = H.accept(decimals, Types.NUMBER, 2);
        noCommas = H.accept(noCommas, Types.BOOLEAN, false);

        const decimalString = Array(decimals).fill('0').join('');
        const comma = noCommas ? '' : '#,';
        return ssf.format(`${comma}##0.${decimalString}_);(${comma}##0.${decimalString})`, number).trim();
    },

    LEFT: (text, numChars) => {
        text = H.accept(text, Types.STRING);
        numChars = H.accept(numChars, Types.NUMBER, 1);

        if (numChars < 0)
            throw FormulaError.VALUE;
        const textArray = getExcelGraphemeClusters(text);
        if (numChars > textArray.length)
            return text;
        return textArray.slice(0, numChars).join('');
    },

    LEFTB: (text, numBytes) => {
        return TextFunctions.LEFT(text, numBytes);
    },

    /**
     * Returns the number of characters in a text string.
     * 
     * Counts the number of characters in text using Excel-compatible
     * grapheme cluster segmentation. Properly handles complex characters
     * like emojis, combining marks, and multi-byte Unicode sequences.
     * 
     * @param {ExcelValue} text - Text string to measure
     * @returns {number} Number of characters in the text
     * 
     * @example
     * // Count regular ASCII characters
     * LEN("Hello") // Returns 5
     * 
     * @example
     * // Count Unicode characters properly
     * LEN("👩‍💻") // Returns 1 (woman technologist emoji)
     * 
     * @example
     * // Count mixed text
     * LEN("Hello 世界") // Returns 8 (5 ASCII + 1 space + 2 Chinese)
     * 
     * @throws {FormulaError.VALUE} When text cannot be converted to string
     * @since 1.0.0
     */
    LEN: (text) => {
        text = H.accept(text, Types.STRING);
        return getExcelGraphemeClusters(text).length;
    },

    LENB: (text) => {
        return TextFunctions.LEN(text);
    },

    /**
     * Converts text to lowercase.
     * 
     * Converts all uppercase letters in the text string to lowercase.
     * Numbers, punctuation, and special characters remain unchanged.
     * 
     * @param {ExcelValue} text - Text to convert to lowercase
     * @returns {string} Text converted to lowercase
     * 
     * @example
     * // Convert mixed case to lowercase
     * LOWER("Hello World") // Returns "hello world"
     * 
     * @example
     * // Mixed content unchanged for non-letters
     * LOWER("ABC123!@#") // Returns "abc123!@#"
     * 
     * @throws {FormulaError.VALUE} When text cannot be converted to string
     * @since 1.0.0
     */
    LOWER: (text) => {
        text = H.accept(text, Types.STRING);
        return text.toLowerCase();
    },

    MID: (text, startNum, numChars) => {
        text = H.accept(text, Types.STRING);
        startNum = H.accept(startNum, Types.NUMBER);
        numChars = H.accept(numChars, Types.NUMBER);
        const textArray = getExcelGraphemeClusters(text);
        if (startNum > textArray.length)
            return '';
        if (startNum < 1 || numChars < 1)
            throw FormulaError.VALUE;
        return textArray.slice(startNum - 1, startNum + numChars - 1).join('');
    },

    MIDB: (text, startNum, numBytes) => {
        return TextFunctions.MID(text, startNum, numBytes);
    },

    NUMBERVALUE: (text, decimalSeparator, groupSeparator) => {
        text = H.accept(text, Types.STRING);
        // TODO: support reading system locale and set separators
        decimalSeparator = H.accept(decimalSeparator, Types.STRING, '.');
        groupSeparator = H.accept(groupSeparator, Types.STRING, ',');

        if (text.length === 0)
            return 0;
        if (decimalSeparator.length === 0 || groupSeparator.length === 0)
            throw FormulaError.VALUE;
        decimalSeparator = decimalSeparator[0];
        groupSeparator = groupSeparator[0];
        if (decimalSeparator === groupSeparator
            || text.indexOf(decimalSeparator) < text.lastIndexOf(groupSeparator))
            throw FormulaError.VALUE;

        const res = text.replace(groupSeparator, '')
            .replace(decimalSeparator, '.')
            // remove chars that not related to number
            .replace(/[^\-0-9.%()]/g, '')
            .match(/([(-]*)([0-9]*[.]*[0-9]+)([)]?)([%]*)/);
        if (!res)
            throw FormulaError.VALUE;
        // ["-123456.78%%", "(-", "123456.78", ")", "%%"]
        const leftParenOrMinus = res[1].length, rightParen = res[3].length, percent = res[4].length;
        let number = Number(res[2]);
        if (leftParenOrMinus > 1 || leftParenOrMinus && !rightParen
            || !leftParenOrMinus && rightParen || isNaN(number))
            throw FormulaError.VALUE;
        number = number / 100 ** percent;
        return leftParenOrMinus ? -number : number;
    },

    PHONETIC: () => {
    },

    PROPER: (text) => {
        text = H.accept(text, [Types.STRING]);
        text = text.toLowerCase();
        text = text.charAt(0).toUpperCase() + text.slice(1);
        return text.replace(/(?:[^a-zA-Z])([a-zA-Z])/g,
            letter => letter.toUpperCase());
    },

    REPLACE: (old_text, start_num, num_chars, new_text) => {
        old_text = H.accept(old_text, [Types.STRING]);
        start_num = H.accept(start_num, [Types.NUMBER]);
        num_chars = H.accept(num_chars, [Types.NUMBER]);
        new_text = H.accept(new_text, [Types.STRING]);

        let arr = getExcelGraphemeClusters(old_text);
        arr.splice(start_num - 1, num_chars, new_text);

        return arr.join("");
    },

    REPLACEB: (oldText, startNum, numBytes, newText) => {
        return TextFunctions.REPLACE(oldText, startNum, numBytes, newText)
    },

    REPT: (text, number_times) => {
        text = H.accept(text, Types.STRING);
        number_times = H.accept(number_times, Types.NUMBER);
        let str = "";

        for (let i = 0; i < number_times; i++) {
            str += text;
        }
        return str;
    },

    RIGHT: (text, numChars) => {
        text = H.accept(text, Types.STRING);
        numChars = H.accept(numChars, Types.NUMBER, 1);

        if (numChars < 0)
            throw FormulaError.VALUE;
        const textArray = getExcelGraphemeClusters(text);
        const len = textArray.length;
        if (numChars > len)
            return text;
        return textArray.slice(len - numChars).join('');
    },

    RIGHTB: (text, numBytes) => {
        return TextFunctions.RIGHT(text, numBytes);
    },

    SEARCH: (findText, withinText, startNum) => {
        findText = H.accept(findText, Types.STRING);
        withinText = H.accept(withinText, Types.STRING);
        startNum = H.accept(startNum, Types.NUMBER, 1);
        const withinTextArray = getExcelGraphemeClusters(withinText);
        if (startNum < 1 || startNum > withinTextArray.length)
            throw FormulaError.VALUE;

        // transform to js regex expression
        let findTextRegex = WildCard.isWildCard(findText) ? WildCard.toRegex(findText, 'i') : findText;
        // Convert character position to string index for searching
        let searchStartIndex = 0;
        if (startNum > 1) {
            searchStartIndex = withinTextArray.slice(0, startNum - 1).join('').length;
        }
        const res = withinText.slice(searchStartIndex).search(findTextRegex);
        if (res === -1)
            throw FormulaError.VALUE;
        // Convert string index back to character position
        const foundIndex = searchStartIndex + res;
        const charPosition = getExcelGraphemeClusters(withinText.substring(0, foundIndex)).length + 1;
        return charPosition;
    },

    SEARCHB: (findText, withinText, startNum) => {
        return TextFunctions.SEARCH(findText, withinText, startNum)
    },

    SUBSTITUTE: (text, old_text, new_text, instance_num) => {
        text = H.accept(text, [Types.STRING]);
        old_text = H.accept(old_text, [Types.STRING]);
        new_text = H.accept(new_text, [Types.STRING]);
        instance_num = H.accept(instance_num, [Types.NUMBER], null);

        // If old_text is an empty string, return the original text
        if (old_text === '') {
            return text;
        }

        // If instance_num is not specified, replace all matches
        if (instance_num === null) {
            // Use regular expression to replace all matches, need to escape special characters
            const escapedOldText = old_text.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
            return text.replace(new RegExp(escapedOldText, 'g'), new_text);
        }

        // If instance_num is specified, replace the Nth match
        if (instance_num < 1) {
            throw FormulaError.VALUE;
        }

        let count = 0;
        let index = 0;
        let result = text;

        while (index < result.length) {
            const pos = result.indexOf(old_text, index);
            if (pos === -1) {
                // No more matches found
                break;
            }
            
            count++;
            if (count === instance_num) {
                // Found the Nth match, replace it
                result = result.substring(0, pos) + new_text + result.substring(pos + old_text.length);
                break;
            }
            
            index = pos + old_text.length;
        }

        return result;
    },

    T: (value) => {
        // extract the real parameter
        value = H.accept(value);
        if (typeof value === "string")
            return value;
        return '';
    },

    TEXT: (value, formatText) => {
        value = H.accept(value, Types.NUMBER);
        formatText = H.accept(formatText, Types.STRING);
        // I know ssf contains bugs...
        try {
            return ssf.format(formatText, value);
        } catch (e) {
            console.error(e)
            throw FormulaError.VALUE;
        }
    },

    TEXTJOIN: (delimiter, ignoreEmpty, ...texts) => {
        // Verify the number of parameters; at least 3 parameters are required.
        if (texts.length < 1) {
            throw new Error('TEXTJOIN requires at least 3 arguments.');
        }
        
        // Verify the number of parameters, the maximum is 252 text parameters
        if (texts.length > 252) {
            throw new Error('TEXTJOIN supports a maximum of 252 text arguments.');
        }
        
        // Process the delimiter parameter, if it is a number, convert it to text
        delimiter = H.accept(delimiter, Types.STRING);
        
        // Process the ignore_empty parameter
        ignoreEmpty = H.accept(ignoreEmpty, Types.BOOLEAN);
        
        let result: string[] = [];
        
        // Process all text parameters
        H.flattenParams(texts, Types.STRING, false, (item: any) => {
            item = H.accept(item, Types.STRING);
            // Determine whether to include empty strings based on ignoreEmpty
            if (!ignoreEmpty || item !== '') {
                result.push(item);
            }
        });
        
        // Join all text
        const joinedText = result.join(delimiter);
        
        // Check the length limit (Excel cell limit)
        if (joinedText.length > 32767) {
            throw FormulaError.VALUE;
        }
        
        return joinedText;
    },

    TRIM: (text) => {
        text = H.accept(text, [Types.STRING]);
        return text.replace(/^\s+|\s+$/g, '')
    },

    UPPER: (text) => {
        text = H.accept(text, Types.STRING);
        return text.toUpperCase();
    },

    UNICHAR: (number) => {
        number = H.accept(number, [Types.NUMBER]);
        if (number <= 0)
            throw FormulaError.VALUE;
        return String.fromCodePoint(number);
    },

    UNICODE: (text) => {
        return TextFunctions.CODE(text);
    },
};

export default TextFunctions;
