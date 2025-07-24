const FormulaError = require('../error');
const {FormulaHelpers, Types, WildCard} = require('../helpers');
const H = FormulaHelpers;

// Spreadsheet number format
const ssf = require('../../ssf/ssf');

// Change number to Thai pronunciation string
const { bahttext } = require('bahttext');

// full-width and half-width converter
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
const toFull = set => c => set.delta ?
    String.fromCharCode(c.charCodeAt(0) + set.delta) :
    [...set.full][[...set.half].indexOf(c)];
const toHalf = set => c => set.delta ?
    String.fromCharCode(c.charCodeAt(0) - set.delta) :
    [...set.half][[...set.full].indexOf(c)];
const re = (set, way) => set[way + "RE"] || new RegExp("[" + set[way] + "]", "g");
const sets = Object.keys(charsets).map(i => charsets[i]);
const toFullWidth = str0 =>
    sets.reduce((str, set) => str.replace(re(set, "half"), toFull(set)), str0);
const toHalfWidth = str0 =>
    sets.reduce((str, set) => str.replace(re(set, "full"), toHalf(set)), str0);

// Excel-style character segmentation function - prefer to use Intl.Segmenter; otherwise, use the provided algorithm.
const HIGH_SURROGATE_START = 0xd800;
const HIGH_SURROGATE_END = 0xdbff;
const LOW_SURROGATE_START = 0xdc00;
const REGIONAL_INDICATOR_START = 0x1f1e6;
const REGIONAL_INDICATOR_END = 0x1f1ff;
const FITZPATRICK_MODIFIER_START = 0x1f3fb;
const FITZPATRICK_MODIFIER_END = 0x1f3ff;
const VARIATION_MODIFIER_START = 0xfe00;
const VARIATION_MODIFIER_END = 0xfe0f;
const DIACRITICAL_MARKS_START = 0x20d0;
const DIACRITICAL_MARKS_END = 0x20ff;
const ZWJ = 0x200d;

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

// 辅助函数
const betweenInclusive = (value, lower, upper) => value >= lower && value <= upper;

const isFirstOfSurrogatePair = (string) => {
    return string && betweenInclusive(string[0].charCodeAt(0), HIGH_SURROGATE_START, HIGH_SURROGATE_END);
};

const codePointFromSurrogatePair = (pair) => {
    const highOffset = pair.charCodeAt(0) - HIGH_SURROGATE_START;
    const lowOffset = pair.charCodeAt(1) - LOW_SURROGATE_START;
    return (highOffset << 10) + lowOffset + 0x10000;
};

const isRegionalIndicator = (string) => {
    return betweenInclusive(codePointFromSurrogatePair(string), REGIONAL_INDICATOR_START, REGIONAL_INDICATOR_END);
};

const isFitzpatrickModifier = (string) => {
    return betweenInclusive(codePointFromSurrogatePair(string), FITZPATRICK_MODIFIER_START, FITZPATRICK_MODIFIER_END);
};

const isVariationSelector = (string) => {
    return typeof string === 'string' && betweenInclusive(string.charCodeAt(0), VARIATION_MODIFIER_START, VARIATION_MODIFIER_END);
};

const isDiacriticalMark = (string) => {
    return typeof string === 'string' && betweenInclusive(string.charCodeAt(0), DIACRITICAL_MARKS_START, DIACRITICAL_MARKS_END);
};

const isGraphem = (string) => {
    return typeof string === 'string' && GRAPHEMS.indexOf(string.charCodeAt(0)) !== -1;
};

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
    const result = [];
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

const TextFunctions = {
    ASC: (text) => {
        text = H.accept(text, Types.STRING);
        return toHalfWidth(text);
    },

    BAHTTEXT: (number) => {
        number = H.accept(number, Types.NUMBER);
        try {
            return bahttext(number);
        } catch (e) {
            throw Error(`Error in https://github.com/jojoee/bahttext \n${e.toString()}`)
        }
    },

    CHAR: (number) => {
        number = H.accept(number, Types.NUMBER);
        if (number > 255 || number < 1)
            throw FormulaError.VALUE;
        return String.fromCharCode(number);
    },

    CLEAN: (text) => {
        text = H.accept(text, Types.STRING);
        return text.replace(/[\x00-\x1F]/g, '');
    },

    CODE: (text) => {
        text = H.accept(text, Types.STRING);
        if (text.length === 0)
            throw FormulaError.VALUE;
        return text.codePointAt(0);
    },

    CONCAT: (...params) => {
        let text = '';
        // does not allow union
        H.flattenParams(params, Types.STRING, false, item => {
            item = H.accept(item, Types.STRING);
            text += item;
        });
        return text
    },

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

    FINDB: (...params) => {
        return TextFunctions.FIND(...params);
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

    LEFTB: (...params) => {
        return TextFunctions.LEFT(...params);
    },

    LEN: (text) => {
        text = H.accept(text, Types.STRING);
        return getExcelGraphemeClusters(text).length;
    },

    LENB: (...params) => {
        return TextFunctions.LEN(...params);
    },

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

    MIDB: (...params) => {
        return TextFunctions.MID(...params);
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

    REPLACEB: (...params) => {
        return TextFunctions.REPLACE(...params)
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

    RIGHTB: (...params) => {
        return TextFunctions.RIGHT(...params);
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

    SEARCHB: (...params) => {
        return TextFunctions.SEARCH(...params)
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
        if (arguments.length < 3) {
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
        
        let result = [];
        
        // Process all text parameters
        H.flattenParams(texts, Types.STRING, false, item => {
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

module.exports = TextFunctions;
