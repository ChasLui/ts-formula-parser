import type { ExcelValue, FunctionResult } from "./types";
import FormulaError from '../error';
import {FormulaHelpers, Types} from '../helpers';
import type { DateConfig, FormulaValue } from '../../types/index';
const H = FormulaHelpers;

const MS_PER_DAY = 1000 * 60 * 60 * 24;
const d1900 = new Date(Date.UTC(1900, 0, 1));
const WEEK_STARTS = [
    undefined, 0, 1, undefined, undefined, undefined, undefined, undefined, undefined,
    undefined, undefined, undefined, 1, 2, 3, 4, 5, 6, 0];
const WEEK_TYPES = [
    undefined,
    [1, 2, 3, 4, 5, 6, 7],
    [7, 1, 2, 3, 4, 5, 6],
    [6, 0, 1, 2, 3, 4, 5],
    undefined,
    undefined,
    undefined,
    undefined,
    undefined,
    undefined,
    undefined,
    [7, 1, 2, 3, 4, 5, 6],
    [6, 7, 1, 2, 3, 4, 5],
    [5, 6, 7, 1, 2, 3, 4],
    [4, 5, 6, 7, 1, 2, 3],
    [3, 4, 5, 6, 7, 1, 2],
    [2, 3, 4, 5, 6, 7, 1],
    [1, 2, 3, 4, 5, 6, 7]
];
const WEEKEND_TYPES = [
    undefined,
    [6, 0],
    [0, 1],
    [1, 2],
    [2, 3],
    [3, 4],
    [4, 5],
    [5, 6],
    undefined,
    undefined,
    undefined,
    [0],
    [1],
    [2],
    [3],
    [4],
    [5],
    [6]
];

// Formats: h:mm:ss A, h:mm A, H:mm, H:mm:ss, H A
const timeRegex = /^\s*(\d\d?)\s*(:\s*\d\d?)?\s*(:\s*\d\d?)?\s*(pm|am)?\s*$/i;
// 12-3, 12/3
const dateRegex1 = /^\s*((\d\d?)\s*([-\/])\s*(\d\d?))([\d:.apm\s]*)$/i;
// 3-Dec, 3/Dec
const dateRegex2 = /^\s*((\d\d?)\s*([-/])\s*(jan\w*|feb\w*|mar\w*|apr\w*|may\w*|jun\w*|jul\w*|aug\w*|sep\w*|oct\w*|nov\w*|dec\w*))([\d:.apm\s]*)$/i;
// Dec-3, Dec/3
const dateRegex3 = /^\s*((jan\w*|feb\w*|mar\w*|apr\w*|may\w*|jun\w*|jul\w*|aug\w*|sep\w*|oct\w*|nov\w*|dec\w*)\s*([-/])\s*(\d\d?))([\d:.apm\s]*)$/i;

function parseSimplifiedDate(text: string, nullDate: DateConfig | null = null): Date {
    const fmt1 = text.match(dateRegex1);
    const fmt2 = text.match(dateRegex2);
    const fmt3 = text.match(dateRegex3);
    // Use the incoming default year; if none is provided, use the current year.
    const yearToUse = nullDate && nullDate.year ? nullDate.year : new Date().getFullYear();
    if (fmt1) {
        text = fmt1[1] + fmt1[3] + yearToUse + fmt1[5];
    } else if (fmt2) {
        text = fmt2[1] + fmt2[3] + yearToUse + fmt2[5];
    } else if (fmt3) {
        text = fmt3[1] + fmt3[3] + yearToUse + fmt3[5];
    }
    return new Date(Date.parse(`${text} UTC`));
}

/**
 * Parse time string to date in UTC.
 * @param {string} text
 */
function parseTime(text: string): Date | undefined {
    const res = text.match(timeRegex);
    if (!res) return;

    //  ["4:50:55 pm", "4", ":50", ":55", "pm", ...]
    const minutes = res[2] ? res[2] : ':00';
    const seconds = res[3] ? res[3] : ':00';
    const ampm = res[4] ? ' ' + res[4] : '';

    const date = new Date(Date.parse(`1/1/1900 ${res[1] + minutes + seconds + ampm} UTC`));
    let now = new Date();
    now = new Date(Date.UTC(now.getFullYear(), now.getMonth(), now.getDate(),
        now.getHours(), now.getMinutes(), now.getSeconds(), now.getMilliseconds()));

    return new Date(Date.UTC(now.getUTCFullYear(), now.getUTCMonth(), now.getUTCDate(),
        date.getUTCHours(), date.getUTCMinutes(), date.getUTCSeconds(), date.getUTCMilliseconds()));
}

/**
 * Parse a UTC date to excel serial number.
 * @param {Date|number} date - A UTC date.
 * @returns {number}
 */
function toSerial(date: Date | number): number {
    const dateTime = date instanceof Date ? date.getTime() : date;
    const addOn = (dateTime > -2203891200000) ? 2 : 1;
    return Math.floor((dateTime - d1900.getTime()) / 86400000) + addOn;
}

/**
 * Parse an excel serial number to UTC date.
 * @param serial
 * @returns {Date}
 */
function toDate(serial: number): Date {
    if (serial < 0) {
        throw FormulaError.VALUE;
    }
    if (serial <= 60) {
        return new Date(d1900.getTime() + (serial - 1) * 86400000);
    }
    return new Date(d1900.getTime() + (serial - 2) * 86400000);
}

function parseDateWithExtra(serialOrString: string | number | Date, nullDate: DateConfig | null = null): { date: Date; extra?: string } {
    if (serialOrString instanceof Date) return {date: serialOrString};
    const acceptedValue = H.accept(serialOrString);
    let date: Date;
    if (!isNaN(acceptedValue as number)) {
        const numValue = Number(acceptedValue);
        date = toDate(numValue);
    } else {
        // support time without date
        const timeDate = parseTime(acceptedValue as string);

        if (!timeDate) {
            date = parseSimplifiedDate(acceptedValue as string, nullDate);
        } else {
            date = timeDate;
        }
    }
    return {date};
}

function parseDate(serialOrString: string | number | Date): Date {
    return parseDateWithExtra(serialOrString).date;
}

function compareDateIgnoreTime(date1: Date, date2: Date): boolean {
    return date1.getUTCFullYear() === date2.getUTCFullYear() &&
        date1.getUTCMonth() === date2.getUTCMonth() &&
        date1.getUTCDate() === date2.getUTCDate();
}

function isLeapYear(year: number): boolean {
    if (year === 1900) {
        return true;
    }
    return new Date(year, 1, 29).getMonth() === 1;
}

const DateFunctions =  {
    // global configuration
    _config: {
        nullDate: { year: 1900, month: 1, day: 1 }
    },

    DATE: (year: number, month: number, day: number): number => {
        year = H.accept(year, Types.NUMBER);
        month = H.accept(month, Types.NUMBER);
        day = H.accept(day, Types.NUMBER);
        if (year < 0 || year >= 10000)
            throw FormulaError.NUM;

        // If year is between 0 (zero) and 1899 (inclusive), Excel adds that value to 1900 to calculate the year.
        if (year < 1900) {
            year += 1900;
        }

        return toSerial(Date.UTC(year, month - 1, day));
    },

    DATEDIF: (startDate: string | number | Date, endDate: string | number | Date, unit: string): number => {
        startDate = parseDate(startDate);
        endDate = parseDate(endDate);
        unit = H.accept(unit, Types.STRING).toLowerCase();

        if (startDate > endDate)
            throw FormulaError.NUM;
        const yearDiff = endDate.getUTCFullYear() - startDate.getUTCFullYear();
        const monthDiff = endDate.getUTCMonth() - startDate.getUTCMonth();
        const dayDiff = endDate.getUTCDate() - startDate.getUTCDate();
        let offset;
        switch (unit) {
            case 'y':
                offset = monthDiff < 0 || monthDiff === 0 && dayDiff < 0 ? -1 : 0;
                return offset + yearDiff;
            case 'm':
                offset = dayDiff < 0 ? -1 : 0;
                return yearDiff * 12 + monthDiff + offset;
            case 'd':
                return Math.floor((endDate.getTime() - startDate.getTime()) / MS_PER_DAY);
            case 'md':
                // The months and years of the dates are ignored.
                startDate.setUTCFullYear(endDate.getUTCFullYear());
                if (dayDiff < 0) {
                    startDate.setUTCMonth(endDate.getUTCMonth() - 1)
                } else {
                    startDate.setUTCMonth(endDate.getUTCMonth())
                }
                return Math.floor((endDate.getTime() - startDate.getTime()) / MS_PER_DAY);
            case 'ym':
                // The days and years of the dates are ignored
                offset = dayDiff < 0 ? -1 : 0;
                return (offset + yearDiff * 12 + monthDiff) % 12;
            case 'yd':
                // The years of the dates are ignored.
                if (monthDiff < 0 || monthDiff === 0 && dayDiff < 0) {
                    startDate.setUTCFullYear(endDate.getUTCFullYear() - 1);
                } else {
                    startDate.setUTCFullYear(endDate.getUTCFullYear());
                }
                return Math.floor((endDate.getTime() - startDate.getTime()) / MS_PER_DAY);
            default:
                throw FormulaError.VALUE;
        }
    },

    /**
     * Limitation: Year must be four digit, only support ISO 8016 date format.
     * Does not support date without year, i.e. "5-JUL".
     * @param {string} dateText
     */
    DATEVALUE: (dateText: string): number => {
        dateText = H.accept(dateText, Types.STRING);
        
        // Check if input is pure time (no date component)
        const timeDate = parseTime(dateText);
        if (timeDate) {
            // For pure time strings, DATEVALUE should return 0 (no date part)
            return 0;
        }
        
        const {date} = parseDateWithExtra(dateText, DateFunctions._config.nullDate);
        const serial = toSerial(date);
        if (serial < 0 || serial > 2958465)
            throw FormulaError.VALUE;
        return serial;
    },

    DAY: (serialOrString: string | number | Date): number => {
        const date = parseDate(serialOrString);
        return date.getUTCDate();
    },

    DAYS: (endDate: string | number | Date, startDate: string | number | Date): number => {
        endDate = parseDate(endDate);
        startDate = parseDate(startDate);
        let offset = 0;
        const startTime = startDate instanceof Date ? startDate.getTime() : startDate;
        const endTime = endDate instanceof Date ? endDate.getTime() : endDate;
        if (startTime < -2203891200000 && -2203891200000 < endTime) {
            offset = 1;
        }
        return Math.floor((endTime - startTime) / MS_PER_DAY) + offset;
    },

    DAYS360: (startDate: string | number | Date, endDate: string | number | Date, method?: boolean): number => {
        startDate = parseDate(startDate);
        endDate = parseDate(endDate);
        // default is US method
        method = H.accept(method, Types.BOOLEAN, false);

        if (startDate.getUTCDate() === 31) {
            startDate.setUTCDate(30);
        }
        if (!method && startDate.getUTCDate() < 30 && endDate.getUTCDate() > 30) {
            endDate.setUTCMonth(endDate.getUTCMonth() + 1, 1);
        } else {
            // European method
            if (endDate.getUTCDate() === 31) {
                endDate.setUTCDate(30);
            }
        }

        const yearDiff = endDate.getUTCFullYear() - startDate.getUTCFullYear();
        const monthDiff = endDate.getUTCMonth() - startDate.getUTCMonth();
        const dayDiff = endDate.getUTCDate() - startDate.getUTCDate();

        return (monthDiff) * 30 + dayDiff + yearDiff * 12 * 30;
    },

    EDATE: (startDate: string | number | Date, months: number): number => {
        startDate = parseDate(startDate);
        months = H.accept(months, Types.NUMBER);
        startDate.setUTCMonth(startDate.getUTCMonth() + months);
        return toSerial(startDate);
    },

    EOMONTH: (startDate: string | number | Date, months: number): number => {
        startDate = parseDate(startDate);
        months = H.accept(months, Types.NUMBER);
        startDate.setUTCMonth(startDate.getUTCMonth() + months + 1, 0);
        return toSerial(startDate);
    },

    HOUR: (serialOrString: string | number | Date): number => {
        const date = parseDate(serialOrString);
        return date.getUTCHours();
    },

    ISOWEEKNUM: (serialOrString: string | number | Date): number => {
        const date = parseDate(serialOrString);

        // https://stackoverflow.com/questions/6117814/get-week-of-year-in-javascript-like-in-php
        const d = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
        const dayNum = d.getUTCDay();
        d.setUTCDate(d.getUTCDate() + 4 - dayNum);
        const yearStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
        return Math.ceil((((d.getTime() - yearStart.getTime()) / 86400000) + 1) / 7)
    },

    MINUTE: (serialOrString: string | number | Date): number => {
        const date = parseDate(serialOrString);
        return date.getUTCMinutes();
    },

    MONTH: (serialOrString: string | number | Date): number => {
        const date = parseDate(serialOrString);
        return date.getUTCMonth() + 1;
    },

    NETWORKDAYS: (startDate: string | number | Date, endDate: string | number | Date, holidays?: any): number => {
        startDate = parseDate(startDate);
        endDate = parseDate(endDate);
        let sign = 1;
        if (startDate > endDate) {
            sign = -1;
            const temp = startDate;
            startDate = endDate;
            endDate = temp;
        }
        const holidaysArr: Date[] = [];
        if (holidays != null) {
            H.flattenParams([holidays], Types.NUMBER, false, (item: any) => {
                holidaysArr.push(parseDate(item));
            });
        }
        let numWorkDays = 0;
        while (startDate <= endDate) {
            // Skips Sunday and Saturday
            if (startDate.getUTCDay() !== 0 && startDate.getUTCDay() !== 6) {
                let found = false;
                for (let i = 0; i < holidaysArr.length; i++) {
                    if (compareDateIgnoreTime(startDate, holidaysArr[i])) {
                        found = true;
                        break;
                    }
                }
                if (!found) numWorkDays++;
            }
            startDate.setUTCDate(startDate.getUTCDate() + 1);
        }
        return sign * numWorkDays;

    },

    'NETWORKDAYS.INTL': (startDate: string | number | Date, endDate: string | number | Date, weekend?: string | number, holidays?: any): number => {
        startDate = parseDate(startDate);
        endDate = parseDate(endDate);
        let sign = 1;
        if (startDate > endDate) {
            sign = -1;
            const temp = startDate;
            startDate = endDate;
            endDate = temp;
        }
        weekend = H.accept(weekend, null, 1);
        // Using 1111111 will always return 0.
        if (weekend === '1111111')
            return 0;

        let weekendPattern: number[];
        // using weekend string, i.e, 0000011
        if (typeof weekend === "string" && Number(weekend).toString() !== weekend) {
            if (weekend.length !== 7) throw FormulaError.VALUE;
            const weekendStr = weekend.charAt(6) + weekend.slice(0, 6);
            const weekendArr: number[] = [];
            for (let i = 0; i < weekendStr.length; i++) {
                if (weekendStr.charAt(i) === '1')
                    weekendArr.push(i);
            }
            weekendPattern = weekendArr;
        } else {
            // using weekend number
            if (typeof weekend !== "number")
                throw FormulaError.VALUE;
            const pattern = WEEKEND_TYPES[weekend];
            if (!pattern) throw FormulaError.NUM;
            weekendPattern = pattern;
        }

        const holidaysArr: Date[] = [];
        if (holidays != null) {
            H.flattenParams([holidays], Types.NUMBER, false, (item: any) => {
                holidaysArr.push(parseDate(item));
            });
        }
        let numWorkDays = 0;
        while (startDate <= endDate) {
            let skip = false;
            for (let i = 0; i < weekendPattern.length; i++) {
                if (weekendPattern[i] === startDate.getUTCDay()) {
                    skip = true;
                    break;
                }
            }

            if (!skip) {
                let found = false;
                for (let i = 0; i < holidaysArr.length; i++) {
                    if (compareDateIgnoreTime(startDate, holidaysArr[i])) {
                        found = true;
                        break;
                    }
                }
                if (!found) numWorkDays++;
            }
            startDate.setUTCDate(startDate.getUTCDate() + 1);
        }
        return sign * numWorkDays;

    },

    NOW: (): number => {
        const now = new Date();
        return toSerial(Date.UTC(now.getFullYear(), now.getMonth(), now.getDate(),
            now.getHours(), now.getMinutes(), now.getSeconds(), now.getMilliseconds()))
            + (3600 * now.getHours() + 60 * now.getMinutes() + now.getSeconds()) / 86400;
    },

    SECOND: (serialOrString: string | number | Date): number => {
        const date = parseDate(serialOrString);
        return date.getUTCSeconds();
    },

    TIME: (hour: number, minute: number, second: number): number => {
        hour = H.accept(hour, Types.NUMBER);
        minute = H.accept(minute, Types.NUMBER);
        second = H.accept(second, Types.NUMBER);

        if (hour < 0 || hour > 32767 || minute < 0 || minute > 32767 || second < 0 || second > 32767)
            throw FormulaError.NUM;
        return (3600 * hour + 60 * minute + second) / 86400;
    },

    TIMEVALUE: (timeText: string | number | Date): number => {
        timeText = parseDate(timeText);
        return (3600 * timeText.getUTCHours() + 60 * timeText.getUTCMinutes() + timeText.getUTCSeconds()) / 86400;
    },

    TODAY: (): number => {
        const now = new Date();
        return toSerial(Date.UTC(now.getFullYear(), now.getMonth(), now.getDate()));
    },

    WEEKDAY: (serialOrString: string | number | Date, returnType?: number): number => {
        const date = parseDate(serialOrString);
        const returnTypeValue: number = H.accept(returnType, Types.NUMBER, 1);

        const day = date.getUTCDay();
        const weekTypes = WEEK_TYPES[returnTypeValue];
        if (!weekTypes)
            throw FormulaError.NUM;
        return weekTypes[day]!;

    },

    WEEKNUM: (serialOrString: string | number | Date, returnType?: number): number => {
        const date = parseDate(serialOrString);
        const returnTypeValue: number = H.accept(returnType, Types.NUMBER, 1);
        if (returnTypeValue === 21) {
            return DateFunctions.ISOWEEKNUM(serialOrString);
        }
        const weekStart = WEEK_STARTS[returnTypeValue]!;
        const yearStart = new Date(Date.UTC(date.getUTCFullYear(), 0, 1));
        const offset = yearStart.getUTCDay() < weekStart ? 1 : 0;
        return Math.ceil((((date.getTime() - yearStart.getTime()) / 86400000) + 1) / 7) + offset;
    },

    WORKDAY: (startDate: string | number | Date, days: number, holidays?: any): number => {
        return DateFunctions["WORKDAY.INTL"](startDate, days, 1, holidays);
    },

    'WORKDAY.INTL': (startDate: string | number | Date, days: number, weekend?: string | number, holidays?: any): number => {
        startDate = parseDate(startDate);
        days = H.accept(days, Types.NUMBER);

        weekend = H.accept(weekend, null, 1);
        // Using 1111111 will always return value error.
        if (weekend === '1111111')
            throw FormulaError.VALUE;

        let weekendPattern: number[];
        // using weekend string, i.e, 0000011
        if (typeof weekend === "string" && Number(weekend).toString() !== weekend) {
            if (weekend.length !== 7)
                throw FormulaError.VALUE;
            const weekendStr = weekend.charAt(6) + weekend.slice(0, 6);
            const weekendArr: number[] = [];
            for (let i = 0; i < weekendStr.length; i++) {
                if (weekendStr.charAt(i) === '1')
                    weekendArr.push(i);
            }
            weekendPattern = weekendArr;
        } else {
            // using weekend number
            if (typeof weekend !== "number")
                throw FormulaError.VALUE;
            const pattern = WEEKEND_TYPES[weekend];
            if (!pattern)
                throw FormulaError.NUM;
            weekendPattern = pattern;
        }

        const holidaysArr: Date[] = [];
        if (holidays != null) {
            H.flattenParams([holidays], Types.NUMBER, false, (item: any) => {
                holidaysArr.push(parseDate(item));
            });
        }
        startDate.setUTCDate(startDate.getUTCDate() + 1);
        let cnt = 0;
        while (cnt < days) {
            let skip = false;
            for (let i = 0; i < weekendPattern.length; i++) {
                if (weekendPattern[i] === startDate.getUTCDay()) {
                    skip = true;
                    break;
                }
            }

            if (!skip) {
                let found = false;
                for (let i = 0; i < holidaysArr.length; i++) {
                    if (compareDateIgnoreTime(startDate, holidaysArr[i])) {
                        found = true;
                        break;
                    }
                }
                if (!found) cnt++;
            }
            startDate.setUTCDate(startDate.getUTCDate() + 1);
        }
        return toSerial(startDate) - 1;
    },

    YEAR: (serialOrString: string | number | Date): number => {
        const date = parseDate(serialOrString);
        return date.getUTCFullYear();
    },

    // Warning: may have bugs
    YEARFRAC: (startDate: string | number | Date, endDate: string | number | Date, basis?: number): number => {
        startDate = parseDate(startDate);
        endDate = parseDate(endDate);
        if (startDate > endDate) {
            const temp = startDate;
            startDate = endDate;
            endDate = temp;
        }
        const basisValue: number = H.accept(basis, Types.NUMBER, 0);
        const basisInt = Math.trunc(basisValue);

        if (basisInt < 0 || basisInt > 4)
            throw FormulaError.VALUE;

        // https://github.com/LesterLyu/formula.js/blob/develop/lib/date-time.js#L508
        let sd = startDate.getUTCDate();
        const sm = startDate.getUTCMonth() + 1;
        const sy = startDate.getUTCFullYear();
        let ed = endDate.getUTCDate();
        const em = endDate.getUTCMonth() + 1;
        const ey = endDate.getUTCFullYear();

        switch (basisInt) {
            case 0:
                // US (NASD) 30/360
                if (sd === 31 && ed === 31) {
                    sd = 30;
                    ed = 30;
                } else if (sd === 31) {
                    sd = 30;
                } else if (sd === 30 && ed === 31) {
                    ed = 30;
                }
                return Math.abs((ed + em * 30 + ey * 360) - (sd + sm * 30 + sy * 360)) / 360;
            case 1:
                // Actual/actual
                if (ey - sy < 2) {
                    const yLength = isLeapYear(sy) && sy !== 1900 ? 366 : 365;
                    const days = DateFunctions.DAYS(endDate, startDate);
                    return days / yLength;
                } else {
                    const years = (ey - sy) + 1;
                    const days = (new Date(ey + 1, 0, 1).getTime() - new Date(sy, 0, 1).getTime()) / 1000 / 60 / 60 / 24;
                    const average = days / years;
                    return DateFunctions.DAYS(endDate, startDate) / average;
                }
            case 2:
                // Actual/360
                return Math.abs(DateFunctions.DAYS(endDate, startDate) / 360);
            case 3:
                // Actual/365
                return Math.abs(DateFunctions.DAYS(endDate, startDate) / 365);
            case 4:
                // European 30/360
                return Math.abs((ed + em * 30 + ey * 360) - (sd + sm * 30 + sy * 360)) / 360;
            default:
                return 0;
        }
    },
};

export default DateFunctions;
