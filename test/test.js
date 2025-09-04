import assert from 'assert';
import { expect, describe, it, beforeAll } from 'vitest';
import {FormulaParser} from '../grammar/hooks.js';
import fs from 'fs';
import readline from 'readline';
import './grammar/test.js';
import './grammar/errors.js';
import './grammar/collection.js';
import './grammar/depParser.js';
import './formulas/index.js';

const parser = new FormulaParser(undefined, true);

describe.skip('Parsing Formulas 1', function () {
    let success = 0;
    const formulas = [];
    const failures = [];
    beforeAll(async () => {
        return new Promise((resolve, reject) => {
            const lineReader = readline.createInterface({
                input: fs.createReadStream('./test/formulas.txt'),
                crlfDelay: Infinity
            });
            
            lineReader.on('line', (line) => {
                line = line.slice(1, -1)
                    .replace(/""/g, '"');
                if (line.indexOf('[') === -1)
                    formulas.push(line);
                // else
                //     console.log(`not supported: ${line}`)
                // console.log(line)
            });
            lineReader.on('close', () => {
                lineReader.close();
                resolve();
            });
            lineReader.on('error', (err) => {
                console.error('Error reading formulas.txt:', err);
                reject(err);
            });
        });
    });

    it('formulas parse rate should be 100%', function () {
        // console.log(formulas.length);
        formulas.forEach((formula, index) => {
            // console.log('testing #', index, formula);
            try {
                parser.parse(formula, {row: 2, col: 2});
                success++;
            } catch (e) {
                failures.push(formula);
            }
        });
        if (failures.length > 0) {
            console.log(failures);
            console.log(`Success rate: ${success / formulas.length * 100}%`);
        }

        const logs = parser.logs.sort();
        console.log(`The following functions is not implemented: (${logs.length} in total)\n ${logs.join(', ')}`);
        parser.logs = [];
        assert.strictEqual(success / formulas.length === 1, true);
    });
});

describe('Parsing Formulas 2', () => {
    let success = 0;
    let formulas;
    const failures = [];
    beforeAll(async () => {
        return new Promise((resolve, reject) => {
            fs.readFile('./test/formulas2.json', (err, data) => {
                if (err) {
                    reject(err);
                    return;
                }
                formulas = JSON.parse(data.toString());
                resolve();
            });
        });
    });

    it ('skip', () => '');

    it('custom formulas parse rate should be 100%',  function() {
        formulas.forEach((formula, index)  => {
            // console.log('testing #', index, formula);
            try {
                parser.parse(formula);
                success++;
            } catch (e) {
                failures.push(formula);
            }
        });
        if (failures.length > 0) {
            console.log(failures);
            console.log(`Success rate: ${success / formulas.length * 100}%`);
        }
        assert.strictEqual(success / formulas.length === 1, true);
        const logs = parser.logs.sort();
        console.log(`The following functions is not implemented: (${logs.length} in total)\n ${logs.join(', ')}`);
        parser.logs = [];
    });

});

describe('Get supported formulas', () => {
    it('should support more than 275 functions', () => {
        const functionsNames =  parser.supportedFunctions();
        expect(functionsNames.length).to.greaterThan(275);
        console.log(`Support ${functionsNames.length} functions:\n${functionsNames.join(', ')}`);
        expect(functionsNames.includes('IFNA')).to.eq(true);
        expect(functionsNames.includes('SUMIF')).to.eq(true);
    });
})
