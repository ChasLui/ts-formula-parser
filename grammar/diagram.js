import fs from 'fs';
import * as chevrotain from 'chevrotain';
import {Parser} from './parsing.js';

/**
 * Class for generating syntax diagrams, used to create visual diagrams of Excel formula grammar
 */
export class DiagramGenerator {
    constructor() {
        this.parserInstance = new Parser();
    }

    /**
     * Get serialized grammar definition
     * @returns {Array} Serialized grammar productions
     */
    getSerializedGrammar() {
        return this.parserInstance.getSerializedGastProductions();
    }

    /**
     * Generate HTML format syntax diagram code
     * @returns {string} HTML code string
     */
    generateHTMLCode() {
        const serializedGrammar = this.getSerializedGrammar();
        return chevrotain.createSyntaxDiagramsCode(serializedGrammar);
    }

    /**
     * Write syntax diagram to specified file
     * @param {string} filePath - Output file path
     * @param {boolean} silent - Whether to handle errors silently (no output to console)
     * @returns {boolean|Error} Returns true on success, false or Error object on failure
     */
    writeToFile(filePath, silent = false) {
        try {
            const htmlText = this.generateHTMLCode();
            fs.writeFileSync(filePath, htmlText);
            return true;
        } catch (error) {
            if (!silent) {
                console.error('Failed to write diagram file:', error);
            }
            return false;
        }
    }

    /**
     * Validate the completeness of generated grammar data
     * @returns {Object} Validation result
     */
    validateGrammar() {
        const grammar = this.getSerializedGrammar();
        const result = {
            isValid: true,
            ruleCount: 0,
            errors: []
        };

        if (!Array.isArray(grammar)) {
            result.isValid = false;
            result.errors.push('Grammar is not an array');
            return result;
        }

        result.ruleCount = grammar.length;

        // Verify each rule has required properties
        for (let i = 0; i < grammar.length; i++) {
            const rule = grammar[i];
            if (!rule.type || !rule.name) {
                result.isValid = false;
                result.errors.push(`Rule at index ${i} is missing type or name`);
            }
        }

        return result;
    }

    /**
     * Get statistical information of grammar rules
     * @returns {Object} Statistical information
     */
    getGrammarStats() {
        const grammar = this.getSerializedGrammar();
        const stats = {
            totalRules: grammar.length,
            ruleNames: [],
            ruleTypes: new Set()
        };

        grammar.forEach(rule => {
            stats.ruleNames.push(rule.name);
            stats.ruleTypes.add(rule.type);
        });

        return {
            ...stats,
            ruleTypes: Array.from(stats.ruleTypes)
        };
    }
}

// If this file is run directly, execute default diagram generation
if (import.meta.url === `file://${process.argv[1]}`) {
    const generator = new DiagramGenerator();
    const success = generator.writeToFile("./docs/generated_diagrams.html");
    
    if (success) {
        console.log('Syntax diagrams generated successfully!');
        const stats = generator.getGrammarStats();
        console.log(`Generated ${stats.totalRules} grammar rules`);
    } else {
        console.error('Failed to generate syntax diagrams');
        process.exit(1);
    }
}

export default DiagramGenerator;
