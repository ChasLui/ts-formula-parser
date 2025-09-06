import fs from 'fs';
import * as chevrotain from 'chevrotain';
import { Parser } from './parsing';

// Define interfaces for grammar and validation results
interface ValidationResult {
    isValid: boolean;
    ruleCount: number;
    errors: string[];
}

interface GrammarStats {
    totalRules: number;
    ruleNames: string[];
    ruleTypes: string[];
}

/**
 * Class for generating syntax diagrams, used to create visual diagrams of Excel formula grammar
 */
export class DiagramGenerator {
    private parserInstance: Parser;

    constructor() {
        // Create parser with dummy context for diagram generation
        const dummyContext = {
            callFunction: () => {},
            getVariable: () => {}
        };
        const dummyUtils = {} as any;
        this.parserInstance = new Parser(dummyContext, dummyUtils);
    }

    /**
     * Get serialized grammar definition
     * @returns Serialized grammar productions
     */
    getSerializedGrammar(): any[] {
        return (this.parserInstance as any).getSerializedGastProductions();
    }

    /**
     * Generate HTML format syntax diagram code
     * @returns HTML code string
     */
    generateHTMLCode(): string {
        const serializedGrammar = this.getSerializedGrammar();
        return (chevrotain as any).createSyntaxDiagramsCode(serializedGrammar);
    }

    /**
     * Write syntax diagram to specified file
     * @param filePath - Output file path
     * @param silent - Whether to handle errors silently (no output to console)
     * @returns Returns true on success, false on failure
     */
    writeToFile(filePath: string, silent: boolean = false): boolean {
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
     * @returns Validation result
     */
    validateGrammar(): ValidationResult {
        const grammar = this.getSerializedGrammar();
        const result: ValidationResult = {
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
     * @returns Statistical information
     */
    getGrammarStats(): GrammarStats {
        const grammar = this.getSerializedGrammar();
        const stats = {
            totalRules: grammar.length,
            ruleNames: [] as string[],
            ruleTypes: new Set<string>()
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