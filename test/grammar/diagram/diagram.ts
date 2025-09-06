import { describe, it, expect, beforeAll, afterAll } from 'vitest';
import fs from 'fs';
import path from 'path';
import { DiagramGenerator } from '../../../grammar/diagram';

interface GrammarRule {
    type: string;
    name: string;
    definition?: any[];
}

interface GrammarValidation {
    isValid: boolean;
    ruleCount: number;
    errors: string[];
}

interface GrammarStats {
    totalRules: number;
    ruleNames: string[];
    ruleTypes: string[];
}

describe('DiagramGenerator', () => {
    let generator: DiagramGenerator;
    let testOutputPath: string;

    beforeAll(() => {
        generator = new DiagramGenerator();
        // Use temporary path to avoid affecting actual documentation files
        testOutputPath = path.join(process.cwd(), 'test', 'grammar', 'diagram', 'test_output.html');
    });

    afterAll(() => {
        // Clean up test files
        if (fs.existsSync(testOutputPath)) {
            fs.unlinkSync(testOutputPath);
        }
    });

    describe('Instantiation and basic functionality', () => {
        it('should correctly instantiate DiagramGenerator', () => {
            expect(generator).toBeDefined();
            expect(generator.parserInstance).toBeDefined();
        });

        it('should be able to get serialized grammar', () => {
            const grammar: GrammarRule[] = generator.getSerializedGrammar();
            
            expect(Array.isArray(grammar)).toBe(true);
            expect(grammar.length).toBeGreaterThan(0);
            
            // Verify the structure of the first rule
            const firstRule: GrammarRule = grammar[0];
            expect(firstRule).toHaveProperty('type');
            expect(firstRule).toHaveProperty('name');
            expect(firstRule).toHaveProperty('definition');
        });
    });

    describe('Grammar validation functionality', () => {
        it('should be able to validate grammar completeness', () => {
            const validation: GrammarValidation = generator.validateGrammar();
            
            expect(validation).toHaveProperty('isValid');
            expect(validation).toHaveProperty('ruleCount');
            expect(validation).toHaveProperty('errors');
            
            expect(validation.isValid).toBe(true);
            expect(validation.ruleCount).toBeGreaterThan(0);
            expect(validation.errors).toHaveLength(0);
        });

        it('should be able to handle invalid grammar data', () => {
            // Create a test instance with mock grammar
            const testGenerator = new DiagramGenerator();
            
            // Temporarily replace getSerializedGrammar method to test error cases
            const originalMethod = testGenerator.getSerializedGrammar;
            
            // Test non-array grammar data
            testGenerator.getSerializedGrammar = () => null as any;
            let validation: GrammarValidation = testGenerator.validateGrammar();
            expect(validation.isValid).toBe(false);
            expect(validation.errors).toContain('Grammar is not an array');
            
            // Test rules missing required properties
            testGenerator.getSerializedGrammar = () => [
                { type: 'Rule', name: 'validRule' },
                { type: 'Rule' }, // Missing name
                { name: 'noType' } // Missing type
            ] as any;
            validation = testGenerator.validateGrammar();
            expect(validation.isValid).toBe(false);
            expect(validation.errors.length).toBeGreaterThan(0);
            expect(validation.ruleCount).toBe(3);
            
            // Restore original method
            testGenerator.getSerializedGrammar = originalMethod;
        });

        it('should return correct grammar statistics', () => {
            const stats: GrammarStats = generator.getGrammarStats();
            
            expect(stats).toHaveProperty('totalRules');
            expect(stats).toHaveProperty('ruleNames');
            expect(stats).toHaveProperty('ruleTypes');
            
            expect(stats.totalRules).toBeGreaterThan(0);
            expect(Array.isArray(stats.ruleNames)).toBe(true);
            expect(Array.isArray(stats.ruleTypes)).toBe(true);
            
            // Verify expected grammar rules exist
            const expectedRules = [
                'formulaWithBinaryOp',
                'formula',
                'functionCall',
                'constant',
                'referenceItem'
            ];
            
            expectedRules.forEach((ruleName: string) => {
                expect(stats.ruleNames).toContain(ruleName);
            });
            
            // Verify grammar rule types
            expect(stats.ruleTypes).toContain('Rule');
        });
    });

    describe('HTML generation functionality', () => {
        it('should be able to generate HTML code', () => {
            const htmlCode: string = generator.generateHTMLCode();
            
            expect(typeof htmlCode).toBe('string');
            expect(htmlCode.length).toBeGreaterThan(0);
            
            // Verify generated HTML contains necessary elements
            expect(htmlCode).toContain('<!DOCTYPE html>');
            expect(htmlCode).toContain('<div id="diagrams"');
            expect(htmlCode).toContain('window.serializedGrammar');
            expect(htmlCode).toContain('chevrotain');
        });

        it('generated HTML should contain all grammar rules', () => {
            const htmlCode: string = generator.generateHTMLCode();
            const stats: GrammarStats = generator.getGrammarStats();
            
            // Check if main grammar rules are in HTML
            const criticalRules: string[] = ['formulaWithBinaryOp', 'formula', 'functionCall'];
            criticalRules.forEach((ruleName: string) => {
                expect(htmlCode).toContain(ruleName);
            });
        });

        it('generated HTML should contain correct Chevrotain resource links', () => {
            const htmlCode: string = generator.generateHTMLCode();
            
            expect(htmlCode).toContain('chevrotain@');
            expect(htmlCode).toContain('diagrams.css');
            expect(htmlCode).toContain('railroad-diagrams.js');
            expect(htmlCode).toContain('diagrams_builder.js');
        });
    });

    describe('File writing functionality', () => {
        it('should be able to successfully write HTML file', () => {
            const success: boolean = generator.writeToFile(testOutputPath);
            
            expect(success).toBe(true);
            expect(fs.existsSync(testOutputPath)).toBe(true);
            
            // Verify file content
            const fileContent: string = fs.readFileSync(testOutputPath, 'utf8');
            expect(fileContent.length).toBeGreaterThan(0);
            expect(fileContent).toContain('<!DOCTYPE html>');
            expect(fileContent).toContain('window.serializedGrammar');
        });

        it('written file should contain complete grammar data', () => {
            generator.writeToFile(testOutputPath);
            const fileContent: string = fs.readFileSync(testOutputPath, 'utf8');
            const stats: GrammarStats = generator.getGrammarStats();
            
            // Verify file contains all rule names
            stats.ruleNames.forEach((ruleName: string) => {
                expect(fileContent).toContain(ruleName);
            });
        });

        it('should be able to handle invalid file paths', () => {
            const invalidPath = '/invalid/path/that/does/not/exist/test.html';
            // Use silent: true parameter to avoid stderr output
            const success: boolean = generator.writeToFile(invalidPath, true);
            
            expect(success).toBe(false);
        });

        it('should be able to handle errors in normal mode (by simulating fs error)', () => {
            // Simulate fs.writeFileSync throwing an error
            const originalWriteFileSync = fs.writeFileSync;
            let consoleErrorCalled = false;
            const originalConsoleError = console.error;
            
            // Temporarily replace console.error to capture calls
            console.error = () => { consoleErrorCalled = true; };
            
            try {
                fs.writeFileSync = () => {
                    throw new Error('Simulated file write error');
                };
                
                const result: boolean = generator.writeToFile('/some/path/test.html', false);
                expect(result).toBe(false);
                expect(consoleErrorCalled).toBe(true);
            } finally {
                // Restore original functions
                fs.writeFileSync = originalWriteFileSync;
                console.error = originalConsoleError;
            }
        });
    });

    describe('Specific validation of grammar rules', () => {
        it('should contain all necessary Excel formula grammar rules', () => {
            const grammar: GrammarRule[] = generator.getSerializedGrammar();
            const ruleNames: string[] = grammar.map((rule: GrammarRule) => rule.name);
            
            // Core rules that Excel formula parser should contain
            const requiredRules: string[] = [
                'formulaWithBinaryOp',      // Binary operator formula
                'formulaWithPercentOp',     // Percent operator
                'formulaWithUnaryOp',       // Unary operator
                'formulaWithIntersect',     // Intersection operation
                'formulaWithRange',         // Range operation
                'formula',                  // Basic formula
                'paren',                   // Parenthesized expression
                'constantArray',           // Constant array
                'constant',                // Constant
                'functionCall',            // Function call
                'arguments',               // Function arguments
                'referenceWithoutInfix',   // Reference (without infix)
                'referenceItem',           // Reference item
                'prefixName'               // Prefix name
            ];
            
            requiredRules.forEach((ruleName: string) => {
                expect(ruleNames).toContain(ruleName);
            });
        });

        it('grammar rules should have correct structure', () => {
            const grammar: GrammarRule[] = generator.getSerializedGrammar();
            
            grammar.forEach((rule: GrammarRule, index: number) => {
                // Each rule should have these basic properties
                expect(rule).toHaveProperty('type');
                expect(rule).toHaveProperty('name');
                expect(rule).toHaveProperty('definition');
                
                // Verify type should be 'Rule'
                expect(rule.type).toBe('Rule');
                
                // Name should be non-empty string
                expect(typeof rule.name).toBe('string');
                expect(rule.name.length).toBeGreaterThan(0);
                
                // definition should be array
                expect(Array.isArray(rule.definition)).toBe(true);
            });
        });

        it('should contain operator-related rules', () => {
            const grammar: GrammarRule[] = generator.getSerializedGrammar();
            const allContent: string = JSON.stringify(grammar);
            
            // Verify contains various operators
            const operators: string[] = ['+', '-', '*', '/', '^', '=', '>', '<', '>=', '<=', '<>', '&', '%'];
            operators.forEach((op: string) => {
                expect(allContent).toContain(op);
            });
        });
    });

    describe('Performance and robustness tests', () => {
        it('multiple calls should return consistent results', () => {
            const grammar1: GrammarRule[] = generator.getSerializedGrammar();
            const grammar2: GrammarRule[] = generator.getSerializedGrammar();
            
            expect(grammar1).toEqual(grammar2);
        });

        it('HTML generation should be deterministic', () => {
            const html1: string = generator.generateHTMLCode();
            const html2: string = generator.generateHTMLCode();
            
            expect(html1).toBe(html2);
        });

        it('grammar validation should be idempotent', () => {
            const validation1: GrammarValidation = generator.validateGrammar();
            const validation2: GrammarValidation = generator.validateGrammar();
            
            expect(validation1).toEqual(validation2);
        });
    });
});