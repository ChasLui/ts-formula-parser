#!/usr/bin/env node

/**
 * Performance test for the formula parser
 */

import { performance } from 'perf_hooks';
import { FormulaParser } from '../build/index.mjs';

const parser = new FormulaParser();

// Test formulas with varying complexity
const testCases = [
  // Simple arithmetic
  '=1+2*3',
  '=(10+5)*2',
  
  // Mathematical functions
  '=SUM(1,2,3,4,5)',
  '=AVERAGE(10,20,30,40,50)',
  '=SQRT(16)+POWER(2,3)',
  
  // Complex formulas
  '=IF(SUM(A1:A10)>100,AVERAGE(B1:B10)*1.2,AVERAGE(B1:B10)*0.8)',
  '=VLOOKUP("test",A1:C100,2,FALSE)',
  
  // Date functions
  '=DATE(2023,12,25)',
  '=DATEDIF(DATE(2023,1,1),TODAY(),"D")',
  
  // Statistical functions
  '=NORM.DIST(0,0,1,TRUE)',
  '=BETA.DIST(0.5,2,3,TRUE)',
];

function runPerformanceTest() {
  console.log('🚀 Formula Parser Performance Test');
  console.log('=====================================\n');
  
  const results = [];
  
  for (const formula of testCases) {
    const iterations = 1000;
    const times = [];
    
    // Warmup
    for (let i = 0; i < 10; i++) {
      try {
        parser.parse(formula);
      } catch (e) {
        // Ignore parsing errors for performance test
      }
    }
    
    // Measure
    for (let i = 0; i < iterations; i++) {
      const start = performance.now();
      try {
        parser.parse(formula);
      } catch (e) {
        // Ignore parsing errors for performance test
      }
      const end = performance.now();
      times.push(end - start);
    }
    
    const avgTime = times.reduce((a, b) => a + b, 0) / times.length;
    const minTime = Math.min(...times);
    const maxTime = Math.max(...times);
    
    results.push({
      formula,
      avgTime: avgTime.toFixed(3),
      minTime: minTime.toFixed(3),
      maxTime: maxTime.toFixed(3),
      opsPerSec: Math.round(1000 / avgTime)
    });
  }
  
  // Display results
  console.log('Formula'.padEnd(60) + 'Avg(ms)'.padEnd(10) + 'Min(ms)'.padEnd(10) + 'Max(ms)'.padEnd(10) + 'Ops/sec');
  console.log('-'.repeat(100));
  
  for (const result of results) {
    console.log(
      result.formula.padEnd(60) +
      result.avgTime.padEnd(10) +
      result.minTime.padEnd(10) +
      result.maxTime.padEnd(10) +
      result.opsPerSec.toString()
    );
  }
  
  const totalAvg = results.reduce((sum, r) => sum + parseFloat(r.avgTime), 0) / results.length;
  console.log(`\n📊 Overall average: ${totalAvg.toFixed(3)}ms per parse`);
  console.log(`🏃 Overall throughput: ${Math.round(1000 / totalAvg)} operations/second`);
}

runPerformanceTest();