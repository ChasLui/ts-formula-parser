// Type definitions for jstat
declare module 'jstat' {
  interface JStat {
    beta: {
      pdf(x: number, alpha: number, beta: number): number;
      cdf(x: number, alpha: number, beta: number): number;
      inv(p: number, alpha: number, beta: number): number;
    };
    binomial: {
      pdf(k: number, n: number, p: number): number;
      cdf(k: number, n: number, p: number): number;
    };
    centralF: {
      pdf(x: number, df1: number, df2: number): number;
      cdf(x: number, df1: number, df2: number): number;
      inv(p: number, df1: number, df2: number): number;
    };
    chisquare: {
      pdf(x: number, dof: number): number;
      cdf(x: number, dof: number): number;
      inv(p: number, dof: number): number;
    };
    exponential: {
      pdf(x: number, rate: number): number;
      cdf(x: number, rate: number): number;
    };
    gamma: {
      pdf(x: number, shape: number, scale: number): number;
      cdf(x: number, shape: number, scale: number): number;
      inv(p: number, shape: number, scale: number): number;
    };
    hypgeom: {
      pdf(k: number, N: number, m: number, n: number): number;
      cdf(k: number, N: number, m: number, n: number): number;
    };
    lognormal: {
      pdf(x: number, mu: number, sigma: number): number;
      cdf(x: number, mu: number, sigma: number): number;
      inv(p: number, mu: number, sigma: number): number;
    };
    negbin: {
      pdf(k: number, r: number, p: number): number;
      cdf(k: number, r: number, p: number): number;
    };
    normal: {
      pdf(x: number, mean: number, std: number): number;
      cdf(x: number, mean: number, std: number): number;
      inv(p: number, mean: number, std: number): number;
    };
    poisson: {
      pdf(k: number, lambda: number): number;
      cdf(k: number, lambda: number): number;
    };
    studentt: {
      pdf(x: number, dof: number): number;
      cdf(x: number, dof: number): number;
      inv(p: number, dof: number): number;
    };
    weibull: {
      pdf(x: number, scale: number, shape: number): number;
      cdf(x: number, scale: number, shape: number): number;
    };
    
    // Utility functions
    factorial(n: number): number;
    factorial2(n: number): number;
    factorialln(n: number): number;
    gammaFunc(x: number): number;
    gammaln(x: number): number;
    combinantion(n: number, k: number): number;
    permutation(n: number, k: number): number;
    
    // Matrix and array functions
    transpose(matrix: number[][]): number[][];
    multiply(a: number[][], b: number[][]): number[][];
    
    // Statistical functions
    mean(array: number[]): number;
    median(array: number[]): number;
    mode(array: number[]): number | number[];
    variance(array: number[]): number;
    stdev(array: number[]): number;
    covariance(array1: number[], array2: number[]): number;
    corrcoeff(array1: number[], array2: number[]): number;
  }

  const jStat: JStat;
  export = jStat;
}