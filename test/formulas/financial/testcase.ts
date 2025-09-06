import FormulaError from "../../../formulas/error";

export default {
    ACCRINT: {
        // Basic test cases: issue date, first interest date, settlement date, rate, par value, frequency, basis, calculation method
        // ACCRINT(issue, first_interest, settlement, rate, par, frequency, basis, calc_method)
        
        // Basic functionality tests - using correct calculation results
        'ACCRINT(41640,41731,41823,0.1,1000,1,0)': 50.55555555555556, // Basic calculation
        'ACCRINT(41640,41731,41823,0.05,1000,1,0)': 25.27777777777778, // Low interest rate 
        'ACCRINT(41640,41731,41823,0.1,500,1,0)': 25.27777777777778, // Low par value
        'ACCRINT(41640,41731,41823,0.1,1000,2,0)': 50.55555555555556, // Semi-annual payment
        'ACCRINT(41640,41731,41823,0.1,1000,4,0)': 50.55555555555556, // Quarterly payment
        
        // Different basis tests (date calculation methods)
        'ACCRINT(41640,41731,41823,0.1,1000,1,1)': 50.136986301369866, // Basis 1: actual/actual  
        'ACCRINT(41640,41731,41823,0.1,1000,1,2)': 50.83333333333333, // Basis 2: actual/360
        'ACCRINT(41640,41731,41823,0.1,1000,1,3)': 50.136986301369866, // Basis 3: actual/365
        'ACCRINT(41640,41731,41823,0.1,1000,1,4)': 50.55555555555556, // Basis 4: 30/360
        
        // calc_method tests
        'ACCRINT(41640,41731,41823,0.1,1000,1,0,TRUE)': 50.55555555555556,  // Calculate from issue date
        'ACCRINT(41640,41731,41823,0.1,1000,1,0,FALSE)': 25.27777777777778, // Calculate from first interest date
        
        // Using date strings
        'ACCRINT("1/1/2014","4/1/2014","7/1/2014",0.1,1000,1,0)': 50,
        'ACCRINT("2014-01-01","2014-04-01","2014-07-01",0.1,1000,1,0)': 50,
        
        // Boundary value tests - error cases
        'ACCRINT(41640,41731,41823,-0.1,1000,1,0)': FormulaError.NUM, // Negative interest rate
        'ACCRINT(41640,41731,41823,0.1,-1000,1,0)': FormulaError.NUM, // Negative par value
        'ACCRINT(41640,41731,41823,0.1,1000,3,0)': FormulaError.NUM, // Invalid frequency
        'ACCRINT(41640,41731,41823,0.1,1000,1,-1)': FormulaError.NUM, // Invalid basis
        'ACCRINT(41640,41731,41823,0.1,1000,1,5)': FormulaError.NUM, // Invalid basis
        'ACCRINT(41823,41731,41640,0.1,1000,1,0)': FormulaError.NUM, // Issue date later than settlement date
        
        // Zero value tests
        'ACCRINT(41640,41731,41823,0,1000,1,0)': FormulaError.NUM, // Zero interest rate
        'ACCRINT(41640,41731,41823,0.1,0,1,0)': FormulaError.NUM, // Zero par value
        
        // Very small value tests
        'ACCRINT(41640,41731,41823,0.001,1000,1,0)': 0.5055555555555555, // Very small interest rate
        'ACCRINT(41640,41731,41823,0.1,1,1,0)': 0.050555555555555555, // Very small par value
        
        // Same date tests (minimum time difference)
        'ACCRINT(41640,41640,41641,0.1,1000,1,0)': 0.2777777777777778, // Accrued interest for 1 day
    }
};