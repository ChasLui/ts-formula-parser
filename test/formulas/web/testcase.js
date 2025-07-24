const FormulaError = require('../../../formulas/error');

module.exports = {
    ENCODEURL: {
        'ENCODEURL("http://contoso.sharepoint.com/teams/Finance/Documents/April Reports/Profit and Loss Statement.xlsx")':
            'http%3A%2F%2Fcontoso.sharepoint.com%2Fteams%2FFinance%2FDocuments%2FApril%20Reports%2FProfit%20and%20Loss%20Statement.xlsx',
        'ENCODEURL(123)': '123',
        'ENCODEURL(TRUE)': 'TRUE',
        'ENCODEURL(FALSE)': 'FALSE',
        'ENCODEURL()': FormulaError.ARG_MISSING(['STRING']),
        'ENCODEURL(null)': FormulaError.NAME,
        'ENCODEURL({1,2,3})': '1', // Take the first element of the array
    },
    FILTERXML: {
        'FILTERXML("<root><a>1</a></root>", "//a")': FormulaError.NOT_IMPLEMENTED('FILTERXML'), // 未实现，抛出 NOT_IMPLEMENTED 错误
    },
    WEBSERVICE: {
        // test error handling for invalid URL (fetch exists in Node.js environment)
        'WEBSERVICE("invalid-url")': FormulaError.ERROR(),
        'WEBSERVICE()': FormulaError.ARG_MISSING(['STRING']),
        'WEBSERVICE(null)': FormulaError.NAME,
    }
};
