const FormulaError = require('../error');
const {FormulaHelpers, Types} = require('../helpers');
const H = FormulaHelpers;

const WebFunctions = {
    ENCODEURL: text => {
        return encodeURIComponent(H.accept(text, Types.STRING));
    },

    FILTERXML: () => {
        // Not implemented due to extra dependency
        throw FormulaError.NOT_IMPLEMENTED('FILTERXML');
    },

    WEBSERVICE: (context, url) => {
        if (typeof fetch === "function") {
            url = H.accept(url, Types.STRING);
            // validate url
            try {
                new URL(url);
            } catch (err) {
                throw FormulaError.ERROR('WEBSERVICE request failed: Invalid URL format');
            }
            return fetch(url).then(res => res.text()).catch(err => {
                throw FormulaError.ERROR('WEBSERVICE request failed: ' + err.message);
            });
        } else {
            // Not implemented for Node.js due to extra dependency
            // Sample code for Node.js
            // const fetch = require('node-fetch');
            // url = H.accept(url, Types.STRING);
            // return fetch(url).then(res => res.text());
            throw FormulaError.ERROR('WEBSERVICE only available to browser with fetch. ' +
                'If you want to use WEBSERVICE in Node.js, please override this function: \n' +
                'new FormulaParser({\n' +
                '    functionsNeedContext: {\n' +
                '        WEBSERVICE: (context, url) => {...}}\n' +
                '})');
        }
    }
}

module.exports = WebFunctions;
