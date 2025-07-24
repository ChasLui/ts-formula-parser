const path = require("path");
const BundleAnalyzerPlugin = require('webpack-bundle-analyzer').BundleAnalyzerPlugin;

// 是否启用bundle分析器
const enableAnalyzer = process.env.ANALYZE === 'true';

// 基础配置 - 非压缩版本
const baseConfig = {
    mode: "production",
    entry: "./index.js",
    resolve: {
        extensions: ['.js']
    },
    plugins: [],
    optimization: {
        minimize: false  // 非压缩版本
    }
};

// 基础配置 - 压缩版本
const baseMinConfig = {
    mode: "production",
    entry: "./index.js",
    resolve: {
        extensions: ['.js']
    },
    plugins: [],
    optimization: {
        minimize: true  // 压缩版本
    }
};

// UMD 格式配置
const umdConfig = {
    ...baseConfig,
    output: {
        path: path.resolve(__dirname, "./build/"),
        filename: "parser.umd.js",
        library: "FormulaParser",
        libraryTarget: "umd",
        globalObject: 'this'
    },
    plugins: [
        ...(enableAnalyzer ? [new BundleAnalyzerPlugin({
            analyzerMode: 'static',
            openAnalyzer: true,
            generateStatsFile: true,
            reportFilename: 'bundle-report.html'
        })] : [])
    ]
};

// UMD 压缩版本
const umdMinConfig = {
    ...baseMinConfig,
    output: {
        path: path.resolve(__dirname, "./build/"),
        filename: "parser.umd.min.js",
        library: "FormulaParser",
        libraryTarget: "umd",
        globalObject: 'this'
    }
};

// ESM 格式配置
const esmConfig = {
    ...baseConfig,
    experiments: {
        outputModule: true
    },
    output: {
        path: path.resolve(__dirname, "./build/"),
        filename: "parser.esm.js",
        library: {
            type: "module"
        }
    }
};

// ESM 压缩版本
const esmMinConfig = {
    ...baseMinConfig,
    experiments: {
        outputModule: true
    },
    output: {
        path: path.resolve(__dirname, "./build/"),
        filename: "parser.esm.min.js",
        library: {
            type: "module"
        }
    }
};

// IIFE 格式配置
const iifeConfig = {
    ...baseConfig,
    output: {
        path: path.resolve(__dirname, "./build/"),
        filename: "parser.iife.js",
        library: "FormulaParser",
        libraryTarget: "var"
    }
};

// IIFE 压缩版本
const iifeMinConfig = {
    ...baseMinConfig,
    output: {
        path: path.resolve(__dirname, "./build/"),
        filename: "parser.iife.min.js",
        library: "FormulaParser",
        libraryTarget: "var"
    }
};

// CommonJS 格式配置
const cjsConfig = {
    ...baseConfig,
    output: {
        path: path.resolve(__dirname, "./build/"),
        filename: "parser.cjs.js",
        library: {
            type: "commonjs2"
        }
    }
};

// CommonJS 压缩版本
const cjsMinConfig = {
    ...baseMinConfig,
    output: {
        path: path.resolve(__dirname, "./build/"),
        filename: "parser.cjs.min.js",
        library: {
            type: "commonjs2"
        }
    }
};

module.exports = [
    umdConfig, umdMinConfig,
    esmConfig, esmMinConfig,
    iifeConfig, iifeMinConfig,
    cjsConfig, cjsMinConfig
];
