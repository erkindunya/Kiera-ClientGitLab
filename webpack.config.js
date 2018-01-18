
var babelOptions = {
    "presets": [
        "es2016",
        "es2015"
    ]
};


module.exports = {
    entry: [
        'babel-polyfill',
        "./scripts/Kiera.ts"
    ],
    resolve: {
        modules: ['.', './node_modules'],
        extensions: [".ts", ".js"]
    },
    output: {
        filename: "bundle-test.js"
    },
    module: {
        rules: [
            {
                test: /\.ts$/,
                exclude: /(node_modules|bower_components)/,
                use: [
                    {
                        loader: 'babel-loader',
                        options: babelOptions
                    },
                    {
                        loader: 'ts-loader'
                    }
                ]
            }
        ]
    }
};