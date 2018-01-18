var ExtractTextPlugin = require('extract-text-webpack-plugin');

var babelOptions = {
    "presets": [
        "es2016",
        "es2015"
    ]
};


module.exports = {
    entry: [
        'babel-polyfill',
        "./scripts/Kiera.ts",
        "./styles/main.scss"
    ],
    resolve: {
        modules: ['.', './node_modules'],
        extensions: [".ts", ".js"]
    },
    output: {
        filename: "./dist/bundle-test.js"
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
            },
            { // regular css files
                test: /\.css$/,
                loader: ExtractTextPlugin.extract({
                    use: 'css-loader?importLoaders=1',
                }),
            },
            { // sass / scss loader for webpack
                test: /\.(sass|scss)$/,
                loader: ExtractTextPlugin.extract(['css-loader', 'sass-loader'])
            }
        ]
    },
    plugins: [
        new ExtractTextPlugin({
            filename: './dist/bundle-test.css',
            allChunks: true
          })
    ]
};