const merge = require('webpack-merge');
const common = require('./webpack.common.js');
const webpack = require('webpack');
var ExtractTextPlugin = require('extract-text-webpack-plugin');


module.exports = merge(common, {
    output: {
        filename: "./dist/bundle.js"
    },
    plugins: [
        new ExtractTextPlugin({
            filename: './dist/bundle.css',
            allChunks: true
        }),
        new webpack.DefinePlugin({
            DIRECTLINE_SECRET: JSON.stringify('uVAIRSveWlA.cwA.rC0.86qtVL1RV8zm3Hh7p3I7RZE9j5eddaK-G_l9lmFOXI8'),
            ROOT_SITES: JSON.stringify(['', 'https://mykier']),
            DEVELOPMENT: JSON.stringify(false),
            MYKIER_URL: JSON.stringify('https://mykier')
        })
    ]
});