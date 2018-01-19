const merge = require('webpack-merge');
const common = require('./webpack.common.js');
const webpack = require('webpack');
var ExtractTextPlugin = require('extract-text-webpack-plugin');


module.exports = merge(common, {
    output: {
        filename: "./dist/bundle-test.js"
    },
    plugins: [
        new ExtractTextPlugin({
            filename: './dist/bundle-test.css',
            allChunks: true
        }),
        new webpack.DefinePlugin({
            DIRECTLINE_SECRET: JSON.stringify('oCx5Pd_G9OQ.cwA.ZTc.kOI4BcVwzJNfokl651HPhLueIXCO0rjkoMAUOE1D0ak'),
        })
    ]
});