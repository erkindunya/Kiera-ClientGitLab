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
            DIRECTLINE_SECRET: JSON.stringify('kfWac62Fmic.cwA.xSk.zzRLxsumU1cMyFLOpuEIE19XX92kl7D4o5UMCZxSnOk'),
        })
    ]
});