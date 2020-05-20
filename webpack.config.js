const path = require('path');
const webpack = require('webpack');
const UglifyJsPlugin = require('terser-webpack-plugin');

module.exports = {
    entry: {
        'm365-wrapper': './src/index.ts',
        'm365-wrapper.min': './src/index.ts'
    },
    output: {
        path: path.resolve(__dirname, '_bundles'),
        filename: '[name].js',
        libraryTarget: 'umd',
        library: 'M365Wrapper',
        umdNamedDefine: true
    },
    resolve: {
        extensions: ['.ts', '.tsx', '.js']
    },
    devtool: 'source-map',
    // plugins: [
    //     new webpack.optimize.UglifyJsPlugin({
    //         minimize: true,
    //         sourceMap: true,
    //         include: /\.min\.js$/,
    //     })
    // ],
    optimization: {
        minimize: false,
        minimizer: [new UglifyJsPlugin( {
            sourceMap: true
        })],
      },
    module: {
        rules: [{
            test: /\.tsx?$/,
            loader: 'awesome-typescript-loader',
            exclude: /node_modules/,
            query: {
                declaration: false,
            }
        }]
    }
};