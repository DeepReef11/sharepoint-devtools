const { merge } = require('webpack-merge');
const TerserPlugin = require('terser-webpack-plugin');
const common = require('./webpack.common.js');

module.exports = merge(common, {
  mode: 'production',
  devtool: false,
  optimization: {
    // Tree shaking and module concatenation
    usedExports: true,
    sideEffects: true,
    // Better module IDs for caching
    moduleIds: 'deterministic',
    minimizer: [
      new TerserPlugin({
        terserOptions: {
          compress: {
            // The content script runs on every SharePoint page, and this
            // extension's users live in DevTools — chatty logging is noise in
            // their console. Diagnostic logging is stripped from release builds;
            // console.error and console.warn are kept so real failures stay
            // reportable. Use `npm run build:dev` to keep everything.
            pure_funcs: ['console.log', 'console.debug', 'console.info'],
          },
        },
      }),
    ],
  },
});
