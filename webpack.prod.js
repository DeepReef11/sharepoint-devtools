const { merge } = require('webpack-merge');
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
  },
});
