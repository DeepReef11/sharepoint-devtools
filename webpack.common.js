const path = require('path');
const CopyWebpackPlugin = require('copy-webpack-plugin');
const { CleanWebpackPlugin } = require('clean-webpack-plugin');

module.exports = {
  entry: {
    background: './src/background/service-worker.ts',
    content: './src/content/content-script.ts',
    options: './src/options/options.ts',
  },
  module: {
    rules: [
      {
        test: /\.tsx?$/,
        use: {
          loader: 'ts-loader',
          options: {
            transpileOnly: true,
          },
        },
        exclude: /node_modules/,
      },
      {
        test: /\.css$/,
        use: ['style-loader', 'css-loader'],
      },
    ],
  },
  resolve: {
    extensions: ['.tsx', '.ts', '.js'],
  },
  output: {
    filename: '[name].js',
    path: path.resolve(__dirname, 'dist'),
    chunkFilename: '[name].[contenthash:8].chunk.js',
  },
  optimization: {
    splitChunks: {
      chunks: 'all',
      cacheGroups: {
        // Vendor code (node_modules) - split to improve caching
        vendor: {
          test: /[\\/]node_modules[\\/]/,
          name: 'vendors',
          priority: 10,
          reuseExistingChunk: true,
        },
        // Common code shared between chunks
        common: {
          minChunks: 2,
          name: 'common',
          priority: 5,
          reuseExistingChunk: true,
        },
      },
    },
  },
  plugins: [
    new CleanWebpackPlugin(),
    new CopyWebpackPlugin({
      patterns: [
        { from: 'public', to: '.' },
        { from: 'src/options/options.html', to: 'options.html' },
        // options.html links this by relative path. Without it the whole page
        // renders unstyled and, worse, `.modal { display: none }` never applies,
        // so both dialogs sit permanently expanded in the page body.
        { from: 'src/options/options.css', to: 'options.css' },
      ],
    }),
  ],
};
