const webpack = require('webpack');

module.exports = function override(config, env) {
  // Add fallbacks for Node.js core modules
  config.resolve.fallback = {
    ...config.resolve.fallback,
    fs: false,
    path: false,
    stream: require.resolve('stream-browserify'),
    zlib: require.resolve('browserify-zlib'),
    util: require.resolve('util'),
    crypto: require.resolve('crypto-browserify'),
    buffer: require.resolve('buffer'),
    process: require.resolve('process/browser'),
  };

  // Add plugins for browser polyfills
  config.plugins = [
    ...config.plugins,
    new webpack.ProvidePlugin({
      process: 'process/browser',
      Buffer: ['buffer', 'Buffer'],
    }),
  ];

  return config;
};
