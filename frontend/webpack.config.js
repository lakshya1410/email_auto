const path = require('path');
const HtmlWebpackPlugin = require('html-webpack-plugin');
const fs = require('fs');
const os = require('os');

const certPath = path.join(os.homedir(), '.office-addin-dev-certs');

module.exports = {
  mode: 'development',
  entry: {
    taskpane: './src/taskpane.js'
  },
  output: {
    path: path.resolve(__dirname, 'dist'),
    filename: '[name].js',
    clean: true
  },
  devtool: 'source-map',
  devServer: {
    static: {
      directory: path.join(__dirname, 'dist'),
    },
    port: 3000,
    hot: true,
    https: {
      key: fs.readFileSync(path.join(certPath, 'localhost.key')),
      cert: fs.readFileSync(path.join(certPath, 'localhost.crt'))
    },
    headers: {
      'Access-Control-Allow-Origin': '*',
      'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, PATCH, OPTIONS',
      'Access-Control-Allow-Headers': 'X-Requested-With, content-type, Authorization'
    },
    allowedHosts: 'all',
    historyApiFallback: {
      index: '/taskpane.html'
    },
    client: {
      overlay: {
        warnings: false,
        errors: true,
      },
    },
  },
  plugins: [
    new HtmlWebpackPlugin({
      template: './src/taskpane.html',
      filename: 'taskpane.html',
      chunks: ['taskpane'],
      inject: 'body'
    })
  ],
  module: {
    rules: [
      {
        test: /\.js$/,
        exclude: /node_modules/,
        use: {
          loader: 'babel-loader',
          options: {
            presets: ['@babel/preset-env']
          }
        }
      },
      {
        test: /\.css$/,
        use: ['style-loader', 'css-loader']
      }
    ]
  }
};
