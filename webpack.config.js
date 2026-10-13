const path = require('path');
const HtmlWebpackPlugin = require('html-webpack-plugin');
const CopyWebpackPlugin = require('copy-webpack-plugin');
const devCerts = require('office-addin-dev-certs');

module.exports = async (env, argv) => {
  const isDev = argv.mode !== 'production';
  const httpsOptions = isDev ? await devCerts.getHttpsServerOptions() : undefined;

  return {
    entry: {
      taskpane: './src/taskpane/taskpane.ts'
    },
    devtool: isDev ? 'source-map' : false,
    resolve: {
      extensions: ['.ts', '.js', '.json']
    },
    module: {
      rules: [
        {
          test: /\.ts$/,
          exclude: /node_modules/,
          use: {
            loader: 'ts-loader',
            options: {
              // The project-wide tsconfig is no-emit for standalone checks;
              // Webpack must still receive transpiled modules from ts-loader.
              compilerOptions: { noEmit: false }
            }
          }
        },
        { test: /\.css$/, use: ['style-loader', 'css-loader'] }
      ]
    },
    plugins: [
      new HtmlWebpackPlugin({
        filename: 'taskpane.html',
        template: './src/taskpane/taskpane.html',
        chunks: ['taskpane']
      }),
      new CopyWebpackPlugin({
        patterns: [
          { from: 'assets', to: 'assets', noErrorOnMissing: true }
        ]
      })
    ],
    output: {
      filename: '[name].js',
      path: path.resolve(__dirname, 'dist'),
      clean: true
    },
    devServer: {
      server: httpsOptions ? { type: 'https', options: httpsOptions } : 'https',
      port: 3000,
      headers: { 'Access-Control-Allow-Origin': '*' },
      static: { directory: path.join(__dirname, 'dist') },
      hot: true
    }
  };
};
