const path = require('path');
const HtmlWebpackPlugin = require('html-webpack-plugin');
const CopyWebpackPlugin = require('copy-webpack-plugin');
const MiniCssExtractPlugin = require('mini-css-extract-plugin');
const devCerts = require('office-addin-dev-certs');

const isProduction = process.env.NODE_ENV === 'production';

async function getHttpsOptions() {
  try {
    const httpsOptions = await devCerts.getHttpsServerOptions();
    return {
      ca: httpsOptions.ca,
      key: httpsOptions.key,
      cert: httpsOptions.cert,
    };
  } catch (error) {
    console.warn('Unable to get HTTPS certificates. Falling back to HTTP.');
    console.warn('Run "npx office-addin-dev-certs install" to install certificates.');
    return false;
  }
}

module.exports = async (env, argv) => {
  const mode = argv.mode || 'development';
  const isDev = mode === 'development';
  const httpsOptions = isDev ? await getHttpsOptions() : false;

  return {
    mode,
    devtool: isDev ? 'source-map' : false,

    entry: {
      taskpane: './src/taskpane/taskpane.ts',
      commands: './src/commands/commands.ts',
    },

    output: {
      path: path.resolve(__dirname, 'dist'),
      filename: '[name].bundle.js',
      clean: true,
    },

    resolve: {
      extensions: ['.ts', '.tsx', '.js', '.jsx'],
      alias: {
        '@': path.resolve(__dirname, 'src'),
      },
    },

    module: {
      rules: [
        {
          test: /\.tsx?$/,
          use: 'ts-loader',
          exclude: /node_modules/,
        },
        {
          test: /\.css$/,
          use: [
            isDev ? 'style-loader' : MiniCssExtractPlugin.loader,
            'css-loader',
          ],
        },
        {
          test: /\.(png|jpg|jpeg|gif|svg|ico)$/,
          type: 'asset/resource',
          generator: {
            filename: 'assets/[name][ext]',
          },
        },
      ],
    },

    plugins: [
      // Taskpane HTML
      new HtmlWebpackPlugin({
        template: './src/taskpane/taskpane.html',
        filename: 'taskpane.html',
        chunks: ['taskpane'],
      }),

      // Help HTML
      new HtmlWebpackPlugin({
        template: './src/taskpane/help.html',
        filename: 'help.html',
        chunks: [],
        inject: false,
      }),

      // Commands HTML (for function commands)
      new HtmlWebpackPlugin({
        template: './src/commands/commands.html',
        filename: 'commands.html',
        chunks: ['commands'],
      }),

      // Extract CSS in production
      ...(isDev
        ? []
        : [
            new MiniCssExtractPlugin({
              filename: '[name].css',
            }),
          ]),

      // Copy static assets
      new CopyWebpackPlugin({
        patterns: [
          {
            from: 'assets',
            to: 'assets',
            noErrorOnMissing: true,
          },
          {
            from: 'src/taskpane/help.css',
            to: 'help.css',
          },
          {
            from: 'src/taskpane/help.js',
            to: 'help.js',
          },
          {
            from: 'manifest.xml',
            to: 'manifest.xml',
          },
        ],
      }),
    ],

    devServer: {
      static: {
        directory: path.join(__dirname, 'dist'),
      },
      headers: {
        'Access-Control-Allow-Origin': '*',
      },
      server: httpsOptions
        ? {
            type: 'https',
            options: httpsOptions,
          }
        : 'http',
      port: 3000,
      hot: true,
      open: false,
      devMiddleware: {
        writeToDisk: true,
      },
    },

    optimization: {
      splitChunks: {
        chunks: 'all',
        cacheGroups: {
          vendor: {
            test: /[\\/]node_modules[\\/]/,
            name: 'vendors',
            chunks: 'all',
          },
        },
      },
    },

    performance: {
      hints: isDev ? false : 'warning',
      maxAssetSize: 500000,
      maxEntrypointSize: 500000,
    },
  };
};
