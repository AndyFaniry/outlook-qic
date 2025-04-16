/* eslint-disable no-undef */
const webpack = require("webpack");
const Dotenv = require("dotenv-webpack");
const devCerts = require("office-addin-dev-certs");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const HtmlWebpackPlugin = require("html-webpack-plugin");

const URL_DEV = process.env.URL_DEV;
const URL_PROD = process.env.URL_PROD;

async function getHttpsOptions() {
  const httpsOptions = await devCerts.getHttpsServerOptions();
  return { ca: httpsOptions.ca, key: httpsOptions.key, cert: httpsOptions.cert };
}

module.exports = async (env, options) => {
  const dev = options.mode === "development";
  const config = {
    devtool: "source-map",
    entry: {
      polyfill: ["core-js/stable", "regenerator-runtime/runtime"],
      taskpane: ["./src/taskpane/taskpane.js", "./src/taskpane/taskpane.html"],
      acceuil: ["./src/acceuil/acceuil.js", "./src/acceuil/acceuil.html"],
      profil: ["./src/profil/profil.js", "./src/profil/profil.html"],
      parametre: ["./src/parametre/parametre.js", "./src/parametre/parametre.html"],
    },
    output: {
      clean: true,
    },
    resolve: {
      extensions: [".html", ".js"],
    },
    module: {
      rules: [
        {
          test: /\.js$/,
          exclude: /node_modules/,
          use: {
            loader: "babel-loader",
            options: {
              presets: ["@babel/preset-env"],
            },
          },
        },
        {
          test: /\.html$/,
          exclude: /node_modules/,
          use: "html-loader",
        },
        {
          test: /\.(png|jpg|jpeg|gif|ico)$/,
          type: "asset/resource",
          generator: {
            filename: "assets/[name][ext][query]",
          },
        },
      ],
    },
    plugins: [
      new Dotenv(),
      new webpack.ProvidePlugin({
        $: "jquery",
        jQuery: "jquery",
      }),
      new HtmlWebpackPlugin({
        filename: "taskpane.html",
        template: "./src/taskpane/taskpane.html",
        chunks: ["polyfill", "taskpane"],
      }),
      new CopyWebpackPlugin({
        patterns: [
          {
            from: "assets/*",
            to: "assets/[name][ext][query]",
          },
          {
            from: "manifest*.json",
            to: "[name]" + "[ext]",
            transform(content) {
              if (dev) {
                return content;
              } else {
                return content.toString().replace(new RegExp(URL_DEV, "g"), URL_PROD);
              }
            },
          },
        ],
      }),
      new HtmlWebpackPlugin({
        filename: "acceuil.html",
        template: "./src/acceuil/acceuil.html",
        chunks: ["polyfill", "acceuil"],
      }),
      new HtmlWebpackPlugin({
        filename: "profil.html",
        template: "./src/profil/profil.html",
        chunks: ["polyfill", "profil"],
      }),
      new HtmlWebpackPlugin({
        filename: "parametre.html",
        template: "./src/parametre/parametre.html",
        chunks: ["polyfill", "parametre"],
      }),
    ],
    devServer: {
      headers: {
        "Access-Control-Allow-Origin": "*",
        "X-Frame-Options": "ALLOWALL",
      },
      server: {
        type: "https",
        options: env.WEBPACK_BUILD || options.https !== undefined ? options.https : await getHttpsOptions(),
      },
      port: process.env.npm_package_config_dev_server_port || 3000,
    },
  };

  return config;
};
