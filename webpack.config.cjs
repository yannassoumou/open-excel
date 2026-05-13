/* eslint-disable no-undef */

const CopyWebpackPlugin = require("copy-webpack-plugin");
const CustomFunctionsMetadataPlugin = require("custom-functions-metadata-plugin");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const fs = require("fs");
const os = require("os");
const path = require("path");

const urlDev = "https://localhost:3000/";
const urlProd = process.env.VERCEL_URL
  ? `https://${process.env.VERCEL_URL}/`
  : "https://excel-ten-theta.vercel.app/"; // CHANGE THIS TO YOUR PRODUCTION DEPLOYMENT LOCATION

/* global require, module, process, __dirname */

async function getHttpsOptions() {
  const certDir = path.join(os.homedir(), ".office-addin-dev-certs");
  return {
    ca: fs.readFileSync(path.join(certDir, "ca.crt")),
    key: fs.readFileSync(path.join(certDir, "localhost.key")),
    cert: fs.readFileSync(path.join(certDir, "localhost.crt")),
  };
}

module.exports = async (env, options) => {
  const dev = options.mode === "development";
  const config = {
    devtool: "source-map",
    entry: {
      polyfill: ["core-js/stable", "regenerator-runtime/runtime"],
      taskpane: ["./src/taskpane/taskpane.js", "./src/taskpane/taskpane.html", "./src/taskpane/agentChat.css"],
      commands: "./src/commands/commands.js",
      functions: "./src/functions/functions.js",
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
          },
        },
        {
          test: /\.html$/,
          exclude: /node_modules/,
          use: "html-loader",
        },
        {
          test: /\.css$/,
          use: ["style-loader", "css-loader"],
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
      new CustomFunctionsMetadataPlugin({
        output: "functions.json",
        input: "./src/functions/functions.js",
      }),
      new HtmlWebpackPlugin({
        filename: "functions.html",
        template: "./src/functions/functions.html",
        chunks: ["polyfill", "functions"],
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
            from: "manifest*.xml",
            to: "[name]" + "[ext]",
            transform(content) {
              if (dev) {
                return content;
              } else {
                return content.toString().replace(new RegExp(urlDev + "(?:public/)?", "g"), urlProd);
              }
            },
          },
        ],
      }),
      new HtmlWebpackPlugin({
        filename: "commands.html",
        template: "./src/commands/commands.html",
        chunks: ["polyfill", "commands"],
      }),
    ],
  };

  if (!dev) {
    return config;
  }

  config.devServer = {
    host: "0.0.0.0",
    allowedHosts: "all",
    static: {
      directory: path.join(__dirname, "dist"),
      publicPath: "/public",
    },
    headers: {
      "Access-Control-Allow-Origin": "*",
      "ngrok-skip-browser-warning": "true",
    },
    server: {
      type: "https",
      options: await getHttpsOptions(),
    },
    port: process.env.npm_package_config_dev_server_port || 3000,
  };

  return config;
};
