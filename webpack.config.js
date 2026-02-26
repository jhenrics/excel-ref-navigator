const path = require("path");
const fs = require("fs");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const CopyWebpackPlugin = require("copy-webpack-plugin");

module.exports = {
  entry: {
    taskpane: "./src/taskpane.js",
  },
  output: {
    path: path.resolve(__dirname, "dist"),
    filename: "[name].bundle.js",
    clean: true,
  },
  resolve: {
    extensions: [".js"],
  },
  module: {
    rules: [
      {
        test: /\.js$/,
        exclude: /node_modules/,
        use: "babel-loader",
      },
    ],
  },
  plugins: [
    new HtmlWebpackPlugin({
      filename: "taskpane.html",
      template: "./src/taskpane.html",
      chunks: ["taskpane"],
    }),
    new CopyWebpackPlugin({
      patterns: [{ from: "assets", to: "assets", noErrorOnMissing: true }],
    }),
  ],
  devServer: {
    static: {
      directory: path.join(__dirname, "dist"),
    },
    headers: {
      "Access-Control-Allow-Origin": "*",
    },
    server: {
      type: "https",
      options: {
        key: fs.readFileSync(path.resolve(process.env.HOME || process.env.USERPROFILE, ".office-addin-dev-certs/localhost.key")),
        cert: fs.readFileSync(path.resolve(process.env.HOME || process.env.USERPROFILE, ".office-addin-dev-certs/localhost.crt")),
        ca: fs.readFileSync(path.resolve(process.env.HOME || process.env.USERPROFILE, ".office-addin-dev-certs/ca.crt")),
      },
    },
    port: 3000,
    hot: true,
  },
};
