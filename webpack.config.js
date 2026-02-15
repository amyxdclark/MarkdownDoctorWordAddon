const path = require("path");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const CopyWebpackPlugin = require("copy-webpack-plugin");

module.exports = {
  entry: {
    taskpane: "./taskpane.js"
  },
  output: {
    path: path.resolve(__dirname, "dist"),
    filename: "[name].js"
  },
  devServer: {
    static: {
      directory: path.join(__dirname, "dist")
    },
    port: 3000,
    hot: true,
    headers: {
      "Access-Control-Allow-Origin": "*"
    },
    server: "https"
  },
  plugins: [
    new HtmlWebpackPlugin({
      template: "./taskpane.html",
      filename: "taskpane.html",
      chunks: ["taskpane"]
    }),
    new CopyWebpackPlugin({
      patterns: [
        { from: "taskpane.css", to: "taskpane.css" },
        { from: "commands.html", to: "commands.html" },
        { from: "assets/*", to: "assets/[name][ext]" },
        { from: "manifest.xml", to: "manifest.xml" }
      ]
    })
  ],
  mode: "development"
};
