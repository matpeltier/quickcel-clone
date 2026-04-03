const path = require("path");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const MiniCssExtractPlugin = require("mini-css-extract-plugin");

const fs = require("fs");

const homeDir = process.env.USERPROFILE || process.env.HOME;
const certDir = path.join(homeDir, ".office-addin-dev-certs");
const httpsKeyPath = path.join(certDir, "localhost.key");
const httpsCertPath = path.join(certDir, "localhost.crt");

let serverConfig = "https";
if (fs.existsSync(httpsKeyPath) && fs.existsSync(httpsCertPath)) {
  serverConfig = {
    type: "https",
    options: {
      key: fs.readFileSync(httpsKeyPath),
      cert: fs.readFileSync(httpsCertPath),
    },
  };
}

module.exports = {
  devtool: "source-map",
  entry: {
    taskpane: "./src/taskpane/taskpane.tsx",
    commands: "./src/commands/commands.ts",
  },
  output: {
    path: path.resolve(__dirname, "dist"),
    filename: "[name].js",
    clean: true,
  },
  resolve: {
    extensions: [".ts", ".tsx", ".html", ".js"],
    alias: {
      "@shared": path.resolve(__dirname, "src/shared"),
      "@commands": path.resolve(__dirname, "src/commands"),
    },
  },
  module: {
    rules: [
      {
        test: /\.tsx?$/,
        use: "ts-loader",
        exclude: /node_modules/,
      },
      {
        test: /\.css$/,
        use: [MiniCssExtractPlugin.loader, "css-loader"],
      },
      {
        test: /\.html$/,
        use: "html-loader",
      },
    ],
  },
  plugins: [
    new MiniCssExtractPlugin({ filename: "[name].css" }),
    new CopyWebpackPlugin({
      patterns: [
        { from: "src/taskpane/taskpane.html", to: "taskpane.html" },
        { from: "shortcuts.json", to: "shortcuts.json" },
        { from: "assets/", to: "assets/" },
      ],
    }),
  ],
  devServer: {
    static: {
      directory: path.join(__dirname, "dist"),
    },
    compress: true,
    port: 3000,
    server: serverConfig,
    headers: {
      "Access-Control-Allow-Origin": "*",
    },
  },
};
