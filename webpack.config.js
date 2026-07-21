/* eslint-disable no-undef */

const devCerts = require("office-addin-dev-certs");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const HtmlWebpackPlugin = require("html-webpack-plugin");

const urlDev = "https://localhost:3000/";
const urlProd = process.env.PROD_URL || "https://YOUR_DOMAIN_HERE/";

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
      commands: "./src/commands/commands.js",
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
          test: /\.(png|jpg|jpeg|gif|ico)$/,
          type: "asset/resource",
          generator: {
            filename: "assets/[name][ext][query]",
          },
        },
      ],
    },
    plugins: [
      new HtmlWebpackPlugin({
        filename: "index.html",
        inject: false,
        templateContent: () => `<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>Illuminate - WordCount Tracker</title>
  <style>
    :root {
      --bg: #0f1014;
      --card: #171923;
      --text: #f4f6fb;
      --muted: #b6bfd4;
      --primary: #2ccf7b;
      --primary-hover: #21ad67;
      --outline: #6f7aa1;
    }

    * { box-sizing: border-box; }
    body {
      margin: 0;
      min-height: 100vh;
      font-family: "Segoe UI", Tahoma, sans-serif;
      background: radial-gradient(circle at 25% 15%, #20263a 0%, var(--bg) 42%);
      color: var(--text);
      display: grid;
      place-items: center;
      padding: 24px;
    }

    .card {
      width: min(760px, 100%);
      background: linear-gradient(180deg, #1a1f2e 0%, var(--card) 100%);
      border: 1px solid #2c3550;
      border-radius: 14px;
      padding: 24px;
      box-shadow: 0 20px 60px rgba(0, 0, 0, 0.35);
    }

    h1 { margin: 0 0 8px; font-size: 31px; line-height: 1.1; }
    p { margin: 0; color: var(--muted); }
    .actions { margin-top: 18px; display: flex; flex-wrap: wrap; gap: 10px; }

    .btn {
      display: inline-block;
      text-decoration: none;
      border-radius: 10px;
      padding: 11px 14px;
      font-weight: 700;
      font-size: 14px;
      border: 1px solid transparent;
    }

    .btn-primary {
      background: var(--primary);
      color: #08120b;
    }

    .btn-primary:hover { background: var(--primary-hover); }

    .btn-secondary {
      color: var(--text);
      border-color: var(--outline);
      background: transparent;
    }

    .btn-secondary:hover { border-color: #93a0cf; }

    ol {
      margin: 18px 0 0;
      padding-left: 20px;
      color: var(--muted);
      line-height: 1.6;
    }

    code {
      font-family: ui-monospace, SFMono-Regular, Menlo, Consolas, monospace;
      background: #121725;
      border: 1px solid #2a344f;
      border-radius: 6px;
      padding: 2px 5px;
      color: #e7ecfa;
    }
  </style>
</head>
<body>
  <main class="card">
    <h1>Illuminate - WordCount Tracker</h1>
    <p>Track writing time and word changes in Microsoft Word.</p>

    <div class="actions">
      <a class="btn btn-primary" href="manifest.word.prod.template.xml" download>Install to MS Word</a>
      <a class="btn btn-secondary" href="taskpane.html">Open Task Pane Preview</a>
    </div>

    <ol>
      <li>Click <strong>Install to MS Word</strong> to download the manifest file.</li>
      <li>In Word: <strong>Insert</strong> -> <strong>Add-ins</strong> -> <strong>My Add-ins</strong> -> <strong>Upload My Add-in</strong>.</li>
      <li>Select the downloaded <code>manifest.word.prod.template.xml</code> file.</li>
    </ol>
  </main>
</body>
</html>`,
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
                return content.toString().replace(new RegExp(urlDev, "g"), urlProd);
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
    devServer: {
      headers: {
        "Access-Control-Allow-Origin": "*",
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
