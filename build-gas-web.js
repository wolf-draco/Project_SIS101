const fs = require("node:fs");
const path = require("node:path");

const root = __dirname;
const htmlPath = path.join(root, "index.html");
const cssPath = path.join(root, "app.css");
const appPath = path.join(root, "app.js");
const modelPath = path.join(root, "Codigo base", "Blu.js");
const outputPath = path.join(root, "codigo Apps Script", "WebApp.html");

let html = fs.readFileSync(htmlPath, "utf8").replace(/\r\n/g, "\n");
const css = fs.readFileSync(cssPath, "utf8");
const app = fs.readFileSync(appPath, "utf8");
const model = fs.readFileSync(modelPath, "utf8");

const stylesheetTag = '    <link rel="stylesheet" href="app.css">';
const scriptTags = [
  '    <script src="Codigo%20base/Blu.js"></script>',
  '    <script src="app.js"></script>',
].join("\n");

if (!html.includes(stylesheetTag) || !html.includes(scriptTags)) {
  throw new Error("No se encontraron las referencias esperadas a app.css, Blu.js y app.js en index.html.");
}

html = html.replace(stylesheetTag, `    <style>\n${css}\n    </style>`);
html = html.replace(scriptTags, `    <script>\n${model}\n    </script>\n    <script>\n${app}\n    </script>`);
fs.writeFileSync(outputPath, html, "utf8");
console.log(`Página de Apps Script generada: ${path.relative(root, outputPath)}`);