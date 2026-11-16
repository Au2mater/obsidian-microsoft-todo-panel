// Builds the plugin: checks src/main.js parses and writes it, unchanged, to
// main.js, the file Obsidian loads. The plugin is plain JavaScript with no
// dependencies, so the build is a checked copy and its output is reproducible
// byte for byte.
const fs = require("fs");
const path = require("path");
const vm = require("vm");

const root = path.join(__dirname, "..");
const source = fs.readFileSync(path.join(root, "src", "main.js"));
new vm.Script(source.toString("utf8"), { filename: "src/main.js" }); // throws on a syntax error
fs.writeFileSync(path.join(root, "main.js"), source);
console.log(`Built main.js (${source.length} bytes)`);
