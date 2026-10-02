// Copies the app's static files into public/ for Vercel, plus the Excel library from node_modules.
import fs from "fs";
fs.rmSync("public", { recursive: true, force: true });
fs.mkdirSync("public/fonts", { recursive: true });
for (const f of ["index.html", "shim.js"]) fs.copyFileSync(f, "public/" + f);
for (const f of fs.readdirSync("fonts")) fs.copyFileSync("fonts/" + f, "public/fonts/" + f);
fs.copyFileSync("node_modules/exceljs/dist/exceljs.min.js", "public/exceljs.min.js");
console.log("Built public/");
