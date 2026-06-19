import esbuild from "esbuild";
import { cpSync, existsSync } from "fs";
import builtins from "builtin-modules";

const prod = process.argv[2] === "production";

const outdir = "md-to-docx";

function copyStatic() {
  cpSync("manifest.json", `${outdir}/manifest.json`, { force: true });
  if (existsSync("styles.css")) {
    cpSync("styles.css", `${outdir}/styles.css`, { force: true });
  }
}

const context = await esbuild.context({
  entryPoints: ["main.ts"],
  bundle: true,
  external: [
    "obsidian",
    "electron",
    ...builtins,
  ],
  format: "cjs",
  target: "es2020",
  logLevel: "info",
  sourcemap: prod ? false : "inline",
  treeShaking: true,
  outfile: "md-to-docx/main.js",
  platform: "node",
});

if (prod) {
  await context.rebuild();
  copyStatic();
  console.log(`Built to ${outdir}/`);
  process.exit(0);
} else {
  copyStatic();
  await context.watch();
}
