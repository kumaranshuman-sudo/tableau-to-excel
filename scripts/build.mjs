/* Builds the extension into dist/:
 *   dist/extension.<hash>.js   src/main.js bundled with ExcelJS and JSZip (no CDN needed)
 *   dist/index.html            src/index.html pointing at that file – a new hash on every change,
 *                              so Tableau's browser never runs a stale cached copy
 *   dist/js/                   the Tableau Extensions API library
 *
 *   node scripts/build.mjs            one-off build
 *   node scripts/build.mjs --serve    rebuild on every save and serve dist/ on port 5500
 *                                     (the URL in Export.trex)                                  */
import * as esbuild from "esbuild";
import fs from "node:fs";
import http from "node:http";
import path from "node:path";
import { fileURLToPath } from "node:url";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const dist = path.join(root, "dist");
const serve = process.argv.includes("--serve");
const port = Number(process.env.PORT || 5500);

/** Writes index.html for the bundle just built and removes bundles from earlier builds. */
function writePage(result) {
  const outputs = Object.keys(result.metafile.outputs).map(p => path.basename(p));
  const bundle = outputs.find(f => /^extension\.[A-Z0-9]+\.js$/i.test(f));
  if (!bundle) throw new Error("bundle missing from build output");
  const keep = new Set(outputs);
  for (const f of fs.readdirSync(dist)) {
    if (/^extension\.[A-Z0-9]+\.js(\.map)?$/i.test(f) && !keep.has(f)) fs.rmSync(path.join(dist, f));
  }
  const template = fs.readFileSync(path.join(root, "src", "index.html"), "utf8");
  if (!template.includes("__BUNDLE__")) throw new Error("src/index.html has no __BUNDLE__ placeholder");
  fs.writeFileSync(path.join(dist, "index.html"), template.replace("__BUNDLE__", "./" + bundle));
  fs.mkdirSync(path.join(dist, "js"), { recursive: true });
  fs.copyFileSync(path.join(root, "js", "tableau.extensions.1.latest.js"), path.join(dist, "js", "tableau.extensions.1.latest.js"));
  const kb = Math.round(fs.statSync(path.join(dist, bundle)).size / 1024);
  console.log(`[build] ${new Date().toLocaleTimeString()}  dist/${bundle} (${kb} KB)`);
}

const pagePlugin = {
  name: "page",
  setup(build) {
    build.onEnd(result => {
      if (result.errors.length) { console.error(`[build] ${result.errors.length} error(s) – dist/ not updated`); return; }
      writePage(result);
    });
    // a change to the page template is a rebuild too
    build.onLoad({ filter: /src[\\/]main\.js$/ }, () => ({ watchFiles: [path.join(root, "src", "index.html")] }));
  }
};

fs.mkdirSync(dist, { recursive: true });
const options = {
  entryPoints: { extension: path.join(root, "src", "main.js") },
  outdir: dist,
  entryNames: "[name].[hash]",
  bundle: true,
  format: "iife",
  platform: "browser",
  target: ["chrome100", "edge100", "safari15"],
  sourcemap: "linked",
  metafile: true,
  logLevel: "warning",
  plugins: [pagePlugin]
};

if (serve) {
  const ctx = await esbuild.context(options);
  await ctx.watch();
  // esbuild's server sets no cache headers; a browser that heuristically caches index.html would keep
  // pointing at an old bundle. A small proxy in front adds "Cache-Control: no-cache" to every response.
  const inner = await ctx.serve({ servedir: dist, host: "127.0.0.1", port: 0 });
  http.createServer((req, res) => {
    const upstream = http.request({ hostname: "127.0.0.1", port: inner.port, path: req.url, method: req.method, headers: req.headers }, up => {
      res.writeHead(up.statusCode, { ...up.headers, "cache-control": "no-cache" });
      up.pipe(res, { end: true });
    });
    upstream.on("error", () => { res.writeHead(502); res.end("build server unavailable"); });
    req.pipe(upstream, { end: true });
  }).listen(port, "127.0.0.1", () => console.log(`[build] watching src/ – serving dist/ on http://localhost:${port}/index.html`));
} else {
  await esbuild.build(options);
}
