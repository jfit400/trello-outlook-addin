// Builds manifest.prod.xml from manifest.xml, pointing every URL at the hosted site.
// Usage: node scripts/make-manifest.js https://jfit400.github.io/trello-outlook-addin
const fs = require("fs");
const path = require("path");

const base = (process.argv[2] || "").replace(/\/+$/, "");
if (!/^https:\/\/[^/]+/.test(base)) {
  console.error("Usage: node scripts/make-manifest.js https://<your-site>/<path>");
  process.exit(1);
}

const root = path.join(__dirname, "..");
const xml = fs.readFileSync(path.join(root, "manifest.xml"), "utf8")
  .replaceAll("https://localhost:3000", base)
  // Separate ID so the hosted add-in doesn't collide with the localhost dev copy.
  .replace(/<Id>[^<]+<\/Id>/, "<Id>e5d0e026-62fa-4856-9b8a-c473838ec79b</Id>");

fs.writeFileSync(path.join(root, "manifest.prod.xml"), xml);
console.log(`Wrote manifest.prod.xml for ${base}`);
