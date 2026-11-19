// Tiny HTTPS static server for local development. Office add-ins must be served over HTTPS.
// Uses the trusted localhost certificate created by `npm run certs`.
const https = require("https");
const fs = require("fs");
const path = require("path");
const os = require("os");

const PORT = Number(process.env.PORT) || 3000;
const ROOT = path.join(__dirname, "src");
const CERT_DIR = path.join(os.homedir(), ".office-addin-dev-certs");

const TYPES = {
  ".html": "text/html; charset=utf-8",
  ".js": "text/javascript; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".png": "image/png",
  ".svg": "image/svg+xml",
  ".json": "application/json",
  ".xml": "application/xml",
};

let tls;
try {
  tls = {
    key: fs.readFileSync(path.join(CERT_DIR, "localhost.key")),
    cert: fs.readFileSync(path.join(CERT_DIR, "localhost.crt")),
  };
} catch {
  console.error(`No dev certificate found in ${CERT_DIR}. Run "npm run certs" first.`);
  process.exit(1);
}

https.createServer(tls, (req, res) => {
  const urlPath = decodeURIComponent(new URL(req.url, "https://localhost").pathname);
  const file = path.join(ROOT, urlPath === "/" ? "index.html" : urlPath);
  if (!file.startsWith(ROOT + path.sep)) {
    res.writeHead(403).end();
    return;
  }
  fs.readFile(file, (err, data) => {
    if (err) {
      res.writeHead(404, { "Content-Type": "text/plain" }).end("Not found");
      return;
    }
    res.writeHead(200, {
      "Content-Type": TYPES[path.extname(file)] || "application/octet-stream",
      "Cache-Control": "no-store",
    });
    res.end(data);
  });
}).listen(PORT, () => console.log(`Trello for Outlook running at https://localhost:${PORT}`));
