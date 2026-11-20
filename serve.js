// Server HTTPS lokal lintas-OS (menggantikan skrip lama yang memakai "~", tidak dikenali di Windows).
const fs = require("fs");
const os = require("os");
const path = require("path");
const httpServer = require("http-server");

const dir = path.join(os.homedir(), ".office-addin-dev-certs");
const cert = path.join(dir, "localhost.crt");
const key = path.join(dir, "localhost.key");
if (!fs.existsSync(cert) || !fs.existsSync(key)) {
  console.error("Sertifikat belum ada. Jalankan dulu: npm run certs");
  process.exit(1);
}
const server = httpServer.createServer({
  root: __dirname,
  cache: -1,
  https: { cert, key },
});
server.listen(3000, () => console.log("Siap di https://localhost:3000 (Ctrl+C untuk berhenti)"));
