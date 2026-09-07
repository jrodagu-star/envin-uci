// Genera envin-data.js sin NHC ni iniciales a partir de envin-data.local.js
const fs = require("fs");
const vm = require("vm");
const path = require("path");
const dir = __dirname;
const localPath = path.join(dir, "envin-data.local.js");
const pubPath = path.join(dir, "envin-data.js");
if (!fs.existsSync(localPath)) {
  console.error("No existe envin-data.local.js. Ejecuta primero export-envin.ps1.");
  process.exit(1);
}
const ctx = { window: {} };
vm.runInNewContext(fs.readFileSync(localPath, "utf8"), ctx);
const E = ctx.window.ENVIN;
if (!E) {
  console.error("envin-data.local.js no define window.ENVIN");
  process.exit(1);
}
const ids = Object.create(null);
let n = 0;
function stayId(nhc, f0) {
  const k = String(nhc || "").trim() + "|" + (f0 || "");
  if (!ids[k]) ids[k] = String(++n);
  return ids[k];
}
(E.ingresos || []).forEach((p) => stayId(p.nhc, p.f0));
(E.infecciones || []).forEach((inf) => stayId(inf.nhc, inf.f0));
(E.atb || []).forEach((a) => stayId(a.nhc, a.f0));
(E.ingresos || []).forEach((p) => {
  p.nhc = stayId(p.nhc, p.f0);
  delete p.ini;
});
(E.infecciones || []).forEach((inf) => {
  inf.nhc = stayId(inf.nhc, inf.f0);
});
(E.atb || []).forEach((a) => {
  a.nhc = stayId(a.nhc, a.f0);
});
E.meta = E.meta || {};
E.meta.anon = true;
E.meta.fuente = "ENVINPAZ.accdb (anonimizado)";
fs.writeFileSync(pubPath, "window.ENVIN=" + JSON.stringify(E) + ";\n");
console.log("OK " + (fs.statSync(pubPath).size / 1048576).toFixed(2) + " MB · " + n + " estancias");
