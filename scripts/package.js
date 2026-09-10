#!/usr/bin/env node
//
// Builds the store-ready extension zip.
//
// This exists because hand-zipping the project folder produces an archive
// with everything nested under a top-level folder, which both the Edge
// Add-ons and Chrome Web Store reject on upload -- the manifest has to sit at
// the archive root. "Load unpacked" is happy either way, so the mistake
// survives local testing and only surfaces at submission. This script writes
// paths relative to the extension root, so the shape is right by
// construction.
//
// It also ships only what the extension needs. The tests, this script, the
// README and package.json are development files; a reviewer opening the
// package should see an extension, not a repository.
//
// No dependencies on purpose: the zip is written directly (zlib is in Node's
// standard library), so `npm run package` works on a clean checkout, on
// Windows as well as Unix, without pulling a toolchain in to make a zip.

const fs = require('fs');
const path = require('path');
const zlib = require('zlib');

const ROOT = path.join(__dirname, '..');
const OUT_DIR = path.join(ROOT, 'dist');

// An allow-list, not an ignore-list: anything added to the repo later has to
// be named here to reach users. Getting that backwards is how development
// files end up shipped.
const SHIP_FILES = ['manifest.json', 'background.js', 'LICENSE', 'PRIVACY.md'];
const SHIP_DIRS = ['content', 'icons', 'options', 'popup', 'styles'];

function listDir(dir) {
  const out = [];
  for (const entry of fs.readdirSync(path.join(ROOT, dir), { withFileTypes: true })) {
    const rel = `${dir}/${entry.name}`;
    if (entry.isDirectory()) out.push(...listDir(rel));
    else out.push(rel);
  }
  return out;
}

function collectFiles() {
  const files = [];
  for (const f of SHIP_FILES) {
    if (!fs.existsSync(path.join(ROOT, f))) throw new Error(`missing required file: ${f}`);
    files.push(f);
  }
  for (const d of SHIP_DIRS) {
    if (!fs.existsSync(path.join(ROOT, d))) throw new Error(`missing required directory: ${d}`);
    files.push(...listDir(d));
  }
  return files.sort();
}

// Every path the manifest points at has to be in the package. A typo here
// means an extension that installs and then breaks on the one code path that
// needed the missing file, which is a miserable thing to discover from a
// store review queue.
// The store enforces a 132-character ceiling on manifest.description, and
// rejects the upload rather than truncating. It is an easy limit to blow
// through when the description is rewritten to describe the real scope, and a
// slow way to find out -- so it fails here instead.
const DESCRIPTION_LIMIT = 132;

function checkDescription(manifest) {
  const d = manifest.description || '';
  if (!d) throw new Error('manifest.json has no description; the store requires one.');
  if (d.length > DESCRIPTION_LIMIT) {
    throw new Error(
      `manifest.json description is ${d.length} characters; the store allows ${DESCRIPTION_LIMIT}.\n` +
        `  Long-form copy belongs in the store listing, not the manifest.`
    );
  }
}

function checkManifestReferences(manifest, shipped) {
  const refs = new Set();
  const add = (v) => v && refs.add(v);

  Object.values(manifest.icons || {}).forEach(add);
  Object.values((manifest.action || {}).default_icon || {}).forEach(add);
  add((manifest.action || {}).default_popup);
  add(manifest.options_page);
  add((manifest.background || {}).service_worker);
  (manifest.content_scripts || []).forEach((cs) => {
    (cs.js || []).forEach(add);
    (cs.css || []).forEach(add);
  });

  const have = new Set(shipped);
  const missing = [...refs].filter((r) => !have.has(r));
  if (missing.length) {
    throw new Error(`manifest.json references files that are not in the package:\n  ${missing.join('\n  ')}`);
  }
  return refs.size;
}

// --- minimal zip writer -----------------------------------------------------

const CRC_TABLE = (() => {
  const table = new Int32Array(256);
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    table[n] = c;
  }
  return table;
})();

function crc32(buf) {
  let crc = -1;
  for (let i = 0; i < buf.length; i++) crc = (crc >>> 8) ^ CRC_TABLE[(crc ^ buf[i]) & 0xff];
  return (crc ^ -1) >>> 0;
}

function dosTime(date) {
  return ((date.getHours() << 11) | (date.getMinutes() << 5) | (date.getSeconds() / 2)) & 0xffff;
}

function dosDate(date) {
  return (((date.getFullYear() - 1980) << 9) | ((date.getMonth() + 1) << 5) | date.getDate()) & 0xffff;
}

function buildZip(files) {
  const chunks = [];
  const central = [];
  let offset = 0;

  for (const rel of files) {
    const full = path.join(ROOT, rel);
    const raw = fs.readFileSync(full);
    const mtime = fs.statSync(full).mtime;
    const deflated = zlib.deflateRawSync(raw);
    // Storing beats deflating when deflating made it bigger (tiny icons).
    const useDeflate = deflated.length < raw.length;
    const body = useDeflate ? deflated : raw;
    const method = useDeflate ? 8 : 0;
    const crc = crc32(raw);
    // Zip entry names always use forward slashes, whatever the host OS uses.
    const name = Buffer.from(rel.split(path.sep).join('/'), 'utf8');

    const local = Buffer.alloc(30);
    local.writeUInt32LE(0x04034b50, 0);
    local.writeUInt16LE(20, 4);
    local.writeUInt16LE(0, 6);
    local.writeUInt16LE(method, 8);
    local.writeUInt16LE(dosTime(mtime), 10);
    local.writeUInt16LE(dosDate(mtime), 12);
    local.writeUInt32LE(crc, 14);
    local.writeUInt32LE(body.length, 18);
    local.writeUInt32LE(raw.length, 22);
    local.writeUInt16LE(name.length, 26);
    local.writeUInt16LE(0, 28);
    chunks.push(local, name, body);

    const cd = Buffer.alloc(46);
    cd.writeUInt32LE(0x02014b50, 0);
    cd.writeUInt16LE(20, 4);
    cd.writeUInt16LE(20, 6);
    cd.writeUInt16LE(0, 8);
    cd.writeUInt16LE(method, 10);
    cd.writeUInt16LE(dosTime(mtime), 12);
    cd.writeUInt16LE(dosDate(mtime), 14);
    cd.writeUInt32LE(crc, 16);
    cd.writeUInt32LE(body.length, 20);
    cd.writeUInt32LE(raw.length, 24);
    cd.writeUInt16LE(name.length, 28);
    cd.writeUInt32LE(0, 38);
    cd.writeUInt32LE(offset, 42);
    central.push(cd, name);

    offset += local.length + name.length + body.length;
  }

  const centralBuf = Buffer.concat(central);
  const eocd = Buffer.alloc(22);
  eocd.writeUInt32LE(0x06054b50, 0);
  eocd.writeUInt16LE(files.length, 8);
  eocd.writeUInt16LE(files.length, 10);
  eocd.writeUInt32LE(centralBuf.length, 12);
  eocd.writeUInt32LE(offset, 16);

  return Buffer.concat([Buffer.concat(chunks), centralBuf, eocd]);
}

// --- run --------------------------------------------------------------------

function main() {
  const manifest = JSON.parse(fs.readFileSync(path.join(ROOT, 'manifest.json'), 'utf8'));
  checkDescription(manifest);
  const files = collectFiles();
  const refCount = checkManifestReferences(manifest, files);

  const zip = buildZip(files);
  fs.mkdirSync(OUT_DIR, { recursive: true });

  // The archive is named from version_name when there is one, so a beta build
  // cannot end up in a file called plain "v1.12.0". manifest.version itself
  // must stay numeric -- the browser rejects anything else -- which is why the
  // beta marker lives in version_name and gets carried through to here rather
  // than being typed into the file name by hand.
  const label = (manifest.version_name || manifest.version).replace(/[^0-9A-Za-z.]+/g, '-');
  const outPath = path.join(OUT_DIR, `D365ImportAssistant-v${label}.zip`);
  fs.writeFileSync(outPath, zip);

  const kb = (n) => `${(n / 1024).toFixed(1)} KB`;
  console.log(`Packaged ${manifest.name} ${manifest.version_name || manifest.version}`);
  if (manifest.version_name) {
    console.log(`  store version ${manifest.version}, shown to users as "${manifest.version_name}"`);
  }
  console.log(`  ${files.length} files, ${refCount} of them referenced by the manifest`);
  console.log(`  manifest.json is at the archive root (required by the store)`);
  console.log(`  ${path.relative(ROOT, outPath)} — ${kb(zip.length)}`);
}

try {
  main();
} catch (err) {
  console.error(`Packaging failed: ${err.message}`);
  process.exit(1);
}
