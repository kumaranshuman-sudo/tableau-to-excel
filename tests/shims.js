"use strict";
/* Test shims: a minimal XML DOM (enough for parseTableauFormatting) and a
 * JSZip-compatible zip reader/writer built on node:zlib. */
const zlib = require("zlib");

/* ── XML DOM ─────────────────────────────────────────────────────────────── */
function decode(s) {
  return s.replace(/&(#x[0-9a-f]+|#\d+|lt|gt|amp|quot|apos);/gi, (m, e) => {
    if (e[0] === "#") return String.fromCodePoint(e[1] === "x" || e[1] === "X" ? parseInt(e.slice(2), 16) : +e.slice(1));
    return { lt: "<", gt: ">", amp: "&", quot: '"', apos: "'" }[e.toLowerCase()];
  });
}
class Node_ {
  constructor(type, tag) { this.nodeType = type; this.tagName = tag; this.childNodes = []; this.attrs = {}; this.parentNode = null; this.data = ""; }
  getAttribute(n) { return Object.prototype.hasOwnProperty.call(this.attrs, n) ? this.attrs[n] : null; }
  hasAttribute(n) { return this.getAttribute(n) !== null; }
  get textContent() { return this.nodeType === 3 ? this.data : this.childNodes.map(c => c.textContent).join(""); }
  get children() { return this.childNodes.filter(c => c.nodeType === 1); }
  get firstChild() { return this.childNodes[0] || null; }
  get localName() { return this.tagName; }
  get nodeName() { return this.tagName; }
  getElementsByTagName(tag) {
    const out = [];
    const walk = n => n.childNodes.forEach(c => { if (c.nodeType === 1) { if (tag === "*" || c.tagName === tag) out.push(c); walk(c); } });
    walk(this);
    return out;
  }
}
function parseXml(xml) {
  const doc = new Node_(9, "#document");
  let cur = doc, i = 0;
  const re = /<!--[\s\S]*?-->|<\?[\s\S]*?\?>|<!\[CDATA\[([\s\S]*?)\]\]>|<!DOCTYPE[^>]*>|<\/([^\s>]+)\s*>|<([^\s/>]+)((?:\s+[^\s=]+\s*=\s*(?:"[^"]*"|'[^']*'))*)\s*(\/?)>/g;
  let m;
  const text = t => { if (t) { const n = new Node_(3, "#text"); n.data = decode(t); n.parentNode = cur; cur.childNodes.push(n); } };
  while ((m = re.exec(xml))) {
    text(xml.slice(i, m.index));
    i = re.lastIndex;
    if (m[1] !== undefined) { const n = new Node_(3, "#text"); n.data = m[1]; n.parentNode = cur; cur.childNodes.push(n); }
    else if (m[2]) cur = cur.parentNode || doc;
    else if (m[3]) {
      const el = new Node_(1, m[3]);
      const ar = /([^\s=]+)\s*=\s*("([^"]*)"|'([^']*)')/g;
      let a;
      while ((a = ar.exec(m[4] || ""))) el.attrs[a[1]] = decode(a[3] !== undefined ? a[3] : a[4]);
      el.parentNode = cur; cur.childNodes.push(el);
      if (!m[5]) cur = el;
    }
  }
  doc.documentElement = doc.childNodes.find(c => c.nodeType === 1);
  return doc;
}
class DOMParser { parseFromString(s) { return parseXml(s); } }

/* ── zip (JSZip subset: loadAsync, file(), generateAsync) ────────────────── */
function readZip(buf) {
  buf = Buffer.from(buf);
  let eocd = buf.length - 22;
  while (eocd >= 0 && buf.readUInt32LE(eocd) !== 0x06054b50) eocd--;
  const n = buf.readUInt16LE(eocd + 10);
  let p = buf.readUInt32LE(eocd + 16);
  const files = {};
  for (let k = 0; k < n; k++) {
    const method = buf.readUInt16LE(p + 10), csize = buf.readUInt32LE(p + 20);
    const nlen = buf.readUInt16LE(p + 28), elen = buf.readUInt16LE(p + 30), clen = buf.readUInt16LE(p + 32);
    const lho = buf.readUInt32LE(p + 42);
    const name = buf.toString("utf8", p + 46, p + 46 + nlen);
    const lnlen = buf.readUInt16LE(lho + 26), lelen = buf.readUInt16LE(lho + 28);
    const raw = buf.subarray(lho + 30 + lnlen + lelen, lho + 30 + lnlen + lelen + csize);
    if (!name.endsWith("/")) files[name] = method === 8 ? zlib.inflateRawSync(raw) : Buffer.from(raw);
    p += 46 + nlen + elen + clen;
  }
  return files;
}
function writeZip(files) {
  const locals = [], centrals = [];
  let off = 0;
  for (const [name, data0] of Object.entries(files)) {
    const data = Buffer.isBuffer(data0) ? data0 : Buffer.from(data0, "utf8");
    const comp = zlib.deflateRawSync(data);
    const nameB = Buffer.from(name, "utf8");
    const crc = zlib.crc32(data);
    const lh = Buffer.alloc(30);
    lh.writeUInt32LE(0x04034b50, 0); lh.writeUInt16LE(20, 4); lh.writeUInt16LE(0x0800, 6); lh.writeUInt16LE(8, 8);
    lh.writeUInt32LE(crc, 14); lh.writeUInt32LE(comp.length, 18); lh.writeUInt32LE(data.length, 22); lh.writeUInt16LE(nameB.length, 26);
    const ch = Buffer.alloc(46);
    ch.writeUInt32LE(0x02014b50, 0); ch.writeUInt16LE(20, 4); ch.writeUInt16LE(20, 6); ch.writeUInt16LE(0x0800, 8); ch.writeUInt16LE(8, 10);
    ch.writeUInt32LE(crc, 16); ch.writeUInt32LE(comp.length, 20); ch.writeUInt32LE(data.length, 24); ch.writeUInt16LE(nameB.length, 28);
    ch.writeUInt32LE(off, 42);
    locals.push(lh, nameB, comp);
    centrals.push(ch, nameB);
    off += 30 + nameB.length + comp.length;
  }
  const cd = Buffer.concat(centrals);
  const end = Buffer.alloc(22);
  const count = Object.keys(files).length;
  end.writeUInt32LE(0x06054b50, 0); end.writeUInt16LE(count, 8); end.writeUInt16LE(count, 10);
  end.writeUInt32LE(cd.length, 12); end.writeUInt32LE(off, 16);
  return Buffer.concat([...locals, cd, end]);
}
const JSZip = {
  async loadAsync(buf) {
    const files = readZip(buf);
    return {
      file(name, content) {
        if (content === undefined) return files[name] ? { async: async () => files[name].toString("utf8") } : null;
        files[name] = Buffer.from(content, "utf8");
        return this;
      },
      async generateAsync() { return new Uint8Array(writeZip(files)); }
    };
  }
};

module.exports = { DOMParser, JSZip, writeZip, readZip };
