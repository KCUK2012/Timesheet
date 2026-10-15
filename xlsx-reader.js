/*
 * Outlook Timesheet: minimal reader for .xlsx workbooks.
 * Reads cell values only (text, numbers, true/false). No formatting, no formulas.
 * Written for this add-in so that no third-party spreadsheet library is needed.
 * Uses the browser's built-in DecompressionStream ("deflate-raw") and DOMParser.
 */
(function (root) {
  "use strict";

  function u16(dv, o) { return dv.getUint16(o, true); }
  function u32(dv, o) { return dv.getUint32(o, true); }

  /** Lists the files inside a zip archive: { name: { method, offset, size } }. */
  function readZipDirectory(buf) {
    var dv = new DataView(buf);
    var eocd = -1;
    for (var i = buf.byteLength - 22; i >= Math.max(0, buf.byteLength - 65557); i--) {
      if (u32(dv, i) === 0x06054b50) { eocd = i; break; }
    }
    if (eocd < 0) throw new Error("The workbook is not a valid .xlsx file.");
    var count = u16(dv, eocd + 10);
    var p = u32(dv, eocd + 16);
    var decoder = new TextDecoder("utf-8");
    var entries = {};
    for (var n = 0; n < count; n++) {
      if (u32(dv, p) !== 0x02014b50) throw new Error("The workbook's zip directory is damaged.");
      var method = u16(dv, p + 10);
      var compSize = u32(dv, p + 20);
      var nameLen = u16(dv, p + 28), extraLen = u16(dv, p + 30), commentLen = u16(dv, p + 32);
      var local = u32(dv, p + 42);
      var name = decoder.decode(new Uint8Array(buf, p + 46, nameLen));
      entries[name] = { method: method, local: local, size: compSize };
      p += 46 + nameLen + extraLen + commentLen;
    }
    return entries;
  }

  async function readEntry(buf, entry) {
    var dv = new DataView(buf);
    if (u32(dv, entry.local) !== 0x04034b50) throw new Error("The workbook's zip data is damaged.");
    var start = entry.local + 30 + u16(dv, entry.local + 26) + u16(dv, entry.local + 28);
    var bytes = new Uint8Array(buf, start, entry.size);
    if (entry.method === 0) return new TextDecoder("utf-8").decode(bytes);
    if (entry.method !== 8) throw new Error("The workbook uses an unsupported compression method.");
    if (typeof DecompressionStream === "undefined") {
      throw new Error("This version of Outlook cannot open .xlsx files. Please update Outlook.");
    }
    var stream = new Blob([bytes]).stream().pipeThrough(new DecompressionStream("deflate-raw"));
    return new Response(stream).text();
  }

  function parseXml(text) {
    var doc = new DOMParser().parseFromString(text, "application/xml");
    if (doc.getElementsByTagName("parsererror").length) throw new Error("The workbook contains unreadable XML.");
    return doc;
  }
  function byTag(node, tag) { return Array.prototype.slice.call(node.getElementsByTagNameNS("*", tag)); }
  function childByTag(node, tag) {
    for (var c = node.firstElementChild; c; c = c.nextElementSibling) if (c.localName === tag) return c;
    return null;
  }
  /** Text of a shared or inline string, ignoring phonetic runs. */
  function stringText(si) {
    var out = "";
    byTag(si, "t").forEach(function (t) {
      if (!t.parentNode || t.parentNode.localName !== "rPh") out += t.textContent;
    });
    return out;
  }
  function colIndex(ref) {
    var letters = String(ref || "").replace(/[^A-Z]/gi, "").toUpperCase(), n = 0;
    for (var i = 0; i < letters.length; i++) n = n * 26 + (letters.charCodeAt(i) - 64);
    return n - 1;
  }
  function resolveTarget(target) {
    target = String(target || "");
    if (target.charAt(0) === "/") return target.slice(1);
    return "xl/" + target.replace(/^\.\//, "");
  }

  /**
   * Reads a workbook. Returns { "Sheet name": [[cell, cell, ...], ...], ... }
   * with every cell as a trimmed string.
   */
  async function read(buf) {
    var entries = readZipDirectory(buf);
    async function file(name) {
      return entries[name] ? readEntry(buf, entries[name]) : null;
    }
    var wbXml = await file("xl/workbook.xml");
    var relsXml = await file("xl/_rels/workbook.xml.rels");
    if (!wbXml || !relsXml) throw new Error("The workbook is not a valid .xlsx file.");
    var rels = {};
    byTag(parseXml(relsXml), "Relationship").forEach(function (r) {
      rels[r.getAttribute("Id")] = resolveTarget(r.getAttribute("Target"));
    });
    var shared = [];
    var ssXml = await file("xl/sharedStrings.xml");
    if (ssXml) shared = byTag(parseXml(ssXml), "si").map(stringText);

    var result = {};
    var sheets = byTag(parseXml(wbXml), "sheet");
    for (var s = 0; s < sheets.length; s++) {
      var sh = sheets[s];
      var rid = sh.getAttribute("r:id") ||
        sh.getAttributeNS("http://schemas.openxmlformats.org/officeDocument/2006/relationships", "id");
      var path = rels[rid];
      var xml = path ? await file(path) : null;
      if (!xml) continue;
      var rows = [];
      byTag(parseXml(xml), "row").forEach(function (rowEl) {
        var row = [];
        var next = 0;
        byTag(rowEl, "c").forEach(function (c) {
          var ref = c.getAttribute("r");
          var idx = ref ? colIndex(ref) : next;
          next = idx + 1;
          var t = c.getAttribute("t");
          var v = childByTag(c, "v");
          var val = "";
          if (t === "s") val = v ? (shared[parseInt(v.textContent, 10)] || "") : "";
          else if (t === "inlineStr") { var is = childByTag(c, "is"); val = is ? stringText(is) : ""; }
          else if (t === "b") val = v && v.textContent === "1" ? "TRUE" : "FALSE";
          else val = v ? v.textContent : "";
          row[idx] = String(val).trim();
        });
        for (var k = 0; k < row.length; k++) if (row[k] === undefined) row[k] = "";
        rows.push(row);
      });
      result[sh.getAttribute("name")] = rows;
    }
    return result;
  }

  root.KCXlsx = { read: read, colIndex: colIndex };
})(window);
