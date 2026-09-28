/* Menulis nilai ke sel tertentu pada berkas .xlsx asli tanpa mengubah bagian lain:
 * style, merge, validasi data, rumus, dan sheet lain tetap seperti semula.
 * Bergantung pada MiniZip (zip.js).
 */
(function (root) {
  'use strict';
  const Z = root.MiniZip || (typeof require !== 'undefined' ? require('./zip.js') : null);

  function xmlEsc(s) {
    return String(s)
      .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '')
      .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  }
  function attrUnesc(s) {
    return s.replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&amp;/g, '&');
  }
  function colNum(letters) {
    let n = 0;
    for (const ch of letters) n = n * 26 + (ch.charCodeAt(0) - 64);
    return n;
  }
  // cari tag pembuka "<tag attr="val"" diikuti spasi, '>' atau '/'
  function findTag(xml, prefix, from, until) {
    let i = from;
    for (;;) {
      i = xml.indexOf(prefix, i);
      if (i < 0 || (until != null && i >= until)) return -1;
      const ch = xml[i + prefix.length];
      if (ch === ' ' || ch === '>' || ch === '/' || ch === '\t' || ch === '\n' || ch === '\r') return i;
      i += prefix.length;
    }
  }

  function setCell(xml, ref, text) {
    const m = /^([A-Z]+)(\d+)$/.exec(ref);
    if (!m) throw new Error('Alamat sel tidak valid: ' + ref);
    const col = m[1], row = +m[2];
    const cellXml = s => `<c r="${ref}"${s ? ` s="${s}"` : ''} t="inlineStr"><is><t xml:space="preserve">${xmlEsc(text)}</t></is></c>`;

    const sdStart = xml.indexOf('<sheetData');
    let rs = findTag(xml, `<row r="${row}"`, sdStart);
    if (rs < 0) {
      // sisipkan baris baru sesuai urutan
      const sdEnd = xml.indexOf('</sheetData>');
      if (sdEnd < 0) throw new Error('sheetData tidak ditemukan');
      const rx = /<row r="(\d+)"/g;
      rx.lastIndex = sdStart;
      let at = sdEnd, mm;
      while ((mm = rx.exec(xml)) && mm.index < sdEnd) { if (+mm[1] > row) { at = mm.index; break; } }
      return xml.slice(0, at) + `<row r="${row}">${cellXml('')}</row>` + xml.slice(at);
    }
    const openEnd = xml.indexOf('>', rs);
    if (xml[openEnd - 1] === '/') {
      // <row .../> -> <row ...>cell</row>
      return xml.slice(0, openEnd - 1) + '>' + cellXml('') + '</row>' + xml.slice(openEnd + 1);
    }
    const re = xml.indexOf('</row>', openEnd);
    const cs = findTag(xml, `<c r="${ref}"`, openEnd, re);
    if (cs >= 0) {
      const tagEnd = xml.indexOf('>', cs);
      const selfClose = xml[tagEnd - 1] === '/';
      const ce = selfClose ? tagEnd + 1 : xml.indexOf('</c>', tagEnd) + 4;
      const sm = /\ss="(\d+)"/.exec(xml.slice(cs, tagEnd));
      return xml.slice(0, cs) + cellXml(sm ? sm[1] : '') + xml.slice(ce);
    }
    // sel belum ada: sisipkan sesuai urutan kolom
    const target = colNum(col);
    const rx = /<c r="([A-Z]+)\d+"/g;
    rx.lastIndex = openEnd;
    let at = re, mm;
    while ((mm = rx.exec(xml)) && mm.index < re) { if (colNum(mm[1]) > target) { at = mm.index; break; } }
    return xml.slice(0, at) + cellXml('') + xml.slice(at);
  }

  function ensureFullCalc(wbXml) {
    if (/<calcPr\b[^>]*fullCalcOnLoad="(1|true)"/.test(wbXml)) return wbXml;
    if (/<calcPr\b/.test(wbXml)) return wbXml.replace(/<calcPr\b/, '<calcPr fullCalcOnLoad="1"');
    const after = ['</definedNames>', '</externalReferences>', '</functionGroups>', '</sheets>'].find(t => wbXml.includes(t));
    return wbXml.replace(after, after + '<calcPr fullCalcOnLoad="1"/>');
  }

  async function sheetPath(zip, sheetName) {
    const wb = Z.decode(await Z.extract(zip, zip.find('xl/workbook.xml')));
    const rels = Z.decode(await Z.extract(zip, zip.find('xl/_rels/workbook.xml.rels')));
    let rid = null;
    for (const m of wb.matchAll(/<sheet\b[^>]*>/g)) {
      const name = /\bname="([^"]*)"/.exec(m[0]);
      if (name && attrUnesc(name[1]) === sheetName) { rid = /\br:id="([^"]*)"/.exec(m[0]) || /\bid="([^"]*)"/.exec(m[0]); rid = rid && rid[1]; break; }
    }
    if (!rid) throw new Error(`Sheet "${sheetName}" tidak ditemukan di workbook.xml`);
    let target = null;
    for (const m of rels.matchAll(/<Relationship\b[^>]*>/g)) {
      const id = /\bId="([^"]*)"/.exec(m[0]);
      if (id && id[1] === rid) { target = /\bTarget="([^"]*)"/.exec(m[0])[1]; break; }
    }
    if (!target) throw new Error('Relasi sheet tidak ditemukan');
    const path = target.startsWith('/') ? target.slice(1) : 'xl/' + target.replace(/^\.\//, '');
    return { path, wbXml: wb };
  }

  // cells: [{ref:'N9', value:'B'}]; mengembalikan Blob xlsx baru
  async function patch(file, sheetName, cells) {
    const zip = await Z.open(file);
    const { path, wbXml } = await sheetPath(zip, sheetName);
    const entry = zip.find(path);
    if (!entry) throw new Error('Berkas sheet tidak ditemukan: ' + path);
    let xml = Z.decode(await Z.extract(zip, entry));
    for (const c of cells) xml = setCell(xml, c.ref, c.value);
    const repl = new Map([[entry.name, Z.encode(xml)]]);
    const wb2 = ensureFullCalc(wbXml);
    if (wb2 !== wbXml) repl.set(zip.find('xl/workbook.xml').name, Z.encode(wb2));
    return Z.rewrite(zip, repl);
  }

  const api = { patch, setCell };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.XlsxPatch = api;
})(this);
