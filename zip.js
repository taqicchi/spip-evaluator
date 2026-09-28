/* Pembaca/penulis ZIP minimal (tanpa pustaka luar).
 * - Membaca direktori pusat saja; isi berkas diekstrak saat diminta (hemat memori untuk ZIP bukti ratusan MB).
 * - Menulis ulang ZIP dengan MENYALIN byte terkompresi asli untuk entri yang tidak diubah,
 *   sehingga berkas xlsx hasil tetap identik kecuali bagian yang diganti.
 * Memakai CompressionStream/DecompressionStream('deflate-raw') bawaan peramban (juga ada di Node 18+).
 */
(function (root) {
  'use strict';

  const td = new TextDecoder('utf-8');
  const te = new TextEncoder();

  // blob: Blob/File (peramban) atau objek {size, slice(a,b)->{arrayBuffer()}}
  async function readBytes(blob, start, end) {
    return new Uint8Array(await blob.slice(start, end).arrayBuffer());
  }
  const u16 = (b, o) => b[o] | (b[o + 1] << 8);
  const u32 = (b, o) => (b[o] | (b[o + 1] << 8) | (b[o + 2] << 16) | (b[o + 3] << 24)) >>> 0;

  async function open(blob) {
    const size = blob.size;
    const tailLen = Math.min(size, 65557);
    const tail = await readBytes(blob, size - tailLen, size);
    let eocd = -1;
    for (let i = tail.length - 22; i >= 0; i--) if (u32(tail, i) === 0x06054b50) { eocd = i; break; }
    if (eocd < 0) throw new Error('Bukan berkas ZIP yang valid');
    let count = u16(tail, eocd + 10);
    let cdSize = u32(tail, eocd + 12);
    let cdOff = u32(tail, eocd + 16);
    // ZIP64
    if (cdOff === 0xffffffff || count === 0xffff) {
      const loc = eocd - 20;
      if (loc >= 0 && u32(tail, loc) === 0x07064b50) {
        const z64off = Number(new DataView(tail.buffer, tail.byteOffset + loc + 8, 8).getBigUint64(0, true));
        const z = await readBytes(blob, z64off, z64off + 56);
        const dv = new DataView(z.buffer, z.byteOffset);
        count = Number(dv.getBigUint64(32, true));
        cdSize = Number(dv.getBigUint64(40, true));
        cdOff = Number(dv.getBigUint64(48, true));
      }
    }
    const cd = await readBytes(blob, cdOff, cdOff + cdSize);
    const entries = [];
    let p = 0;
    for (let i = 0; i < count && p + 46 <= cd.length; i++) {
      if (u32(cd, p) !== 0x02014b50) break;
      const flags = u16(cd, p + 8), method = u16(cd, p + 10);
      const time = u16(cd, p + 12), date = u16(cd, p + 14);
      const crc = u32(cd, p + 16);
      let csize = u32(cd, p + 20), usize = u32(cd, p + 24);
      const nlen = u16(cd, p + 28), xlen = u16(cd, p + 30), clen = u16(cd, p + 32);
      let off = u32(cd, p + 42);
      const nameBytes = cd.subarray(p + 46, p + 46 + nlen);
      // extra ZIP64
      if (csize === 0xffffffff || usize === 0xffffffff || off === 0xffffffff) {
        let x = p + 46 + nlen;
        const xe = x + xlen;
        while (x + 4 <= xe) {
          const id = u16(cd, x), sz = u16(cd, x + 2);
          if (id === 1) {
            const dv = new DataView(cd.buffer, cd.byteOffset + x + 4, sz);
            let k = 0;
            if (usize === 0xffffffff) { usize = Number(dv.getBigUint64(k, true)); k += 8; }
            if (csize === 0xffffffff) { csize = Number(dv.getBigUint64(k, true)); k += 8; }
            if (off === 0xffffffff) { off = Number(dv.getBigUint64(k, true)); }
          }
          x += 4 + sz;
        }
      }
      entries.push({ name: td.decode(nameBytes), nameBytes: nameBytes.slice(), flags, method, time, date, crc, csize, usize, offset: off, dir: nameBytes[nlen - 1] === 47 });
      p += 46 + nlen + xlen + clen;
    }
    return { blob, entries, find: n => entries.find(e => e.name === n || e.name === n.replace(/^\//, '')) };
  }

  async function rawData(zip, e) {
    const h = await readBytes(zip.blob, e.offset, e.offset + 30);
    if (u32(h, 0) !== 0x04034b50) throw new Error('Header lokal ZIP rusak: ' + e.name);
    const start = e.offset + 30 + u16(h, 26) + u16(h, 28);
    return readBytes(zip.blob, start, start + e.csize);
  }

  async function streamAll(stream) {
    const chunks = [];
    let n = 0;
    const r = stream.getReader();
    for (;;) { const { done, value } = await r.read(); if (done) break; chunks.push(value); n += value.length; }
    const out = new Uint8Array(n);
    let o = 0;
    for (const c of chunks) { out.set(c, o); o += c.length; }
    return out;
  }
  async function inflate(bytes) {
    return streamAll(new Blob([bytes]).stream().pipeThrough(new DecompressionStream('deflate-raw')));
  }
  async function deflate(bytes) {
    return streamAll(new Blob([bytes]).stream().pipeThrough(new CompressionStream('deflate-raw')));
  }

  async function extract(zip, e) {
    const raw = await rawData(zip, e);
    if (e.method === 0) return raw;
    if (e.method === 8) return inflate(raw);
    throw new Error(`Metode kompresi ${e.method} tidak didukung (${e.name})`);
  }

  let CRC_T = null;
  function crc32(b) {
    if (!CRC_T) {
      CRC_T = new Uint32Array(256);
      for (let n = 0; n < 256; n++) { let c = n; for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1; CRC_T[n] = c >>> 0; }
    }
    let c = 0xffffffff;
    for (let i = 0; i < b.length; i++) c = CRC_T[(c ^ b[i]) & 0xff] ^ (c >>> 8);
    return (c ^ 0xffffffff) >>> 0;
  }
  function dosNow() {
    const d = new Date();
    return {
      time: (d.getHours() << 11) | (d.getMinutes() << 5) | (d.getSeconds() >> 1),
      date: ((d.getFullYear() - 1980) << 9) | ((d.getMonth() + 1) << 5) | d.getDate(),
    };
  }

  // replace: Map<nama entri, Uint8Array isi baru (tidak terkompresi)>
  async function rewrite(zip, replace) {
    const parts = [], central = [];
    let off = 0;
    const now = dosNow();
    for (const e of zip.entries) {
      let data, method = e.method, crc = e.crc, usize = e.usize, time = e.time, date = e.date;
      const flags = e.flags & 0x0800; // pertahankan penanda UTF-8, buang data descriptor
      if (replace.has(e.name)) {
        const plain = replace.get(e.name);
        data = await deflate(plain);
        method = 8; crc = crc32(plain); usize = plain.length; time = now.time; date = now.date;
      } else {
        data = await rawData(zip, e);
      }
      if (data.length >= 0xffffffff || usize >= 0xffffffff || off >= 0xffffffff) throw new Error('Berkas terlalu besar (ZIP64 tidak didukung untuk penulisan)');
      const name = e.nameBytes;
      const lh = new Uint8Array(30 + name.length);
      const lv = new DataView(lh.buffer);
      lv.setUint32(0, 0x04034b50, true); lv.setUint16(4, 20, true); lv.setUint16(6, flags, true); lv.setUint16(8, method, true);
      lv.setUint16(10, time, true); lv.setUint16(12, date, true); lv.setUint32(14, crc, true);
      lv.setUint32(18, data.length, true); lv.setUint32(22, usize, true); lv.setUint16(26, name.length, true); lv.setUint16(28, 0, true);
      lh.set(name, 30);
      const ch = new Uint8Array(46 + name.length);
      const cv = new DataView(ch.buffer);
      cv.setUint32(0, 0x02014b50, true); cv.setUint16(4, 20, true); cv.setUint16(6, 20, true); cv.setUint16(8, flags, true);
      cv.setUint16(10, method, true); cv.setUint16(12, time, true); cv.setUint16(14, date, true); cv.setUint32(16, crc, true);
      cv.setUint32(20, data.length, true); cv.setUint32(24, usize, true); cv.setUint16(28, name.length, true);
      cv.setUint32(42, off, true);
      ch.set(name, 46);
      parts.push(lh, data);
      central.push(ch);
      off += lh.length + data.length;
    }
    let cdSize = 0;
    for (const c of central) cdSize += c.length;
    const end = new Uint8Array(22);
    const ev = new DataView(end.buffer);
    ev.setUint32(0, 0x06054b50, true);
    ev.setUint16(8, central.length, true); ev.setUint16(10, central.length, true);
    ev.setUint32(12, cdSize, true); ev.setUint32(16, off, true);
    return new Blob([...parts, ...central, end], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  }

  const api = { open, extract, rewrite, crc32, decode: b => td.decode(b), encode: s => te.encode(s) };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.MiniZip = api;
})(this);
