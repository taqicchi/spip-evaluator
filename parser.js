/* Parser & pemeriksa otomatis Kertas Kerja (KK) Evaluasi SPIP.
 * Bekerja di peramban (window.SpipParser) maupun Node (module.exports).
 * Input: workbook SheetJS yang dibaca dengan {dense:true, cellFormula:true, sheetStubs:true}.
 */
(function (root) {
  'use strict';

  const STAGES = ['PM', 'PK', 'EV'];
  const STAGE_LABEL = { PM: 'PM', PK: 'PK', EV: 'Evaluasi' };
  const LEVELS = ['A', 'B', 'C', 'D', 'E'];
  const SCORE = { A: 5, B: 4, C: 3, D: 2, E: 1 };

  // ---------- util sel ----------
  function colName(c) {
    let s = '';
    c += 1;
    while (c > 0) { const m = (c - 1) % 26; s = String.fromCharCode(65 + m) + s; c = Math.floor((c - 1) / 26); }
    return s;
  }
  function addr(r, c) { return colName(c) + (r + 1); }
  function norm(s) { return String(s || '').replace(/\s+/g, ' ').trim(); }
  function low(s) { return norm(s).toLowerCase(); }

  class Grid {
    constructor(ws) {
      this.data = ws['!data'] || [];
      this.nrows = this.data.length;
      let nc = 0;
      for (const row of this.data) if (row && row.length > nc) nc = row.length;
      this.ncols = nc;
      this.mergeTL = new Map();
      this.merges = ws['!merges'] || [];
      for (const m of this.merges) {
        const size = (m.e.r - m.s.r + 1) * (m.e.c - m.s.c + 1);
        if (size <= 1 || size > 5000) continue;
        for (let r = m.s.r; r <= m.e.r; r++)
          for (let c = m.s.c; c <= m.e.c; c++)
            if (r !== m.s.r || c !== m.s.c) this.mergeTL.set(r * 16384 + c, [m.s.r, m.s.c]);
      }
    }
    cell(r, c) { const row = this.data[r]; return row ? row[c] : undefined; }
    isFormula(r, c) { const x = this.cell(r, c); return !!(x && x.f); }
    // teks nilai (tanpa resolusi merge). Sel rumus tanpa nilai tersimpan -> ''
    text(r, c) {
      const x = this.cell(r, c);
      if (!x || x.t === 'z' || x.t === 'e') return '';
      const v = x.w != null ? x.w : x.v;
      return v == null ? '' : String(v).trim();
    }
    num(r, c) {
      const x = this.cell(r, c);
      if (!x || x.t === 'z') return null;
      if (x.t === 'n') return x.v;
      const n = parseFloat(String(x.v).replace(/\./g, '').replace(',', '.'));
      return isFinite(n) ? n : null;
    }
    // teks dengan resolusi merge (untuk header & sel gabungan)
    mtext(r, c) {
      const tl = this.mergeTL.get(r * 16384 + c);
      return tl ? this.text(tl[0], tl[1]) : this.text(r, c);
    }
    mtop(r, c) { const tl = this.mergeTL.get(r * 16384 + c); return tl ? tl : [r, c]; }
    rowHasText(r, c0, c1) {
      for (let c = c0; c <= c1; c++) if (this.text(r, c)) return true;
      return false;
    }
  }

  // ---------- klasifikasi isi teks ----------
  const PLACEHOLDER_RX = /(bahwa|telah|uraian aoi|uraian penyebab|kluster|klaster|[.…\s|\-:])/gi;
  function isPlaceholder(s) {
    const t = norm(s);
    if (!t) return true;
    return t.replace(PLACEHOLDER_RX, '').length < 4;
  }
  function isReal(s) { return !isPlaceholder(s); }
  const NEG_RX = /\b(belum|tidak\s+(ada|terdapat|dilakukan|dilaksanakan|tersedia|memadai|lengkap|sesuai|seluruh|semua)|kurang\s+(memadai|optimal)|masih\s+(belum|kurang))\b/i;
  function negSnippet(s) {
    const m = NEG_RX.exec(s);
    if (!m) return null;
    const i = Math.max(0, m.index - 50);
    return (i > 0 ? '…' : '') + s.slice(i, m.index + m[0].length + 60).replace(/\s+/g, ' ') + '…';
  }

  function flag(level, code, msg, cell, extra) {
    return Object.assign({ level, code, msg, cell: cell || '' }, extra || {});
  }

  // ---------- deteksi jenis sheet ----------
  function findText(g, rx, r0, r1, c0, c1) {
    for (let r = r0; r <= Math.min(r1, g.nrows - 1); r++)
      for (let c = c0; c <= Math.min(c1, g.ncols - 1); c++)
        if (rx.test(g.text(r, c))) return [r, c];
    return null;
  }

  function findIndexRow(g, maxRow) {
    // baris penomoran kolom: kolom A = 1 dan kolom berikutnya = 2 atau rumus (=A7+1)
    for (let r = 0; r < Math.min(maxRow, g.nrows); r++) {
      const a = g.cell(r, 0);
      if (!a || !(g.num(r, 0) === 1 || (a.f && /^0?\+?1$/.test(a.f)))) continue;
      for (let c = 1; c <= 3; c++) {
        const x = g.cell(r, c);
        if (x && ((x.f && /^\$?[A-Z]{1,2}\$?\d+\+1$/.test(x.f)) || x.v === 2 || x.v === '2')) return r;
      }
    }
    return -1;
  }

  function classify(g) {
    if (findText(g, /^uraian parameter$/i, 0, 15, 0, 10) && findText(g, /^uraian hasil pengujian/i, 0, 15, 0, 60)) return 'sp';
    if (findText(g, /PENYIMPULAN NILAI MATURITAS/i, 0, 10, 0, 10) && findText(g, /^skor$/i, 0, 12, 0, 30)) return 'lead';
    if (findIndexRow(g, 20) >= 0) return 'table';
    return 'info';
  }

  // ---------- Struktur & Proses (KK 3.x) ----------
  function kodeText(g, r, c, sheetFlags) {
    const x = g.cell(r, c);
    if (!x) return '';
    let t = g.text(r, c);
    if (x.t === 'n' && x.v > 1000) {
      // "1.8" yang terbaca Excel sebagai tanggal 1 Agustus
      if (!/^\d+\.\d+$/.test(t)) {
        const d = new Date(Math.round((x.v - 25569) * 86400000));
        t = d.getUTCDate() + '.' + (d.getUTCMonth() + 1);
      }
      sheetFlags.push(flag('info', 'KODE_TANGGAL', `Kode subunsur "${t}" tersimpan sebagai tanggal (${x.v}). Ketik ulang sebagai teks agar tidak berubah.`, addr(r, c)));
    }
    return t;
  }

  function parseSP(g, name) {
    const sheetFlags = [];
    const hp = findText(g, /^uraian parameter$/i, 0, 15, 0, 10);
    const hdr = hp[0];
    const idx = findIndexRow(g, 20);
    const band0 = hdr, band1 = idx > hdr ? idx - 1 : hdr + 3;
    const find = (rx, from) => { const p = findText(g, rx, hdr, hdr + 1, from || 0, 20); return p ? p[1] : -1; };
    const col = {
      kode: find(/^kode$/i), sub: find(/^uraian subunsur$/i), no: find(/^no$/i), param: hp[1],
      kp: find(/^kode parameter$/i), lvl: find(/^grad/i), krit: find(/^kriteria$/i),
      penj: find(/^penjelasan$/i), cara: find(/^cara pengujian$/i),
    };
    const kpNames = [];
    if (col.kp >= 0) for (let c = col.kp; c < col.kp + 3; c++) kpNames.push({ c, n: g.text(hdr + 1, c) || '' });

    // --- blok isian (per satker / per tahap) ---
    const startC = Math.max(col.cara, col.krit) + 1;
    const uraianCols = [];
    for (let r = band0; r <= band1; r++)
      for (let c = startC; c < g.ncols; c++)
        if (/^uraian hasil pengujian/i.test(g.text(r, c))) uraianCols.push([r, c]);
    uraianCols.sort((a, b) => a[1] - b[1]);

    const blocks = [];
    uraianCols.forEach(([fr, uc], i) => {
      const nextU = i + 1 < uraianCols.length ? uraianCols[i + 1][1] : g.ncols;
      const b = { fr, uraian: uc, grade: {}, gradeSingle: -1, aoiK: -1, aoiU: -1, sebabK: -1, sebabU: -1, kes: null };
      for (let c = uc + 1; c < nextU; c++) {
        const h = low(g.text(fr, c));
        if (!h) break;
        if (/^grade (pm)$/.test(h)) b.grade.PM = c;
        else if (/^grade pk$/.test(h)) b.grade.PK = c;
        else if (/^grade eval/.test(h)) b.grade.EV = c;
        else if (/^grade$/.test(h)) b.gradeSingle = c;
        else if (/^(kl[ua]ster) aoi/.test(h)) b.aoiK = c;
        else if (/^uraian aoi/.test(h)) b.aoiU = c;
        else if (/^(kl[ua]ster) penyebab/.test(h)) b.sebabK = c;
        else if (/^uraian penyebab/.test(h)) b.sebabU = c;
        else if (/^kesimpulan akhir/.test(h)) { b.kes = { c, stage: /pm/.test(h) ? 'PM' : /pk/.test(h) ? 'PK' : 'EV' }; break; }
        else break;
      }
      // label blok: teks header di atas kolom uraian
      const lab = [];
      for (let r = band0; r < fr; r++) {
        const t = g.mtext(r, uc);
        if (t && !/^(hasil pengujian terkait|indeks kk|:)/i.test(t) && !lab.includes(t)) lab.push(t);
      }
      b.label = lab.join(' — ');
      blocks.push(b);
    });

    // kesimpulan global (gaya KK3.1)
    const globalKes = {};
    for (let r = band0; r <= band1; r++)
      for (let c = startC; c < g.ncols; c++) {
        const h = low(g.text(r, c));
        const m = /^kesimpulan akhir (pm|pk|evaluasi)$/.exec(h);
        if (m && !blocks.some(b => b.kes && b.kes.c === c)) globalKes[m[1] === 'pm' ? 'PM' : m[1] === 'pk' ? 'PK' : 'EV'] = c;
      }

    // --- susun unit ---
    const units = [];
    const multi = blocks.filter(b => Object.keys(b.grade).length > 0);
    const single = blocks.filter(b => Object.keys(b.grade).length === 0 && b.gradeSingle >= 0);
    multi.forEach((b, i) => {
      const st = {};
      for (const s of STAGES) if (b.grade[s] != null) st[s] = { grade: b.grade[s], uraian: b.uraian, aoiK: b.aoiK, aoiU: b.aoiU, sebabK: b.sebabK, sebabU: b.sebabU };
      units.push({ id: 'u' + i, label: b.label.replace(/sektor\/fokus:\s*$/i, '').replace(/\s—\s*$/, '') || 'Satker ' + (i + 1), sharedUraian: true, stages: st });
    });
    if (single.length) {
      const st = {};
      single.forEach((b, i) => {
        const s = b.kes ? b.kes.stage : STAGES[i] || ('S' + i);
        st[s] = { grade: b.gradeSingle, uraian: b.uraian, aoiK: b.aoiK, aoiU: b.aoiU, sebabK: b.sebabK, sebabU: b.sebabU, kes: b.kes ? b.kes.c : -1 };
      });
      units.push({ id: 'u' + units.length, label: single[0].label || 'Asesor', sharedUraian: false, stages: st });
    }

    // --- baris data ---
    const subs = [];
    let sub = null, param = null;
    const dataStart = idx >= 0 ? idx + 1 : hdr + 3;
    for (let r = dataStart; r < g.nrows; r++) {
      const lv = g.text(r, col.lvl).toUpperCase();
      const subTxt = col.sub >= 0 ? g.text(r, col.sub) : '';
      const noTxt = col.no >= 0 ? g.text(r, col.no) : '';
      if (/^petunjuk/i.test(subTxt) || /^petunjuk/i.test(g.text(r, 0))) break;
      if (!LEVELS.includes(lv)) {
        if (subTxt && !noTxt) {
          sub = { kode: kodeText(g, r, col.kode, sheetFlags), nama: subTxt, row: r, params: [] };
          subs.push(sub); param = null;
        }
        continue;
      }
      if (!sub) { sub = { kode: '?', nama: '(tanpa subunsur)', row: r, params: [] }; subs.push(sub); }
      if (noTxt || !param || lv === 'A' && param.levels.some(l => l.lvl === 'A')) {
        const kode = kpNames.map(k => ({ n: k.n, v: g.text(r, k.c) })).filter(k => k.v && k.v !== '-').map(k => k.n || k.v);
        param = { sheet: name, row: r, no: noTxt, uraian: g.mtext(r, col.param), kodeParam: kode, sub, levels: [] };
        sub.params.push(param);
      }
      param.levels.push({
        lvl: lv, row: r,
        kriteria: col.krit >= 0 ? g.text(r, col.krit) : '',
        penjelasan: col.penj >= 0 ? g.text(r, col.penj) : '',
        cara: col.cara >= 0 ? g.text(r, col.cara) : '',
      });
    }

    const params = [];
    for (const s of subs) for (const p of s.params) params.push(p);

    // --- isi nilai per parameter/unit/tahap ---
    for (const p of params) {
      const r0 = p.row;
      p.key = `${name}!${r0}`;
      p.units = units.map(u => {
        const res = { unit: u, stages: {} };
        for (const s of Object.keys(u.stages)) {
          const d = u.stages[s];
          const gt = g.mtext(r0, d.grade).toUpperCase();
          res.stages[s] = {
            grade: gt, gradeCell: addr(r0, d.grade),
            uraian: p.levels.map(l => ({ lvl: l.lvl, text: g.mtext(l.row, d.uraian), cell: addr(l.row, d.uraian) })),
            aoiK: d.aoiK >= 0 ? g.mtext(r0, d.aoiK) : '', aoiU: d.aoiU >= 0 ? g.mtext(r0, d.aoiU) : '',
            sebabK: d.sebabK >= 0 ? g.mtext(r0, d.sebabK) : '', sebabU: d.sebabU >= 0 ? g.mtext(r0, d.sebabU) : '',
            aoiCell: d.aoiU >= 0 ? addr(r0, d.aoiU) : '',
            kesCached: d.kes >= 0 ? g.text(r0, d.kes) : '',
          };
        }
        res.active = Object.values(res.stages).some(st => LEVELS.includes(st.grade) || st.uraian.some(x => isReal(x.text)) || isReal(st.aoiU));
        return res;
      });
      p.kesCached = {};
      for (const s of Object.keys(globalKes)) p.kesCached[s] = g.text(r0, globalKes[s]);
    }

    const activeUnitIds = new Set();
    for (const p of params) for (const pu of p.units) if (pu.active) activeUnitIds.add(pu.unit.id);
    if (!activeUnitIds.size && units.length) activeUnitIds.add(units[0].id);

    const model = { type: 'sp', name, units, activeUnits: units.filter(u => activeUnitIds.has(u.id)), subs, params, sheetFlags };
    computeSP(model);
    checkSP(model);
    return model;
  }

  // kesimpulan per parameter per tahap: modus grade antar satker (seri -> grade terendah, meniru MATCH pada E..A)
  function modeGrade(grades) {
    const cnt = { E: 0, D: 0, C: 0, B: 0, A: 0 };
    let any = false;
    for (const gr of grades) if (cnt[gr] != null) { cnt[gr]++; any = true; }
    if (!any) return '';
    let best = '', bc = 0;
    for (const k of ['E', 'D', 'C', 'B', 'A']) if (cnt[k] > bc) { bc = cnt[k]; best = k; }
    return best;
  }

  function computeSP(m) {
    for (const p of m.params) {
      p.kes = {};
      for (const s of STAGES) {
        const grades = p.units.filter(u => u.stages[s]).map(u => u.stages[s].grade);
        p.kes[s] = modeGrade(grades);
      }
    }
    for (const sub of m.subs) {
      sub.score = {};
      for (const s of STAGES) {
        const sc = sub.params.map(p => SCORE[p.kes[s]]).filter(x => x != null);
        sub.score[s] = sc.length ? { value: sc.reduce((a, b) => a + b, 0) / sc.length, complete: sc.length === sub.params.length, n: sc.length, of: sub.params.length } : null;
      }
    }
  }

  function checkSP(m) {
    // peta duplikasi uraian
    const dup = new Map();
    for (const p of m.params) for (const pu of p.units) for (const s of Object.keys(pu.stages)) {
      if (s !== 'PM' && pu.unit.sharedUraian) continue;
      for (const u of pu.stages[s].uraian) {
        if (!isReal(u.text) || norm(u.text).length < 60) continue;
        const k = low(u.text);
        if (!dup.has(k)) dup.set(k, []);
        dup.get(k).push({ p, unit: pu.unit, stage: s, lvl: u.lvl, cell: u.cell });
      }
    }

    for (const p of m.params) {
      p.flags = [];
      for (const pu of p.units) {
        if (!pu.active) continue;
        const u = pu.unit, ul = m.activeUnits.length > 1 ? u.label + ' · ' : '';
        const stages = Object.keys(pu.stages);
        for (const s of stages) {
          const st = pu.stages[s];
          const tag = ul + STAGE_LABEL[s];
          const base = { unit: u.id, stage: s };
          const hasOwnText = (!u.sharedUraian || s === 'PM') && st.uraian.some(x => isReal(x.text));
          if (!st.grade) {
            if (hasOwnText) p.flags.push(flag('warn', 'GRADE_KOSONG', `${tag}: uraian hasil pengujian sudah diisi tetapi grade kosong.`, st.gradeCell, base));
            continue;
          }
          if (!LEVELS.includes(st.grade)) { p.flags.push(flag('error', 'GRADE_TIDAK_VALID', `${tag}: grade "${st.grade}" bukan A–E.`, st.gradeCell, base)); continue; }
          const gi = LEVELS.indexOf(st.grade);
          const need = st.uraian.filter(x => LEVELS.indexOf(x.lvl) >= gi);
          const miss = need.filter(x => !isReal(x.text));
          if (miss.length) p.flags.push(flag('warn', 'BUKTI_KURANG', `${tag}: grade ${st.grade} berarti level ${need.map(x => x.lvl).join(', ')} terpenuhi, tetapi uraian level ${miss.map(x => x.lvl).join(', ')} kosong/masih template.`, miss[0].cell, base));
          const above = st.uraian.filter(x => LEVELS.indexOf(x.lvl) < gi && isReal(x.text));
          if (above.length) p.flags.push(flag('info', 'URAIAN_DI_ATAS_GRADE', `${tag}: ada uraian pada level ${above.map(x => x.lvl).join(', ')} (di atas grade ${st.grade}). Pastikan uraian itu menjelaskan mengapa level belum terpenuhi, atau pertimbangkan grade lebih tinggi.`, above[0].cell, base));
          for (const x of need) {
            if (!isReal(x.text)) continue;
            const sn = negSnippet(x.text);
            if (sn) p.flags.push(flag('warn', 'PERNYATAAN_NEGATIF', `${tag}: uraian level ${x.lvl} (dianggap terpenuhi) memuat pernyataan negatif: "${sn}"`, x.cell, base));
            else if (norm(x.text).length < 40) p.flags.push(flag('info', 'URAIAN_PENDEK', `${tag}: uraian level ${x.lvl} sangat singkat ("${norm(x.text)}").`, x.cell, base));
          }
          if (st.grade !== 'A') {
            if (!isReal(st.aoiU) && !isReal(st.aoiK)) p.flags.push(flag('warn', 'AOI_KOSONG', `${tag}: grade ${st.grade} (belum A) tetapi Kluster/Uraian AoI kosong.`, st.aoiCell, base));
            else if (!isReal(st.sebabU) && !isReal(st.sebabK)) p.flags.push(flag('info', 'PENYEBAB_KOSONG', `${tag}: AoI terisi tetapi penyebab kosong.`, st.aoiCell, base));
          } else if (isReal(st.aoiU)) {
            p.flags.push(flag('info', 'AOI_PADA_A', `${tag}: grade A tetapi ada uraian AoI.`, st.aoiCell, base));
          }
          if (!u.sharedUraian && s !== 'PM' && pu.stages.PM) {
            const a = pu.stages.PM.uraian.map(x => low(x.text)).join('|'), b = st.uraian.map(x => low(x.text)).join('|');
            if (a === b && st.uraian.some(x => isReal(x.text))) p.flags.push(flag('info', 'URAIAN_DISALIN', `${tag}: uraian hasil pengujian sama persis dengan PM.`, st.uraian[0].cell, base));
          }
        }
        // perubahan grade antar tahap
        const seq = STAGES.filter(s => pu.stages[s] && LEVELS.includes(pu.stages[s].grade));
        for (let i = 1; i < seq.length; i++) {
          const a = pu.stages[seq[i - 1]].grade, b = pu.stages[seq[i]].grade;
          if (a !== b) {
            const up = LEVELS.indexOf(b) < LEVELS.indexOf(a);
            p.flags.push(flag(up ? 'warn' : 'info', up ? 'GRADE_NAIK' : 'GRADE_TURUN', `${ul}Grade ${up ? 'naik' : 'turun'}: ${STAGE_LABEL[seq[i - 1]]} ${a} → ${STAGE_LABEL[seq[i]]} ${b}.${up ? ' Pastikan ada bukti tambahan.' : ''}`, pu.stages[seq[i]].gradeCell, { unit: u.id, stage: seq[i] }));
          }
        }
        // duplikasi (satu temuan per unit/tahap)
        for (const s of Object.keys(pu.stages)) {
          if (s !== 'PM' && u.sharedUraian) continue;
          const inside = new Set(), outside = [];
          let firstCell = '';
          for (const x of pu.stages[s].uraian) {
            if (!isReal(x.text) || norm(x.text).length < 60) continue;
            const others = (dup.get(low(x.text)) || []).filter(o => o.cell !== x.cell || o.p.sheet !== p.sheet);
            for (const o of others) {
              if (o.p === p) { inside.add(x.lvl); inside.add(o.lvl); }
              else outside.push(`${o.p.sub.kode} no.${o.p.no} lv ${o.lvl}` + (o.p.sheet !== p.sheet ? ` (${o.p.sheet})` : ''));
              if (!firstCell) firstCell = x.cell;
            }
          }
          const parts = [];
          if (inside.size) parts.push(`uraian level ${[...inside].sort().join(', ')} identik satu sama lain`);
          const uniq = [...new Set(outside)];
          if (uniq.length) parts.push(`ada uraian yang identik dengan ${uniq.slice(0, 3).join('; ')}${uniq.length > 3 ? ` (+${uniq.length - 3} lainnya)` : ''}`);
          if (parts.length) p.flags.push(flag('info', 'URAIAN_DUPLIKAT', `${ul}${STAGE_LABEL[s]}: ${parts.join('; ')} — kemungkinan salin-tempel.`, firstCell, { unit: u.id, stage: s }));
        }
      }
      // kesimpulan tersimpan vs hitung ulang
      for (const s of STAGES) {
        const cached = (p.kesCached[s] || (p.units[0] && p.units[0].stages[s] && p.units[0].stages[s].kesCached) || '').toUpperCase();
        if (cached && LEVELS.includes(cached) && p.kes[s] && cached !== p.kes[s])
          p.flags.push(flag('warn', 'KESIMPULAN_BEDA', `Kesimpulan ${STAGE_LABEL[s]} di sheet (${cached}) berbeda dengan hasil hitung ulang (${p.kes[s]}).`, '', { stage: s }));
      }
      p.flagCount = countFlags(p.flags);
    }
    m.flagCount = countFlags([].concat(m.sheetFlags, ...m.params.map(p => p.flags)));
  }

  function countFlags(fl) {
    const c = { error: 0, warn: 0, info: 0 };
    for (const f of fl) c[f.level]++;
    return c;
  }

  // ---------- KKLEAD_SPIP ----------
  function parseLead(g, name) {
    const flags = [];
    let stageRow = -1, skorRow = -1;
    const stageCols = {};
    for (let r = 0; r < 12 && r < g.nrows; r++)
      for (let c = 0; c < g.ncols; c++) {
        const t = low(g.text(r, c));
        if (t === 'pm' || t === 'pk' || t === 'evaluasi') { stageRow = r; stageCols[t === 'pm' ? 'PM' : t === 'pk' ? 'PK' : 'EV'] = c; }
      }
    for (let r = 0; r < 12 && r < g.nrows; r++) if (/^skor$/i.test(g.text(r, 3)) || /^skor$/i.test(g.text(r, stageCols.PM))) skorRow = r;
    const labelCol = 2;
    const rows = [];
    for (let r = (skorRow > 0 ? skorRow + 1 : 6); r < g.nrows; r++) {
      const label = g.text(r, labelCol);
      if (!label) continue;
      const m = /\((\d+\.\d+)\)\s*$/.exec(label);
      const entry = { row: r, label, kode: m ? m[1] : '', vals: {} };
      let hasVal = false;
      for (const s of STAGES) {
        const c = stageCols[s];
        if (c == null) continue;
        const x = g.cell(r, c);
        const formula = !!(x && x.f);
        const v = g.num(r, c);
        if (x && (formula || v != null)) hasVal = true;
        entry.vals[s] = { formula, value: formula ? null : v, cell: addr(r, c), f: x && x.f };
      }
      entry.section = !hasVal;
      rows.push(entry);
    }
    // sel skor yang diketik manual padahal tahap lain memakai rumus
    for (const e of rows) {
      if (e.section) continue;
      const fs = STAGES.filter(s => e.vals[s] && e.vals[s].formula);
      if (!fs.length) continue;
      for (const s of STAGES) {
        const v = e.vals[s];
        if (v && !v.formula && v.value != null)
          flags.push(flag('warn', 'SKOR_MANUAL', `Skor ${STAGE_LABEL[s]} "${e.label}" diketik manual (${v.value}), bukan rumus — tidak mengikuti isian KK.`, v.cell, { stage: s, row: e.row }));
      }
    }
    const model = { type: 'lead', name, rows, flags, stageCols };
    model.flagCount = countFlags(flags);
    return model;
  }

  // ---------- tabel bertahap generik (KKE, KK4, KK5, KK6-8, KKLEAD I/II/III) ----------
  function parseTable(g, name) {
    const idx = findIndexRow(g, 20);
    // baris header: baris teks tepat di atas baris penomoran
    let lastCol = 0, width = 0;
    for (let c = 0; c < g.ncols; c++) if (g.cell(idx, c) && (g.cell(idx, c).f || g.text(idx, c))) width++;
    const hdrRows = [];
    for (let r = idx - 1; r >= 1 && hdrRows.length < 3; r--) {
      let n = 0;
      for (let c = 0; c < g.ncols; c++) if (g.mtext(r, c)) n++;
      if (n === 0 || (hdrRows.length > 0 && n < 4) || (hdrRows.length >= 2 && n < width * 0.6)) break;
      hdrRows.unshift(r);
    }
    for (let c = 0; c < g.ncols; c++) if (g.cell(idx, c) && (g.cell(idx, c).f || g.text(idx, c))) lastCol = c;
    for (const r of hdrRows) for (let c = 0; c < g.ncols; c++) if (g.text(r, c) && c > lastCol) lastCol = c;

    // baris tahap (PM/PK/EVALUASI) di atas header
    const stageStarts = [];
    for (let r = 1; r < idx; r++) {
      for (let c = 0; c <= lastCol; c++) {
        const t = low(g.text(r, c));
        if (t === 'pm' || t === 'pk' || t === 'evaluasi') {
          const mg = g.merges.find(m => m.s.r === r && m.s.c === c);
          stageStarts.push({ s: t === 'pm' ? 'PM' : t === 'pk' ? 'PK' : 'EV', c0: c, c1: mg ? mg.e.c : null });
        }
      }
      if (stageStarts.length) break;
    }
    stageStarts.sort((a, b) => a.c0 - b.c0);
    // sel "PM" bisa di-merge lebih sempit dari bloknya -> blok berakhir tepat sebelum tahap berikutnya
    stageStarts.forEach((st, i) => {
      if (i + 1 < stageStarts.length) st.c1 = stageStarts[i + 1].c0 - 1;
      else st.c1 = Math.max(st.c1 || st.c0, i > 0 ? Math.min(lastCol, st.c0 + (stageStarts[i - 1].c1 - stageStarts[i - 1].c0)) : lastCol);
    });

    const cols = [];
    for (let c = 0; c <= lastCol; c++) {
      const parts = [];
      for (const r of hdrRows) { const t = norm(g.mtext(r, c)); if (t && !parts.includes(t)) parts.push(t); }
      let stage = null;
      for (const st of stageStarts) if (c >= st.c0 && c <= st.c1) stage = st.s;
      const label = parts.join(' › ');
      if (!stage) {
        const m = /\b(pm|pk|evaluasi)\b/i.exec(label);
        if (m) stage = m[1].toLowerCase() === 'pm' ? 'PM' : m[1].toLowerCase() === 'pk' ? 'PK' : 'EV';
      }
      if (/^analisis/i.test(label)) stage = null;
      const key = label.toLowerCase().replace(/\b(pm|pk|evaluasi)\b/g, '').replace(/\s+/g, ' ').trim();
      cols.push({ c, name: colName(c), label: label || colName(c), stage, key, no: g.text(idx, c) || (g.cell(idx, c) && g.cell(idx, c).f ? String(c + 1) : '') });
    }
    // kolom rumus: sel pertama berisi rumus -> ditandai (nilainya tidak tersimpan)
    const rows = [];
    const notes = [];
    let inNotes = false;
    for (let r = idx + 1; r < g.nrows; r++) {
      const first = (() => { for (let c = 0; c <= lastCol + 3 && c < g.ncols; c++) { const t = g.text(r, c); if (t) return t; } return ''; })();
      if (/^(petunjuk pengisian|analisis\s*:|simpulan\s*:|catatan\s*:|kriteria$)/i.test(first)) inNotes = true;
      if (inNotes) {
        const t = []; for (let c = 0; c < g.ncols; c++) { const x = g.text(r, c); if (x) t.push(x); }
        if (t.length) notes.push({ row: r, text: t.join('  ·  ') });
        continue;
      }
      const vals = [];
      let n = 0, formulas = 0;
      for (const col of cols) {
        const t = g.mtop(r, col.c)[0] === r ? g.text(r, col.c) : '';
        if (g.isFormula(r, col.c)) formulas++;
        if (t) n++;
        vals.push(t);
      }
      if (!n) continue;
      rows.push({ row: r, vals, section: n === 1 && vals.findIndex(v => v) <= 1 && vals.slice(2).every(v => !v) && !/^\d+$/.test(vals[0]) });
    }
    // notes di sisi kanan (mis. Petunjuk di kolom AJ) diabaikan dari tabel
    const model = { type: 'table', name, cols, rows, notes, hdrRows, idx, stages: STAGES.filter(s => cols.some(c => c.stage === s)) };
    checkTable(model);
    return model;
  }

  function checkTable(m) {
    // kelompokkan kolom tahap berdasarkan label
    const byKey = new Map();
    m.cols.forEach((c, i) => {
      if (!c.stage) return;
      if (!byKey.has(c.key)) byKey.set(c.key, {});
      const o = byKey.get(c.key);
      if (o[c.stage] == null) o[c.stage] = i;
    });
    m.compareKeys = [...byKey.entries()].filter(([, o]) => Object.keys(o).length > 1);
    const allFlags = [];
    for (const row of m.rows) {
      row.flags = [];
      if (row.section) continue;
      const hasIdentity = m.cols.some((c, i) => !c.stage && row.vals[i]);
      for (const [key, o] of m.compareKeys) {
        const present = STAGES.filter(s => o[s] != null);
        for (let i = 1; i < present.length; i++) {
          const a = row.vals[o[present[i - 1]]], b = row.vals[o[present[i]]];
          if (a && b && low(a) !== low(b) && !/uraian|analisis/.test(key) && norm(a).length < 40)
            row.flags.push(flag('info', 'BEDA_TAHAP', `${m.cols[o[present[i]]].label}: ${STAGE_LABEL[present[i - 1]]} "${a}" → ${STAGE_LABEL[present[i]]} "${b}".`, addr(row.row, m.cols[o[present[i]]].c)));
        }
      }
      // Y/T -> T tanpa AoI
      m.cols.forEach((c, i) => {
        if (!/y\/t$/i.test(c.label)) return;
        const v = norm(row.vals[i]).toUpperCase();
        if (!v) {
          if (hasIdentity && c.stage === 'PM') row.flags.push(flag('warn', 'YT_KOSONG', `${c.label} (PM) belum diisi.`, addr(row.row, c.c)));
          return;
        }
        if (v !== 'Y' && v !== 'T') { row.flags.push(flag('error', 'YT_TIDAK_VALID', `${c.label} (${STAGE_LABEL[c.stage] || ''}) berisi "${v}", seharusnya Y/T.`, addr(row.row, c.c))); return; }
        if (v === 'T') {
          let aoi = false;
          for (let j = i + 1; j < m.cols.length && j <= i + 4; j++) {
            if (/y\/t$/i.test(m.cols[j].label)) break;
            if (/aoi|penyebab/i.test(m.cols[j].label) && isReal(row.vals[j])) aoi = true;
          }
          if (!aoi) row.flags.push(flag('warn', 'T_TANPA_AOI', `${c.label} (${STAGE_LABEL[c.stage] || ''}) = T tetapi AoI/penyebab kosong.`, addr(row.row, c.c)));
        }
      });
      // target ada, realisasi kosong (KK5)
      m.cols.forEach((c, i) => {
        if (!/target$/i.test(c.label)) return;
        const j = m.cols.findIndex((d, k) => k > i && d.stage === c.stage && /realisasi$/i.test(d.label) && !/persentase/i.test(d.label));
        if (j >= 0 && row.vals[i] && !row.vals[j]) row.flags.push(flag('warn', 'REALISASI_KOSONG', `${STAGE_LABEL[c.stage] || ''}: target "${row.vals[i]}" ada tetapi realisasi kosong.`, addr(row.row, m.cols[j].c)));
      });
      allFlags.push(...row.flags);
    }
    m.flagCount = countFlags(allFlags);
  }

  // ---------- info (FAQ, petunjuk, daftar) ----------
  function parseInfo(g, name) {
    const lines = [];
    for (let r = 0; r < g.nrows; r++) {
      const t = [];
      for (let c = 0; c < g.ncols; c++) { const x = g.text(r, c); if (x) t.push(x); }
      if (t.length) lines.push({ row: r, cells: t });
    }
    return { type: 'info', name, lines, flagCount: { error: 0, warn: 0, info: 0 } };
  }

  function parseSheet(ws, name) {
    const g = new Grid(ws);
    const kind = classify(g);
    try {
      if (kind === 'sp') return parseSP(g, name);
      if (kind === 'lead') return parseLead(g, name);
      if (kind === 'table') return parseTable(g, name);
    } catch (e) {
      const m = parseInfo(g, name);
      m.error = String(e && e.stack || e);
      return m;
    }
    return parseInfo(g, name);
  }

  // ---------- rekap skor Struktur & Proses lintas KK 3.1–3.4 ----------
  function combineSP(models) {
    const sps = models.filter(m => m.type === 'sp');
    const flags = [];
    // kode mayoritas per nama subunsur (nama bisa sedikit beda/typo -> pakai 20 huruf awal)
    const nk = s => low(s.nama).replace(/[^a-z]/g, '').slice(0, 20);
    const votes = new Map();
    for (const m of sps) for (const s of m.subs) {
      const k = nk(s);
      if (!votes.has(k)) votes.set(k, {});
      votes.get(k)[s.kode] = (votes.get(k)[s.kode] || 0) + 1;
    }
    const major = k => Object.entries(votes.get(k)).sort((a, b) => b[1] - a[1])[0][0];
    const map = new Map();
    for (const m of sps) for (const s of m.subs) {
      const kode = major(nk(s));
      if (kode !== s.kode) flags.push(flag('warn', 'KODE_SUBUNSUR_BEDA', `${m.name}: subunsur "${s.nama}" berkode ${s.kode}, sedangkan di KK lain ${kode}.`, `${m.name}!A${s.row + 1}`, { sheet: m.name }));
      if (!map.has(kode)) map.set(kode, { kode, nama: s.nama, per: {} });
      map.get(kode).per[m.name] = s.score;
    }
    const rows = [...map.values()].sort((a, b) => a.kode.localeCompare(b.kode, undefined, { numeric: true }));
    for (const r of rows) {
      r.avg = {};
      for (const st of STAGES) {
        const v = Object.values(r.per).map(x => x && x[st] && x[st].value).filter(x => x != null);
        r.avg[st] = v.length ? v.reduce((a, b) => a + b, 0) / v.length : null;
      }
    }
    return { sheets: sps.map(m => m.name), rows, flags };
  }

  const api = { parseSheet, combineSP, isPlaceholder, isReal, STAGES, STAGE_LABEL, LEVELS, SCORE, colName, addr };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.SpipParser = api;
})(this);
