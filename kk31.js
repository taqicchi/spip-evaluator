/* Ruang kerja evaluator KK 3.1 (Struktur & Proses, tingkat satker).
 * PM diisi satker (uraian hasil pengujian + Grade PM). Evaluator (PK/Evaluasi) memverifikasi
 * tiap level dengan bukti, lalu mengisi Grade PK/Evaluasi, Kluster & Uraian AoI, Kluster & Uraian Penyebab.
 */
(function () {
  'use strict';
  const P = window.SpipParser, Z = window.MiniZip, XP = window.XlsxPatch;
  const { LEVELS, SCORE, STAGE_LABEL } = P;
  const $ = (s, el) => (el || document).querySelector(s);
  const $$ = (s, el) => Array.from((el || document).querySelectorAll(s));

  const DEFAULT_LISTS = {
    SPIP: ['Kebijakan belum memadai', 'Implementasi belum sesuai dengan kebijakan', 'Pengendalian yang ada belum diarahkan untuk memitigasi risiko', 'Pengendalian yang ada belum efektif'],
    MRI: ['Perencanaan belum mempertimbangkan risiko', 'Pimpinan belum/kurang mendukung penerapan MR', 'Kebijakan MR belum ada/memadai', 'Struktur MR belum ada/memadai', 'Struktur MR belum berjalan', 'Penerapan MR belum didukung sumber daya yang memadai', 'Risiko kemitraan belum dikelola', 'Risiko strategis K/L/D belum dikelola', 'Risiko strategis unit kerja belum dilkelola', 'Risiko operasional unit kerja belum dikelola', 'Proses MR belum diterapkan sesuai kebijakan', 'RTP belum dilaksanakan', 'RTP tidak dimonitor', 'RTP tidak efektif', 'Keterjadian risiko tidak dipantau', 'Proses MR tidak dilakukan reviu/evaluasi oleh APIP'],
    IEPK: ['kebijakan pengendalian korupsi belum ada/memadai', 'Sistem anti korupsi yang dibangun belum memadai', 'Pengendalian korupsi belum didukung sumber daya yang memadai', 'Pimpinan belum/kurang mendukung penerapan pengendalian korupsi', 'Program pendidikan antikorupsi belum dilakukan/memadai', 'Risiko kecurangan belum dikelola', 'Sistem pengaduan belum dibangun/tidak berjalan/kurang memadai', 'Budaya antikorupsi belum terbangun/memadai', 'Kejadian korupsi belum dilakukan tindakan investigasi secara memadai', 'Pemberian sanksi belum sesuai dengan tindakan korupsi yang dilakukan', 'Perbaikan pengendalian atas korupsi yang terjadi belum dilakukan/memadai'],
    PENYEBAB: ['Man', 'Method', 'Money', 'Material', 'Machine'],
  };
  const PM_HINT_CODES = new Set(['BUKTI_KURANG', 'PERNYATAAN_NEGATIF', 'URAIAN_DI_ATAS_GRADE', 'GRADE_KOSONG', 'GRADE_TIDAK_VALID', 'URAIAN_PENDEK', 'URAIAN_DUPLIKAT']);
  const VERIF = [['ya', '✓ Terbukti'], ['sebagian', '~ Sebagian'], ['tidak', '✗ Tidak']];

  const S = {
    file: null, fileName: '', sheetName: '', model: null, lists: DEFAULT_LISTS,
    evidence: null, stage: 'PK', unitId: '', idx: 0, view: 'param', filter: 'all',
    ev: {}, showPenjelasan: true,
  };

  // ---------- util ----------
  function esc(s) { return String(s == null ? '' : s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c])); }
  function fmt(n, d) { return n == null || isNaN(n) ? '–' : Number(n).toFixed(d == null ? 2 : d).replace('.', ','); }
  function toast(msg, ms) { const t = $('#toast'); t.textContent = msg; t.hidden = false; clearTimeout(toast._t); toast._t = setTimeout(() => { t.hidden = true; }, ms || 2800); }
  function busy(on, text) { $('#busy').hidden = !on; if (text) $('#busyText').textContent = text; }
  const tick = () => new Promise(r => setTimeout(r, 30));
  const unit = () => S.model.units.find(u => u.id === S.unitId);
  const params = () => S.model.params;
  function download(blob, name) {
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob); a.download = name; document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(a.href), 4000);
  }
  function dateStamp() { const d = new Date(); return `${d.getFullYear()}${String(d.getMonth() + 1).padStart(2, '0')}${String(d.getDate()).padStart(2, '0')}`; }

  // ---------- penyimpanan isian evaluator ----------
  const storeKey = () => 'spip-kk31:v1:' + S.fileName;
  function loadStore() {
    try {
      const j = JSON.parse(localStorage.getItem(storeKey()) || '{}');
      S.ev = j.ev || {};
      if (j.stage) S.stage = j.stage;
      if (j.unitId) S.unitId = j.unitId;
    } catch (e) { S.ev = {}; }
  }
  function saveStore() {
    try { localStorage.setItem(storeKey(), JSON.stringify({ ev: S.ev, stage: S.stage, unitId: S.unitId })); }
    catch (e) { toast('Gagal menyimpan di peramban — gunakan Sesi → Simpan sesi.'); }
  }
  const recKey = (p, stage, unitId) => `${stage || S.stage}|${unitId || S.unitId}|${p.row}`;

  // nilai yang sudah ada di berkas untuk tahap ini (mis. melanjutkan KK yang sudah sebagian diisi)
  function fileValues(p, stage, unitId) {
    const pu = p.units.find(x => x.unit.id === unitId);
    const st = pu && pu.stages[stage];
    if (!st) return {};
    const r = v => (P.isReal(v) ? v : '');
    return { grade: LEVELS.includes(st.grade) ? st.grade : '', gradeRaw: st.grade, aoiK: r(st.aoiK), aoiU: r(st.aoiU), sebabK: r(st.sebabK), sebabU: r(st.sebabU) };
  }
  function getRec(p, stage, unitId) {
    stage = stage || S.stage; unitId = unitId || S.unitId;
    const k = recKey(p, stage, unitId);
    if (S.ev[k]) return S.ev[k];
    const f = fileValues(p, stage, unitId);
    return { lv: {}, grade: f.grade || '', tdn: !!(f.gradeRaw && /tidak dapat/i.test(f.gradeRaw)), aoiK: f.aoiK || '', aoiU: f.aoiU || '', sebabK: f.sebabK || '', sebabU: f.sebabU || '', note: '', done: false, fromFile: !!(f.grade || f.aoiU || f.aoiK) };
  }
  function setRec(p, patch) {
    const r = Object.assign({}, getRec(p), patch, { ts: new Date().toISOString() });
    delete r.fromFile;
    S.ev[recKey(p)] = r;
    saveStore();
    return r;
  }
  function setLevel(p, lvl, patch) {
    const r = getRec(p);
    const lv = Object.assign({}, r.lv);
    lv[lvl] = Object.assign({}, lv[lvl] || {}, patch);
    return setRec(p, { lv });
  }
  function hasInput(r) { return !!(r.grade || r.tdn || r.aoiU || r.aoiK || r.sebabU || r.note || Object.keys(r.lv || {}).length); }

  // saran grade: level tertinggi yang ia dan semua level di bawahnya "terbukti"
  function suggest(p, r) {
    const order = p.levels.map(l => l.lvl).sort((a, b) => LEVELS.indexOf(b) - LEVELS.indexOf(a)); // E..A
    let best = '';
    for (const l of order) { if ((r.lv[l] || {}).v === 'ya') best = l; else break; }
    return best;
  }

  // ---------- membuka berkas KK ----------
  async function openFile(file) {
    if (!file) return;
    busy(true, 'Membaca berkas…');
    await tick();
    try {
      const buf = await file.arrayBuffer();
      const names = XLSX.read(buf, { bookSheets: true }).SheetNames;
      const cands = names.filter(n => /^kk\s*3\.?\s*1\b/i.test(n.trim()));
      const want = cands.length ? [cands[0]] : names;
      if (names.includes('REF')) want.push('REF');
      busy(true, 'Menganalisis KK 3.1…');
      await tick();
      const wb = XLSX.read(buf, { dense: true, cellFormula: true, sheetStubs: true, sheets: want });
      let model = null;
      for (const n of want) {
        if (n === 'REF') continue;
        const m = P.parseSheet(wb.Sheets[n], n);
        if (m.type === 'sp' && m.units.some(u => u.sharedUraian)) { model = m; break; }
      }
      if (!model) throw new Error('Sheet KK 3.1 (blok per satker dengan Grade PM/PK/Evaluasi) tidak ditemukan di berkas ini.');
      S.file = file; S.fileName = file.name; S.sheetName = model.name; S.model = model;
      S.lists = readLists(wb.Sheets.REF);
      S.ev = {}; S.stage = 'PK'; S.unitId = '';
      loadStore();
      if (!S.model.units.some(u => u.id === S.unitId)) S.unitId = (model.activeUnits[0] || model.units[0]).id;
      S.idx = 0; S.view = 'param';
    } catch (e) {
      busy(false); console.error(e); alert(e.message); return;
    }
    busy(false);
    $('#fileLabel').textContent = `${S.fileName} · sheet ${S.sheetName}`;
    $('#landing').hidden = true; $('#workspace').hidden = false; $('#actions').hidden = false;
    renderTop(); render();
  }

  function readLists(ws) {
    if (!ws) return DEFAULT_LISTS;
    const d = ws['!data'] || [];
    const txt = (r, c) => { const x = d[r] && d[r][c]; return x && x.t !== 'z' && x.v != null ? String(x.w != null ? x.w : x.v).trim() : ''; };
    const out = Object.assign({}, DEFAULT_LISTS);
    const heads = { SPIP: /^sub unsur spip$/i, MRI: /^area mri$/i, IEPK: /^area iepk$/i, PENYEBAB: /^kl[au]ster penyebab$/i };
    for (let r = 0; r < Math.min(d.length, 80); r++) for (let c = 0; c < 30; c++) {
      const t = txt(r, c);
      for (const k in heads) if (heads[k].test(t)) {
        const vals = [];
        for (let rr = r + 1; rr < d.length && txt(rr, c); rr++) vals.push(txt(rr, c));
        if (vals.length) out[k] = vals;
      }
    }
    return out;
  }
  function aoiList(p) {
    const k = p.kodeParam.map(x => x.toUpperCase());
    if (k.includes('MRI') && !k.includes('SPIP')) return ['MRI', S.lists.MRI];
    if (k.includes('IEPK') && !k.includes('SPIP')) return ['IEPK', S.lists.IEPK];
    return ['SPIP', S.lists.SPIP];
  }

  // ---------- bukti ----------
  const MIME = { pdf: 'application/pdf', jpg: 'image/jpeg', jpeg: 'image/jpeg', png: 'image/png', gif: 'image/gif', webp: 'image/webp', bmp: 'image/bmp', txt: 'text/plain', csv: 'text/plain', htm: 'text/html', html: 'text/html', mp4: 'video/mp4' };
  const ext = n => (/\.([a-z0-9]+)$/i.exec(n) || [])[1] ? /\.([a-z0-9]+)$/i.exec(n)[1].toLowerCase() : '';

  function indexEvidence(items, sourceName, kind, emptyDirs) {
    // items: [{path, size, open()}]
    const byKode = new Map();
    const unmatched = [];
    const place = (path) => {
      const seg = path.split('/').filter(Boolean);
      const i = seg.findIndex(s => /^\d+\.\d+(\s|-|_|$)/.test(s.trim()));
      if (i < 0) return null;
      const kode = /^(\d+\.\d+)/.exec(seg[i].trim())[1];
      const lv = seg[i + 1] && /^[A-E]$/i.test(seg[i + 1].trim()) ? seg[i + 1].trim().toUpperCase() : '?';
      return { kode, folder: seg[i], lv, rest: seg.slice(lv === '?' ? i + 1 : i + 2) };
    };
    const slot = (kode, folder) => {
      if (!byKode.has(kode)) byKode.set(kode, { folder, levels: { A: [], B: [], C: [], D: [], E: [], '?': [] }, dirs: new Set() });
      return byKode.get(kode);
    };
    for (const it of items) {
      const pl = place(it.path);
      if (!pl || !pl.rest.length) { if (!pl) unmatched.push(it); continue; }
      const s = slot(pl.kode, pl.folder);
      s.levels[pl.lv].push(Object.assign({ name: pl.rest[pl.rest.length - 1], sub: pl.rest.slice(0, -1).join('/') }, it));
      if (pl.lv !== '?') s.dirs.add(pl.lv);
    }
    for (const dpath of emptyDirs || []) { const pl = place(dpath); if (pl) { const s = slot(pl.kode, pl.folder); if (pl.lv !== '?') s.dirs.add(pl.lv); } }
    S.evidence = { name: sourceName, kind, byKode, count: items.length, unmatched };
    renderTop();
    render();
    toast(`${items.length} berkas bukti dimuat untuk ${byKode.size} subunsur${unmatched.length ? ` · ${unmatched.length} berkas tidak dikenali foldernya` : ''}.`, 4000);
  }

  function loadFolder(fileList) {
    const files = Array.from(fileList);
    if (!files.length) return;
    const root = (files[0].webkitRelativePath || '').split('/')[0];
    indexEvidence(files.map(f => ({ path: f.webkitRelativePath || f.name, size: f.size, open: async () => f })), root || 'folder', 'folder', []);
  }
  async function loadZip(file) {
    if (!file) return;
    busy(true, 'Membaca daftar isi ZIP…');
    await tick();
    try {
      const z = await Z.open(file);
      const files = z.entries.filter(e => !e.dir);
      const dirs = z.entries.filter(e => e.dir).map(e => e.name);
      indexEvidence(files.map(e => ({
        path: e.name, size: e.usize,
        open: async () => new Blob([await Z.extract(z, e)], { type: MIME[ext(e.name)] || 'application/octet-stream' }),
      })), file.name, 'zip', dirs);
    } catch (e) { alert('ZIP tidak dapat dibaca: ' + e.message); }
    busy(false);
  }
  function evidenceFor(p) {
    if (!S.evidence) return null;
    return S.evidence.byKode.get(p.sub.kode) || null;
  }

  let viewerUrl = null;
  async function openViewer(item, lvl, kode) {
    const v = $('#viewer');
    v.hidden = false;
    $('#viewerTitle').innerHTML = `<b>${esc(item.name)}</b><div class="muted small">${esc(kode)} · level ${esc(lvl)}${item.sub ? ' · ' + esc(item.sub) : ''} · ${fmt(item.size / 1024, 0)} KB</div>`;
    $('#viewerBody').innerHTML = '<div class="muted" style="padding:20px">Memuat…</div>';
    document.body.classList.add('viewing');
    try {
      let blob = await item.open();
      const e = ext(item.name);
      if (MIME[e] && blob.type !== MIME[e]) blob = new Blob([blob], { type: MIME[e] });
      if (viewerUrl) URL.revokeObjectURL(viewerUrl);
      viewerUrl = URL.createObjectURL(blob);
      const a = $('#viewerDownload'); a.href = viewerUrl; a.download = item.name;
      let body;
      if (/^image\//.test(blob.type)) body = `<img src="${viewerUrl}" alt="">`;
      else if (blob.type === 'application/pdf' || /^text\//.test(blob.type) || blob.type === 'video/mp4') body = `<iframe src="${viewerUrl}" title="Pratinjau"></iframe>`;
      else body = `<div class="empty-note">Pratinjau tidak tersedia untuk berkas .${esc(e)}. Gunakan tombol <b>Unduh</b> lalu buka dengan aplikasinya.</div>`;
      $('#viewerBody').innerHTML = body;
    } catch (err) {
      $('#viewerBody').innerHTML = `<div class="empty-note">Gagal membuka berkas: ${esc(err.message)}</div>`;
    }
  }
  function closeViewer() { $('#viewer').hidden = true; document.body.classList.remove('viewing'); $('#viewerBody').innerHTML = ''; }

  // ---------- catatan atas isian satker (PM) ----------
  function pmHints(p) {
    const out = p.flags.filter(f => f.stage === 'PM' && (!f.unit || f.unit === S.unitId) && PM_HINT_CODES.has(f.code))
      .map(f => ({ level: f.level, msg: f.msg.replace(/^(.*? · )?PM: /, ''), cell: f.cell }));
    const pu = p.units.find(x => x.unit.id === S.unitId);
    const pm = pu && pu.stages.PM;
    const ev = evidenceFor(p);
    if (S.evidence && pm) {
      if (!ev) out.push({ level: 'warn', msg: `Folder bukti subunsur ${p.sub.kode} tidak ditemukan di ${S.evidence.name}.` });
      else if (LEVELS.includes(pm.grade)) {
        const gi = LEVELS.indexOf(pm.grade);
        const claimed = p.levels.map(l => l.lvl).filter(l => LEVELS.indexOf(l) >= gi);
        const empty = claimed.filter(l => !ev.levels[l].length);
        if (empty.length) out.push({ level: 'warn', msg: `Grade PM ${pm.grade} mengklaim level ${claimed.join(', ')}, tetapi folder bukti level ${empty.join(', ')} kosong${ev.dirs.size ? '' : ' / tidak ada'}.` });
      }
      if (ev && p.sub.params.length > 1) out.push({ level: 'info', msg: `Folder bukti disusun per subunsur; ${p.sub.params.length} parameter di subunsur ${p.sub.kode} memakai folder yang sama.` });
    }
    return out;
  }

  // ---------- kerangka ----------
  function renderTop() {
    $$('#stageSeg button').forEach(b => b.classList.toggle('on', b.dataset.s === S.stage));
    $('#unitSel').innerHTML = S.model.units.map(u => `<option value="${u.id}">${esc(u.label)}${S.model.activeUnits.includes(u) ? '' : ' (belum diisi satker)'}</option>`).join('');
    $('#unitSel').value = S.unitId;
    $('#buktiLabel').textContent = S.evidence ? `Bukti: ${S.evidence.count} berkas ▾` : 'Bukti: belum dimuat ▾';
    $('#btnRekap').classList.toggle('on', S.view === 'rekap');
  }

  function progress() {
    let done = 0, filled = 0;
    for (const p of params()) { const r = getRec(p); if (r.done) done++; if (r.grade || r.tdn) filled++; }
    return { done, filled, total: params().length };
  }

  function navList() {
    const pr = progress();
    let h = `<div class="navhead"><div class="small muted">${STAGE_LABEL[S.stage]} · ${esc(unit().label)}</div>
      <div><b>${pr.done}</b>/${pr.total} selesai · ${pr.filled} bergrade</div><div class="bar"><i style="width:${(100 * pr.done / pr.total).toFixed(1)}%"></i></div>
      <select id="navFilter" class="navfilter">
        <option value="all">Semua parameter</option><option value="todo">Belum selesai</option><option value="hint">Ada catatan isian satker</option>
        <option value="diff">Grade beda dengan PM</option><option value="klar">Perlu klarifikasi</option></select></div>`;
    let lastSub = null;
    params().forEach((p, i) => {
      const r = getRec(p);
      const pm = (p.units.find(x => x.unit.id === S.unitId) || { stages: {} }).stages.PM;
      const pmg = pm && LEVELS.includes(pm.grade) ? pm.grade : '';
      const g = r.tdn ? 'TDN' : r.grade;
      if (S.filter === 'todo' && r.done) return;
      if (S.filter === 'hint' && !pmHints(p).some(x => x.level !== 'info')) return;
      if (S.filter === 'diff' && !(g && pmg && g !== pmg)) return;
      if (S.filter === 'klar' && !r.klarifikasi) return;
      if (p.sub !== lastSub) { h += `<h4>${esc(p.sub.kode)} ${esc(p.sub.nama)}</h4>`; lastSub = p.sub; }
      const cls = r.done ? 'done' : hasInput(r) ? 'wip' : '';
      h += `<a data-i="${i}" class="${i === S.idx && S.view === 'param' ? 'active' : ''}">
        <span class="st ${cls}" title="${r.done ? 'Selesai' : hasInput(r) ? 'Sedang dikerjakan' : 'Belum'}"></span>
        <span class="nm">${esc(p.no)}. ${esc(p.uraian)}</span>
        <span class="gg"><span class="mini PM">${pmg || '–'}</span>→<span class="mini ${S.stage}">${g === 'TDN' ? '?' : g || '–'}</span></span>${r.klarifikasi ? '<span class="b warn">!</span>' : ''}</a>`;
    });
    return h;
  }

  function render() {
    renderTop();
    $('#nav').innerHTML = navList();
    $('#navFilter').value = S.filter;
    $('#navFilter').onchange = e => { S.filter = e.target.value; render(); };
    $$('#nav a[data-i]').forEach(a => a.onclick = () => { S.idx = +a.dataset.i; S.view = 'param'; render(); $('#content').scrollTop = 0; });
    if (S.view === 'rekap') renderRekap(); else renderParam();
    const act = $('#nav a.active');
    if (act) act.scrollIntoView({ block: 'nearest' });
  }

  // ---------- tampilan parameter ----------
  function renderParam() {
    const p = params()[S.idx];
    const u = unit();
    const pu = p.units.find(x => x.unit.id === u.id);
    const pm = pu && pu.stages.PM;
    const r = getRec(p);
    const sug = suggest(p, r);
    const ev = evidenceFor(p);
    const hints = pmHints(p);
    const [listName, aoiOpts] = aoiList(p);
    const pmg = pm && pm.grade ? pm.grade : '';
    const gi = LEVELS.includes(pmg) ? LEVELS.indexOf(pmg) : 99;
    const other = S.stage === 'EV' ? fileValues(p, 'PK', u.id) : null;
    const stageSt = pu && pu.stages[S.stage];
    const cellRef = c => (c ? `${S.sheetName}!${c}` : '');

    let h = `<div class="phead">
      <div class="crumbs"><span class="muted">${esc(p.sub.kode)} ${esc(p.sub.nama)}</span> · Parameter ${esc(p.no)} ${p.kodeParam.map(k => `<span class="tag">${esc(k)}</span>`).join(' ')}
        <span class="cell muted">${esc(S.sheetName)}!${P.addr(p.row, 3)}</span></div>
      <div class="pnav"><button class="btn sm" id="prev" ${S.idx ? '' : 'disabled'}>‹</button><span>${S.idx + 1}/${params().length}</span><button class="btn sm" id="next" ${S.idx < params().length - 1 ? '' : 'disabled'}>›</button></div>
    </div>
    <h2 class="ptitle">${esc(p.uraian)}</h2>
    <div class="pmline">
      <div class="gbox"><div class="muted small">Grade PM (satker)</div>${LEVELS.includes(pmg) ? `<span class="grade big PM">${pmg}</span>` : `<span class="grade big empty">${esc(pmg) || '–'}</span>`}</div>
      ${other && (other.grade || other.gradeRaw) ? `<div class="gbox"><div class="muted small">Grade PK (di berkas)</div><span class="grade big PK">${esc(other.grade || other.gradeRaw)}</span></div>` : ''}
      <div class="hints">${hints.length ? `<div class="muted small">Catatan atas isian satker</div><ul class="flags">${hints.map(f => `<li class="${f.level}">${esc(f.msg)}${f.cell ? ` <span class="cell">${esc(f.cell)}</span>` : ''}</li>`).join('')}</ul>` : `<div class="muted small">Tidak ada catatan otomatis atas isian satker.</div>`}</div>
    </div>`;
    if (!pu || !pu.active) h += `<div class="empty-note" style="margin:10px 0">${esc(u.label)} belum mengisi PM untuk parameter ini.</div>`;

    // tangga level
    h += `<table class="ladder verif"><thead><tr><th>Level</th><th>Kriteria <label class="small" style="text-transform:none;font-weight:400;margin-left:8px"><input type="checkbox" id="togPenj" ${S.showPenjelasan ? 'checked' : ''}> penjelasan</label></th><th>Uraian hasil pengujian (satker)</th><th>Bukti${ev ? '' : S.evidence ? '' : ' <span class="muted" style="text-transform:none">(belum dimuat)</span>'}</th><th>Verifikasi ${STAGE_LABEL[S.stage]}</th></tr></thead><tbody>`;
    p.levels.forEach((l, i) => {
      const claim = LEVELS.indexOf(l.lvl) >= gi;
      const ur = pm ? pm.uraian[i] : null;
      const txt = ur ? ur.text : '';
      const files = ev ? ev.levels[l.lvl] : [];
      const lv = r.lv[l.lvl] || {};
      h += `<tr class="${claim ? 'claim' : ''} v-${lv.v || ''}" data-l="${l.lvl}">
        <td class="lv"><span class="lvb">${l.lvl}</span>${pmg === l.lvl ? '<div class="marks"><span class="stg PM">PM</span></div>' : ''}${sug === l.lvl ? '<div class="marks"><span class="stg sug">saran</span></div>' : ''}</td>
        <td class="kr">${esc(l.kriteria)}${S.showPenjelasan && (l.penjelasan || l.cara) ? `<div class="pj">${esc(l.penjelasan)}${l.cara ? `\nCara uji: ${esc(l.cara)}` : ''}</div>` : ''}</td>
        <td class="ur">${!txt ? '<span class="ph">kosong</span>' : P.isPlaceholder(txt) ? `<span class="ph">template: ${esc(txt.replace(/\s+/g, ' '))}</span>` : esc(txt)}${ur ? `<div class="cell muted">${esc(ur.cell)}</div>` : ''}</td>
        <td class="bk">${!S.evidence ? '' : !ev ? '<span class="ph">folder subunsur tidak ada</span>' : files.length ? files.map((f, k) => `<button class="filechip" data-l="${l.lvl}" data-k="${k}" title="${esc(f.path)}"><span class="fx">${esc(ext(f.name) || '?')}</span>${esc(f.name)}</button>`).join('') : `<span class="ph">${ev.dirs.has(l.lvl) ? 'folder kosong' : 'tidak ada folder'}</span>`}</td>
        <td class="vf"><span class="seg vseg">${VERIF.map(([v, lab]) => `<button type="button" data-v="${v}" class="${lv.v === v ? 'on ' + v : ''}">${lab}</button>`).join('')}</span>
          <textarea class="lvnote" rows="2" placeholder="Catatan pengujian level ${l.lvl}…">${esc(lv.note || '')}</textarea></td>
      </tr>`;
    });
    if (ev && ev.levels['?'].length) h += `<tr><td class="lv">–</td><td class="kr muted" colspan="2">Berkas di folder subunsur tanpa subfolder level</td><td class="bk" colspan="2">${ev.levels['?'].map((f, k) => `<button class="filechip" data-l="?" data-k="${k}"><span class="fx">${esc(ext(f.name))}</span>${esc(f.name)}</button>`).join('')}</td></tr>`;
    h += `</tbody></table>`;

    // keputusan
    const gradeCell = stageSt ? cellRef(P.addr(p.row, u.stages[S.stage].grade)) : '';
    h += `<section class="decide">
      <div class="drow">
        <div><div class="dl">Grade ${STAGE_LABEL[S.stage]} <span class="cell muted">${esc(gradeCell)}</span></div>
          <span class="seg gseg">${LEVELS.map(g => `<button type="button" data-g="${g}" class="${!r.tdn && r.grade === g ? 'on' : ''}">${g}</button>`).join('')}<button type="button" data-g="TDN" class="${r.tdn ? 'on' : ''}" title="Tidak dapat dinilai: grade dikosongkan, alasan ditulis di Uraian AoI">Tidak dapat dinilai</button></span></div>
        <div class="sugbox">${sug ? `Saran dari verifikasi: <b class="grade ${S.stage}">${sug}</b> ${!r.tdn && r.grade === sug ? '<span class="muted small">(sudah dipakai)</span>' : `<button class="btn sm" id="useSug">Pakai</button>`}` : '<span class="muted small">Tandai level mulai dari E ke atas untuk mendapat saran grade.</span>'}
          ${pmg && (r.grade || sug) ? `<div class="small muted">PM ${pmg} → ${STAGE_LABEL[S.stage]} ${r.tdn ? 'tidak dapat dinilai' : r.grade || sug}${r.grade && LEVELS.indexOf(r.grade) < LEVELS.indexOf(pmg) ? ' · <b style="color:var(--warn)">naik — pastikan bukti tambahan dicatat</b>' : ''}</div>` : ''}
          ${r.fromFile ? '<div class="small muted">Nilai awal diambil dari berkas.</div>' : ''}</div>
      </div>
      <div class="dgrid">
        <label>Kluster AoI <span class="muted small">(daftar ${listName})</span>
          <select id="aoiK"><option value="">–</option>${aoiOpts.map(o => `<option ${o === r.aoiK ? 'selected' : ''}>${esc(o)}</option>`).join('')}${r.aoiK && !aoiOpts.includes(r.aoiK) ? `<option selected>${esc(r.aoiK)}</option>` : ''}</select></label>
        <label>Kluster Penyebab
          <select id="sebabK"><option value="">–</option>${S.lists.PENYEBAB.map(o => `<option ${o === r.sebabK ? 'selected' : ''}>${esc(o)}</option>`).join('')}${r.sebabK && !S.lists.PENYEBAB.includes(r.sebabK) ? `<option selected>${esc(r.sebabK)}</option>` : ''}</select></label>
        <label>Uraian AoI <button class="linkbtn small" id="composeAoi" type="button">susun dari verifikasi</button>
          <textarea id="aoiU" rows="4" placeholder="Area yang perlu diperbaiki…">${esc(r.aoiU)}</textarea></label>
        <label>Uraian Penyebab
          <textarea id="sebabU" rows="4" placeholder="Penyebab belum tercapainya level berikutnya…">${esc(r.sebabU)}</textarea></label>
        <label class="full">Catatan evaluator (tidak ditulis ke KK; ikut di lembar kerja evaluator)
          <textarea id="evNote" rows="2">${esc(r.note)}</textarea></label>
      </div>
      ${S.stage === 'EV' ? '<div class="small muted">Catatan: KK 3.1 hanya punya satu set kolom AoI/penyebab per satker; isian Evaluasi akan menimpa AoI/penyebab dari PK.</div>' : ''}
      <div class="dactions">
        <label class="small"><input type="checkbox" id="klar" ${r.klarifikasi ? 'checked' : ''}> Perlu klarifikasi ke satker</label>
        <span class="spacer"></span>
        <button class="btn" id="prev2" ${S.idx ? '' : 'disabled'}>‹ Sebelumnya</button>
        <button class="btn primary" id="doneNext">${r.done ? 'Selesai ✓ · Berikutnya ›' : 'Tandai selesai & berikutnya ›'}</button>
      </div>
    </section>`;
    const el = $('#content');
    el.innerHTML = h;
    bindParam(p);
  }

  function bindParam(p) {
    const el = $('#content');
    const go = d => { const n = S.idx + d; if (n < 0 || n >= params().length) return; S.idx = n; render(); el.scrollTop = 0; };
    $('#prev').onclick = () => go(-1); $('#next').onclick = () => go(1); $('#prev2').onclick = () => go(-1);
    $('#doneNext').onclick = () => {
      const r = getRec(p);
      if (!r.grade && !r.tdn) { if (!confirm('Grade belum dipilih. Tetap tandai selesai?')) return; }
      setRec(p, { done: true });
      if (S.idx < params().length - 1) go(1); else { render(); toast('Parameter terakhir selesai. Buka Rekap untuk memeriksa lalu Unduh KK terisi.'); }
    };
    $('#togPenj').onchange = e => { S.showPenjelasan = e.target.checked; renderParam(); };
    $$('.vseg button', el).forEach(b => b.onclick = () => {
      const l = b.closest('tr').dataset.l;
      const cur = (getRec(p).lv[l] || {}).v;
      setLevel(p, l, { v: cur === b.dataset.v ? '' : b.dataset.v });
      keepScroll(renderParam); refreshNav();
    });
    $$('.lvnote', el).forEach(t => { let tm; t.oninput = () => { clearTimeout(tm); tm = setTimeout(() => setLevel(p, t.closest('tr').dataset.l, { note: t.value }), 300); }; });
    $$('.gseg button', el).forEach(b => b.onclick = () => {
      const g = b.dataset.g;
      if (g === 'TDN') setRec(p, { tdn: !getRec(p).tdn, grade: '' }); else setRec(p, { grade: getRec(p).grade === g && !getRec(p).tdn ? '' : g, tdn: false });
      keepScroll(renderParam); refreshNav();
    });
    if ($('#useSug')) $('#useSug').onclick = () => { setRec(p, { grade: suggest(p, getRec(p)), tdn: false }); keepScroll(renderParam); refreshNav(); };
    const bindText = (id, field) => { let tm; $('#' + id).oninput = e => { clearTimeout(tm); tm = setTimeout(() => { setRec(p, { [field]: e.target.value }); refreshNav(); }, 300); }; };
    bindText('aoiU', 'aoiU'); bindText('sebabU', 'sebabU'); bindText('evNote', 'note');
    $('#aoiK').onchange = e => { setRec(p, { aoiK: e.target.value }); refreshNav(); };
    $('#sebabK').onchange = e => { setRec(p, { sebabK: e.target.value }); refreshNav(); };
    $('#klar').onchange = e => { setRec(p, { klarifikasi: e.target.checked }); refreshNav(); };
    $('#composeAoi').onclick = () => {
      const t = composeAoi(p);
      const cur = $('#aoiU').value.trim();
      if (cur && !confirm('Ganti Uraian AoI yang sudah ada dengan susunan otomatis?')) return;
      $('#aoiU').value = t; setRec(p, { aoiU: t }); refreshNav();
    };
    $$('.filechip', el).forEach(b => b.onclick = () => {
      const ev = evidenceFor(p);
      openViewer(ev.levels[b.dataset.l][+b.dataset.k], b.dataset.l, `${p.sub.kode} ${p.sub.nama}`);
      $$('.filechip.open', el).forEach(x => x.classList.remove('open')); b.classList.add('open');
    });
  }
  function keepScroll(fn) { const el = $('#content'); const y = el.scrollTop; fn(); el.scrollTop = y; }
  function refreshNav() {
    const y = $('#nav').scrollTop;
    $('#nav').innerHTML = navList();
    $('#navFilter').value = S.filter;
    $('#navFilter').onchange = e => { S.filter = e.target.value; render(); };
    $$('#nav a[data-i]').forEach(a => a.onclick = () => { S.idx = +a.dataset.i; S.view = 'param'; render(); $('#content').scrollTop = 0; });
    $('#nav').scrollTop = y;
  }

  function composeAoi(p) {
    const r = getRec(p);
    const g = r.tdn ? '' : r.grade || suggest(p, r);
    const lines = [];
    if (r.tdn) lines.push('Tidak dapat dinilai: bukti yang disajikan tidak memadai untuk memverifikasi grade PM.');
    const gi = g ? LEVELS.indexOf(g) : LEVELS.length;
    const next = gi > 0 ? LEVELS[gi - 1] : null;
    const lvNext = next && p.levels.find(l => l.lvl === next);
    if (lvNext) lines.push(`Untuk mencapai level ${next}: ${lvNext.kriteria.replace(/\s+/g, ' ').trim()}`);
    for (const l of p.levels) {
      const v = r.lv[l.lvl] || {};
      if ((v.v === 'tidak' || v.v === 'sebagian') && v.note) lines.push(`Level ${l.lvl} (${v.v === 'tidak' ? 'tidak terbukti' : 'sebagian'}): ${v.note.trim()}`);
    }
    return lines.join('\n');
  }

  // ---------- rekap ----------
  function kesimpulan(p, stage) {
    // modus grade antar satker (seri -> grade terendah), sama dengan rumus KK 3.1
    const cnt = { E: 0, D: 0, C: 0, B: 0, A: 0 };
    let any = false;
    for (const u of S.model.units) {
      let g = '';
      if (stage === 'PM') { const pu = p.units.find(x => x.unit.id === u.id); g = pu && pu.stages.PM ? pu.stages.PM.grade : ''; }
      else { const r = getRec(p, stage, u.id); g = r.tdn ? '' : r.grade; }
      if (cnt[g] != null) { cnt[g]++; any = true; }
    }
    if (!any) return '';
    let best = '', bc = 0;
    for (const k of ['E', 'D', 'C', 'B', 'A']) if (cnt[k] > bc) { bc = cnt[k]; best = k; }
    return best;
  }

  function renderRekap() {
    const u = unit();
    const pr = progress();
    let h = `<h1>Rekap ${STAGE_LABEL[S.stage]} · ${esc(u.label)}</h1>
      <div class="tiles">
        <div class="tile"><div class="v">${pr.done}/${pr.total}</div><div class="l">Parameter ditandai selesai</div><div class="bar"><i style="width:${(100 * pr.done / pr.total).toFixed(1)}%"></i></div></div>
        <div class="tile"><div class="v">${pr.filled}</div><div class="l">Parameter sudah bergrade ${STAGE_LABEL[S.stage]}</div></div>
        <div class="tile"><div class="v">${params().filter(p => getRec(p).klarifikasi).length}</div><div class="l">Perlu klarifikasi</div></div>
      </div>
      <h2>Per parameter</h2>
      <table class="grid"><thead><tr><th>Kode</th><th>No</th><th>Parameter</th><th class="PM">PM</th><th class="${S.stage}">${STAGE_LABEL[S.stage]}</th><th>Saran</th><th>AoI</th><th>Penyebab</th><th>Status</th></tr></thead><tbody>`;
    params().forEach((p, i) => {
      const r = getRec(p);
      const pm = (p.units.find(x => x.unit.id === u.id) || { stages: {} }).stages.PM;
      const pmg = pm ? pm.grade : '';
      const g = r.tdn ? 'TDN' : r.grade;
      const sug = suggest(p, r);
      const diff = g && pmg && g !== pmg && g !== 'TDN';
      const needAoi = g && g !== 'A';
      h += `<tr class="click" data-i="${i}"><td>${esc(p.sub.kode)}</td><td>${esc(p.no)}</td><td>${esc(p.uraian.slice(0, 110))}${p.uraian.length > 110 ? '…' : ''}</td>
        <td><span class="grade ${LEVELS.includes(pmg) ? 'PM' : 'empty'}">${esc(pmg) || '–'}</span></td>
        <td class="${diff ? 'diff' : ''}"><span class="grade ${g && g !== 'TDN' ? S.stage : 'empty'}" title="${g === 'TDN' ? 'Tidak dapat dinilai' : ''}">${g === 'TDN' ? '?' : g || '–'}</span></td>
        <td class="muted">${sug || ''}</td>
        <td>${r.aoiK || r.aoiU ? '✓' : needAoi ? '<span class="b warn">kosong</span>' : ''}</td>
        <td>${r.sebabK || r.sebabU ? '✓' : needAoi ? '<span class="b warn">kosong</span>' : ''}</td>
        <td>${r.done ? '<span class="b ok">selesai</span>' : hasInput(r) ? '<span class="b info">proses</span>' : ''}${r.klarifikasi ? ' <span class="b warn">klarifikasi</span>' : ''}</td></tr>`;
    });
    h += `</tbody></table>`;

    // skor subunsur
    h += `<h2>Skor subunsur KK 3.1</h2><p class="muted small">Kesimpulan parameter = grade terbanyak antar satker (seri → terendah); skor A=5…E=1; skor subunsur = rata-rata parameter. Ini hanya bagian KK 3.1 (T1) — KKLEAD II merata-ratakan dengan KK 3.2–3.4.</p>
      <table class="grid"><thead><tr><th>Kode</th><th>Subunsur</th><th class="PM">PM</th><th class="${S.stage}">${STAGE_LABEL[S.stage]}</th><th>Selisih</th></tr></thead><tbody>`;
    for (const sub of S.model.subs) {
      const sc = st => {
        const v = sub.params.map(p => SCORE[kesimpulan(p, st)]).filter(x => x != null);
        return v.length ? { v: v.reduce((a, b) => a + b, 0) / v.length, full: v.length === sub.params.length } : null;
      };
      const a = sc('PM'), b = sc(S.stage);
      const d = a && b ? b.v - a.v : null;
      h += `<tr><td><b>${esc(sub.kode)}</b></td><td>${esc(sub.nama)}</td><td class="num">${a ? fmt(a.v) + (a.full ? '' : '*') : '–'}</td><td class="num">${b ? fmt(b.v) + (b.full ? '' : '*') : '–'}</td>
        <td class="num ${d && d > 0 ? 'diff' : ''}">${d == null ? '' : (d > 0 ? '+' : '') + fmt(d)}</td></tr>`;
    }
    h += `</tbody></table><p class="muted small">* belum semua parameter bernilai.</p>
      <div style="margin-top:16px;display:flex;gap:8px"><button class="btn primary" id="rkWrite">Unduh KK terisi (.xlsx)</button><button class="btn" id="rkSheet">Ekspor lembar kerja evaluator (.xlsx)</button></div>`;
    $('#content').innerHTML = h;
    $$('tr.click').forEach(tr => tr.onclick = () => { S.idx = +tr.dataset.i; S.view = 'param'; render(); $('#content').scrollTop = 0; });
    $('#rkWrite').onclick = writeKK;
    $('#rkSheet').onclick = exportWorksheet;
  }

  // ---------- menulis ke KK ----------
  function collectCells() {
    const cells = [], summary = { grade: 0, aoi: 0, sebab: 0, units: new Set(), tdn: 0 };
    for (const u of S.model.units) {
      const cols = u.stages[S.stage];
      if (!cols) continue;
      for (const p of params()) {
        const k = recKey(p, S.stage, u.id);
        const r = S.ev[k];
        if (!r) continue; // hanya yang disentuh evaluator di aplikasi
        const put = (c, v) => { if (c >= 0 && v != null && String(v).trim() !== '') cells.push({ ref: P.addr(p.row, c), value: String(v).trim() }); };
        if (r.grade && !r.tdn) { put(cols.grade, r.grade); summary.grade++; }
        if (r.tdn) summary.tdn++;
        let aoiU = r.aoiU || '';
        if (r.tdn && !/tidak dapat dinilai/i.test(aoiU)) aoiU = 'Tidak dapat dinilai. ' + aoiU;
        if (r.aoiK || aoiU) summary.aoi++;
        if (r.sebabK || r.sebabU) summary.sebab++;
        put(cols.aoiK, r.aoiK); put(cols.aoiU, aoiU); put(cols.sebabK, r.sebabK); put(cols.sebabU, r.sebabU);
        summary.units.add(u.label);
      }
    }
    return { cells, summary };
  }

  async function writeKK() {
    const { cells, summary } = collectCells();
    if (!cells.length) { alert(`Belum ada isian ${STAGE_LABEL[S.stage]} yang dibuat di aplikasi.`); return; }
    const msg = `Menulis ke salinan "${S.fileName}", sheet ${S.sheetName}, tahap ${STAGE_LABEL[S.stage]}:\n` +
      `• ${summary.grade} grade${summary.tdn ? ` (+${summary.tdn} "tidak dapat dinilai": grade dikosongkan)` : ''}\n• ${summary.aoi} AoI · ${summary.sebab} penyebab\n• Satker: ${[...summary.units].join(', ')}\n\n` +
      `Hanya sel yang Anda isi yang ditulis; sel lain, format, rumus, dan sheet lain tidak berubah. Berkas asli tidak diubah. Lanjutkan?`;
    if (!confirm(msg)) return;
    busy(true, 'Menulis KK…');
    await tick();
    try {
      const blob = await XP.patch(S.file, S.sheetName, cells);
      // verifikasi: baca ulang hasil dan cocokkan setiap sel yang ditulis
      busy(true, 'Memeriksa hasil…');
      await tick();
      const wb = XLSX.read(await blob.arrayBuffer(), { sheets: [S.sheetName] });
      const ws = wb.Sheets[S.sheetName];
      const bad = cells.filter(c => { const x = ws[c.ref]; return !x || String(x.v).trim() !== c.value; });
      if (bad.length) throw new Error(`${bad.length} sel tidak tertulis dengan benar (mis. ${bad[0].ref}). Berkas tidak diunduh.`);
      const base = S.fileName.replace(/\.[^.]+$/, '');
      download(blob, `${base}_${S.stage === 'EV' ? 'Evaluasi' : 'PK'}_${dateStamp()}.xlsx`);
      toast(`${cells.length} sel ditulis & diverifikasi.`, 4000);
    } catch (e) { console.error(e); alert('Gagal menulis KK: ' + e.message); }
    busy(false);
  }

  function exportWorksheet() {
    const rows = [];
    for (const u of S.model.units) for (const p of params()) {
      const r = S.ev[recKey(p, S.stage, u.id)];
      if (!r) continue;
      const pm = (p.units.find(x => x.unit.id === u.id) || { stages: {} }).stages.PM;
      const ev = evidenceFor(p);
      const row = { Satker: u.label, 'Kode subunsur': p.sub.kode, Subunsur: p.sub.nama, No: p.no, Parameter: p.uraian, 'Grade PM': pm ? pm.grade : '' };
      for (const l of LEVELS) {
        const v = r.lv[l] || {};
        row[`Level ${l}`] = { ya: 'Terbukti', sebagian: 'Sebagian', tidak: 'Tidak terbukti' }[v.v] || '';
        row[`Catatan ${l}`] = v.note || '';
        if (ev) row[`Bukti ${l}`] = ev.levels[l].map(f => f.name).join('\n');
      }
      Object.assign(row, {
        'Saran grade': suggest(p, r), [`Grade ${STAGE_LABEL[S.stage]}`]: r.tdn ? 'Tidak dapat dinilai' : r.grade,
        'Kluster AoI': r.aoiK, 'Uraian AoI': r.aoiU, 'Kluster Penyebab': r.sebabK, 'Uraian Penyebab': r.sebabU,
        'Catatan evaluator': r.note, 'Perlu klarifikasi': r.klarifikasi ? 'Ya' : '', Selesai: r.done ? 'Ya' : '', 'Sel grade': `${S.sheetName}!${P.addr(p.row, (u.stages[S.stage] || {}).grade)}`,
      });
      rows.push(row);
    }
    if (!rows.length) { alert('Belum ada isian evaluator.'); return; }
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.json_to_sheet(rows);
    ws['!cols'] = Object.keys(rows[0]).map(k => ({ wch: /Parameter|Uraian|Catatan|Bukti/.test(k) ? 40 : 12 }));
    XLSX.utils.book_append_sheet(wb, ws, 'Lembar kerja ' + STAGE_LABEL[S.stage]);
    XLSX.writeFile(wb, `Lembar_kerja_${S.stage === 'EV' ? 'Evaluasi' : 'PK'}_KK31_${S.fileName.replace(/\.[^.]+$/, '')}_${dateStamp()}.xlsx`);
  }

  // ---------- sesi ----------
  function saveSession() {
    download(new Blob([JSON.stringify({ app: 'spip-kk31', version: 1, file: S.fileName, saved: new Date().toISOString(), ev: S.ev }, null, 1)], { type: 'application/json' }),
      `Sesi_KK31_${S.fileName.replace(/\.[^.]+$/, '')}_${dateStamp()}.json`);
  }
  async function loadSession(f) {
    try {
      const j = JSON.parse(await f.text());
      if (!j.ev) throw new Error('bukan berkas sesi KK 3.1');
      if (j.file && j.file !== S.fileName && !confirm(`Sesi dibuat untuk "${j.file}", berkas terbuka "${S.fileName}". Tetap gabungkan?`)) return;
      Object.assign(S.ev, j.ev); saveStore(); render();
      toast(`${Object.keys(j.ev).length} isian dimuat.`);
    } catch (e) { alert('Berkas sesi tidak valid: ' + e.message); }
  }

  // ---------- event ----------
  const fileInput = $('#fileInput');
  $('#btnOpen').onclick = () => fileInput.click();
  $('#btnOpenOther').onclick = () => { $('#menuMore').open = false; fileInput.click(); };
  fileInput.onchange = () => { openFile(fileInput.files[0]); fileInput.value = ''; };
  $('#btnBuktiFolder').onclick = () => { $('#menuBukti').open = false; $('#folderInput').click(); };
  $('#btnBuktiZip').onclick = () => { $('#menuBukti').open = false; $('#zipInput').click(); };
  $('#folderInput').onchange = e => { loadFolder(e.target.files); e.target.value = ''; };
  $('#zipInput').onchange = e => { loadZip(e.target.files[0]); e.target.value = ''; };
  $$('#stageSeg button').forEach(b => b.onclick = () => { S.stage = b.dataset.s; saveStore(); render(); });
  $('#unitSel').onchange = e => { S.unitId = e.target.value; saveStore(); render(); };
  $('#btnRekap').onclick = () => { S.view = S.view === 'rekap' ? 'param' : 'rekap'; render(); $('#content').scrollTop = 0; };
  $('#btnWrite').onclick = writeKK;
  $('#btnWorksheet').onclick = () => { $('#menuMore').open = false; exportWorksheet(); };
  $('#btnSaveSession').onclick = () => { $('#menuMore').open = false; saveSession(); };
  $('#btnLoadSession').onclick = () => { $('#menuMore').open = false; $('#sessionInput').click(); };
  $('#sessionInput').onchange = e => { if (e.target.files[0]) loadSession(e.target.files[0]); e.target.value = ''; };
  $('#btnClearSession').onclick = () => {
    $('#menuMore').open = false;
    if (!confirm('Hapus semua isian evaluator (semua tahap & satker) untuk berkas ini dari peramban?')) return;
    S.ev = {}; saveStore(); render();
  };
  $('#viewerClose').onclick = closeViewer;
  document.addEventListener('keydown', e => {
    if (e.key === 'Escape') { closeViewer(); return; }
    if (!S.model || S.view !== 'param') return;
    const t = e.target;
    if (t && (t.tagName === 'INPUT' || t.tagName === 'TEXTAREA' || t.tagName === 'SELECT' || t.isContentEditable) || e.ctrlKey || e.metaKey || e.altKey) return;
    const p = params()[S.idx];
    if (e.key === 'ArrowRight') { if (S.idx < params().length - 1) { S.idx++; render(); $('#content').scrollTop = 0; } e.preventDefault(); }
    else if (e.key === 'ArrowLeft') { if (S.idx > 0) { S.idx--; render(); $('#content').scrollTop = 0; } e.preventDefault(); }
    else if (/^[a-e]$/i.test(e.key)) { setRec(p, { grade: e.key.toUpperCase(), tdn: false }); keepScroll(renderParam); refreshNav(); }
  });
  const drop = $('#drop');
  ['dragenter', 'dragover'].forEach(ev => document.addEventListener(ev, e => { e.preventDefault(); drop.classList.add('over'); }));
  ['dragleave', 'drop'].forEach(ev => document.addEventListener(ev, e => { e.preventDefault(); if (ev === 'drop' || e.target === document.documentElement) drop.classList.remove('over'); }));
  document.addEventListener('drop', e => {
    const f = e.dataTransfer.files[0];
    if (!f) return;
    if (/\.zip$/i.test(f.name) && S.model) loadZip(f); else openFile(f);
  });

  // untuk pengujian otomatis
  window.KK31 = { openFile, loadZip, S, render };
})();
