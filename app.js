/* Antarmuka Asisten Evaluasi SPIP. Semua pemrosesan terjadi di peramban. */
(function () {
  'use strict';
  const P = window.SpipParser;
  const { STAGES, STAGE_LABEL, LEVELS, SCORE } = P;
  const $ = (s, el) => (el || document).querySelector(s);
  const $$ = (s, el) => Array.from((el || document).querySelectorAll(s));

  const state = {
    fileName: '', buf: null, sheetNames: [], hidden: {},
    models: [], byName: {}, combo: null,
    view: { kind: 'summary' }, reviews: {}, ui: {},
  };
  const STATUS = [
    ['', 'Belum'], ['ok', 'Sesuai'], ['klarifikasi', 'Klarifikasi'], ['tidak', 'Tidak sesuai'],
  ];
  const LEVEL_LABEL = { error: 'Galat', warn: 'Perlu perhatian', info: 'Info' };

  // ---------- util ----------
  function esc(s) {
    return String(s == null ? '' : s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  }
  function hl(text, q) {
    const e = esc(text);
    if (!q) return e;
    const rx = new RegExp('(' + q.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + ')', 'gi');
    return e.replace(rx, '<mark>$1</mark>');
  }
  function fmt(n, d) { return n == null || isNaN(n) ? '–' : Number(n).toFixed(d == null ? 2 : d).replace('.', ','); }
  function toast(msg) {
    const t = $('#toast'); t.textContent = msg; t.hidden = false;
    clearTimeout(toast._t); toast._t = setTimeout(() => { t.hidden = true; }, 2600);
  }
  function busy(on, text) { $('#busy').hidden = !on; if (text) $('#busyText').textContent = text; }
  // setTimeout (bukan requestAnimationFrame) agar tetap jalan saat tab tidak terlihat
  const nextFrame = () => new Promise(r => setTimeout(r, 30));

  // ---------- penyimpanan catatan reviu ----------
  const storeKey = () => 'spip-eval:v1:' + state.fileName;
  function loadReviews() {
    try { state.reviews = JSON.parse(localStorage.getItem(storeKey()) || '{}') || {}; } catch (e) { state.reviews = {}; }
  }
  function saveReviews() {
    try { localStorage.setItem(storeKey(), JSON.stringify(state.reviews)); } catch (e) { toast('Gagal menyimpan ke peramban — gunakan "Simpan sesi".'); }
  }
  function rev(key) { return state.reviews[key] || {}; }
  function setRev(key, patch) {
    const r = Object.assign({}, state.reviews[key] || {}, patch, { ts: new Date().toISOString() });
    if (!r.status && !r.grade && !r.note) delete state.reviews[key]; else state.reviews[key] = r;
    saveReviews();
  }
  const spKey = (p, unitId) => `${p.sheet}!${p.row}#${unitId}`;
  const rowKey = (m, row) => `${m.name}!${row.row}`;

  // ---------- membuka berkas ----------
  async function openFile(file) {
    if (!file) return;
    busy(true, 'Membaca daftar sheet…');
    await nextFrame();
    try {
      state.buf = await file.arrayBuffer();
      state.fileName = file.name;
      const wb = XLSX.read(state.buf, { bookSheets: true });
      state.sheetNames = wb.SheetNames;
      state.hidden = {};
      const meta = wb.Workbook && wb.Workbook.Sheets;
      if (meta) meta.forEach((s, i) => { if (s.Hidden) state.hidden[wb.SheetNames[i]] = true; });
    } catch (e) {
      busy(false); alert('Berkas tidak dapat dibaca: ' + e.message); return;
    }
    busy(false);
    $('#fileLabel').textContent = file.name;
    showSheetDialog();
  }

  function sheetGroup(n) {
    const s = n.toUpperCase().replace(/\s+/g, '');
    if (/^KK3\./.test(s)) return 'sp';
    if (/^KKE|^KKLEADI$/.test(s)) return 'pt';
    if (/^KK(4|5|6|7|8|9)|^KKLEADIII$/.test(s)) return 'ct';
    if (/^KKLEAD/.test(s)) return 'lead';
    return 'other';
  }

  function showSheetDialog() {
    const dlg = $('#sheetDialog');
    const loaded = new Set(state.models.map(m => m.name));
    const defaultOn = n => loaded.size ? loaded.has(n) : !state.hidden[n] && sheetGroup(n) !== 'other';
    $('#sheetDialogInfo').textContent = `${state.sheetNames.length} sheet ditemukan. Sheet referensi/tersembunyi tidak dicentang secara bawaan.`;
    $('#sheetList').innerHTML = state.sheetNames.map(n =>
      `<label><input type="checkbox" value="${esc(n)}" ${defaultOn(n) ? 'checked' : ''}> ${esc(n)} ${state.hidden[n] ? '<span class="hid">(tersembunyi)</span>' : ''}</label>`).join('');
    const presets = [
      ['Semua', () => true], ['Tidak ada', () => false],
      ['Struktur & Proses (KK 3.x)', n => sheetGroup(n) === 'sp' || /KKLEAD_SPIP/i.test(n)],
      ['Penetapan Tujuan (KKE)', n => sheetGroup(n) === 'pt'],
      ['Pencapaian Tujuan (KK 4–8)', n => sheetGroup(n) === 'ct'],
      ['Semua KK (tanpa referensi)', n => !state.hidden[n] && sheetGroup(n) !== 'other'],
    ];
    $('#sheetPresets').innerHTML = presets.map((p, i) => `<button type="button" class="btn sm" data-i="${i}">${esc(p[0])}</button>`).join('');
    $$('#sheetPresets button').forEach(b => b.onclick = () => {
      const f = presets[+b.dataset.i][1];
      $$('#sheetList input').forEach(cb => { cb.checked = f(cb.value); });
    });
    dlg.returnValue = '';
    dlg.showModal();
  }

  async function loadSelected(names) {
    if (!names.length) { toast('Tidak ada sheet dipilih.'); return; }
    busy(true, `Membaca ${names.length} sheet…`);
    await nextFrame();
    try {
      const wb = XLSX.read(state.buf, { dense: true, cellFormula: true, sheetStubs: true, sheets: names });
      busy(true, 'Menganalisis isi kertas kerja…');
      await nextFrame();
      state.models = names.filter(n => wb.Sheets[n]).map(n => P.parseSheet(wb.Sheets[n], n));
    } catch (e) {
      busy(false); console.error(e); alert('Gagal membaca sheet: ' + e.message); return;
    }
    state.byName = {};
    state.models.forEach(m => { state.byName[m.name] = m; });
    state.combo = P.combineSP(state.models);
    loadReviews();
    busy(false);
    $('#landing').hidden = true;
    $('#workspace').hidden = false;
    $('#btnExport').hidden = false;
    $('#btnSheets').hidden = false;
    $('#menuMore').hidden = false;
    state.ui = {};
    go({ kind: 'summary' });
  }

  // ---------- navigasi ----------
  function progressOf(m) {
    if (m.type !== 'sp') return null;
    let total = 0, done = 0;
    for (const p of m.params) for (const u of m.activeUnits) { total++; if (rev(spKey(p, u.id)).status) done++; }
    return { total, done };
  }
  function flagBadges(fc) {
    if (!fc) return '';
    return (fc.error ? `<span class="b error" title="Galat">${fc.error}</span>` : '') + (fc.warn ? `<span class="b warn" title="Perlu perhatian">${fc.warn}</span>` : '');
  }
  function renderNav() {
    const v = state.view;
    const groups = [
      ['Struktur & Proses', m => m.type === 'sp'],
      ['Rekap nilai', m => m.type === 'lead'],
      ['Tabel kertas kerja', m => m.type === 'table'],
      ['Lainnya', m => m.type === 'info'],
    ];
    let h = `<a data-kind="summary" class="${v.kind === 'summary' ? 'active' : ''}"><span class="nm">Ringkasan</span></a>`;
    const tot = totalFlags();
    h += `<a data-kind="findings" class="${v.kind === 'findings' ? 'active' : ''}"><span class="nm">Semua temuan otomatis</span>${flagBadges(tot)}</a>`;
    for (const [title, f] of groups) {
      const ms = state.models.filter(f);
      if (!ms.length) continue;
      h += `<h4>${esc(title)}</h4>`;
      for (const m of ms) {
        const pr = progressOf(m);
        h += `<a data-kind="sheet" data-name="${esc(m.name)}" class="${v.kind === 'sheet' && v.name === m.name ? 'active' : ''}"><span class="nm">${esc(m.name)}</span>${pr ? `<span class="prog">${pr.done}/${pr.total}</span>` : ''}${flagBadges(m.flagCount)}</a>`;
      }
    }
    $('#nav').innerHTML = h;
    $$('#nav a').forEach(a => a.onclick = () => go(a.dataset.kind === 'sheet' ? { kind: 'sheet', name: a.dataset.name } : { kind: a.dataset.kind }));
  }
  function totalFlags() {
    const c = { error: 0, warn: 0, info: 0 };
    for (const m of state.models) if (m.flagCount) for (const k in c) c[k] += m.flagCount[k];
    if (state.combo) for (const f of state.combo.flags) c[f.level]++;
    return c;
  }

  function go(view, opts) {
    state.view = view;
    closeDrawer();
    renderNav();
    const el = $('#content');
    if (view.kind === 'summary') renderSummary(el);
    else if (view.kind === 'findings') renderFindings(el);
    else {
      const m = state.byName[view.name];
      if (m.type === 'sp') renderSP(el, m);
      else if (m.type === 'table') renderTable(el, m);
      else if (m.type === 'lead') renderLead(el, m);
      else renderInfo(el, m);
    }
    if (!(opts && opts.keepScroll)) el.scrollTop = 0;
  }

  // ---------- ringkasan ----------
  function renderSummary(el) {
    const tot = totalFlags();
    let items = 0, done = 0;
    state.models.forEach(m => { const p = progressOf(m); if (p) { items += p.total; done += p.done; } });
    const sp = state.models.filter(m => m.type === 'sp');
    let h = `<h1>Ringkasan</h1><div class="muted">${esc(state.fileName)} · ${state.models.length} sheet dibaca</div>`;
    h += `<div class="tiles">
      <div class="tile"><div class="v">${sp.reduce((a, m) => a + m.params.length, 0)}</div><div class="l">Parameter Struktur &amp; Proses (${sp.length} KK)</div></div>
      <div class="tile"><div class="v">${tot.error + tot.warn}</div><div class="l">Temuan perlu perhatian <span class="b error">${tot.error}</span> <span class="b warn">${tot.warn}</span> · info ${tot.info}</div></div>
      <div class="tile"><div class="v">${done}/${items}</div><div class="l">Parameter × satker sudah direviu</div><div class="bar"><i style="width:${items ? (100 * done / items).toFixed(1) : 0}%"></i></div></div>
    </div>`;

    // tabel per sheet
    h += `<h2>Per sheet</h2><table class="grid"><thead><tr><th>Sheet</th><th>Jenis</th><th>Isi</th><th>Galat</th><th>Perhatian</th><th>Info</th><th>Reviu</th></tr></thead><tbody>`;
    const jenis = { sp: 'Struktur & Proses', table: 'Tabel KK', lead: 'Rekap nilai', info: 'Informasi' };
    for (const m of state.models) {
      const pr = progressOf(m);
      const isi = m.type === 'sp' ? `${m.params.length} parameter · ${m.activeUnits.map(u => esc(u.label)).join(', ')}`
        : m.type === 'table' ? `${m.rows.filter(r => !r.section).length} baris data` : m.type === 'lead' ? `${m.rows.length} baris` : `${m.lines.length} baris teks`;
      h += `<tr class="click" data-name="${esc(m.name)}"><td><b>${esc(m.name)}</b></td><td>${jenis[m.type]}</td><td class="muted">${isi}</td>
        <td class="num">${m.flagCount.error || ''}</td><td class="num">${m.flagCount.warn || ''}</td><td class="num">${m.flagCount.info || ''}</td>
        <td>${pr ? `${pr.done}/${pr.total}` : ''}</td></tr>`;
    }
    h += `</tbody></table>`;

    // rekap skor
    if (state.combo && state.combo.rows.length) h += scoreTableHtml();
    el.innerHTML = h;
    $$('tr.click', el).forEach(tr => tr.onclick = () => go({ kind: 'sheet', name: tr.dataset.name }));
    $$('[data-jump]', el).forEach(a => a.onclick = e => { e.preventDefault(); jump(a.dataset.jump); });
  }

  function leadModel() {
    const leads = state.models.filter(m => m.type === 'lead');
    return leads.find(m => /kklead/i.test(m.name)) || leads[0];
  }

  function scoreTableHtml() {
    const c = state.combo, lead = leadModel();
    // tampilkan tahap yang punya nilai hasil hitung ulang atau nilai di KKLEAD
    const stg = STAGES.filter(s => c.rows.some(r => r.avg[s] != null) || (lead && lead.rows.some(x => x.vals[s] && (x.vals[s].value != null))));
    if (!stg.length) stg.push('PM');
    let h = `<h2>Rekap skor subunsur Struktur &amp; Proses (dihitung ulang)</h2>
      <p class="muted small">Skor parameter: A=5 … E=1. Kesimpulan parameter KK 3.1 = grade terbanyak antar satker (seri → grade terendah, sama dengan rumus KK). Skor subunsur per KK = rata-rata parameter; skor gabungan = rata-rata KK 3.1–3.4 (seperti KKLEAD II, sebelum veto KK 4).
      ${lead ? 'Kolom "di KKLEAD" = nilai yang tercatat di sheet ' + esc(lead.name) + '.' : ''}</p>`;
    if (c.flags.length) h += `<ul class="flags" style="margin-bottom:10px">${c.flags.map(f => `<li class="${f.level}">${esc(f.msg)} <span class="cell">${esc(f.cell)}</span></li>`).join('')}</ul>`;
    h += `<div class="scroll-x"><table class="grid score"><thead><tr><th rowspan="2">Kode</th><th rowspan="2">Subunsur</th>`;
    for (const sh of c.sheets) h += `<th colspan="${stg.length}">${esc(sh)}</th>`;
    h += `<th colspan="${stg.length}">Gabungan</th>${lead ? `<th colspan="${stg.length}">Di KKLEAD</th>` : ''}</tr><tr>`;
    const heads = stg.map(s => `<th class="${s}">${STAGE_LABEL[s]}</th>`).join('');
    for (let i = 0; i < c.sheets.length + 1 + (lead ? 1 : 0); i++) h += heads;
    h += `</tr></thead><tbody>`;
    for (const r of c.rows) {
      h += `<tr><td><b>${esc(r.kode)}</b></td><td>${esc(r.nama)}</td>`;
      for (const sh of c.sheets) for (const s of stg) {
        const v = r.per[sh] && r.per[sh][s];
        h += `<td class="num" title="${v ? `${v.n}/${v.of} parameter bernilai` : 'tidak ada'}">${v ? fmt(v.value) + (v.complete ? '' : '*') : '–'}</td>`;
      }
      for (const s of stg) h += `<td class="num"><b>${fmt(r.avg[s])}</b></td>`;
      if (lead) {
        const lr = lead.rows.find(x => x.kode === r.kode);
        for (const s of stg) {
          const v = lr && lr.vals[s];
          let cell = '–', cls = '';
          if (v && v.formula) cell = '<span class="muted" title="Rumus; nilai tidak tersimpan di berkas">rumus</span>';
          else if (v && v.value != null) {
            cell = fmt(v.value) + ' <span class="b warn" title="Diketik manual">manual</span>';
            if (r.avg[s] != null && Math.abs(r.avg[s] - v.value) > 0.005) cls = 'diff';
          }
          h += `<td class="num ${cls}">${cell}</td>`;
        }
      }
      h += `</tr>`;
    }
    h += `</tbody></table></div><p class="muted small">* sebagian parameter belum bernilai — di Excel rumus subunsur akan kosong sampai semua parameter terisi.</p>`;
    return h;
  }

  // ---------- semua temuan ----------
  function allFindings() {
    const out = [];
    for (const m of state.models) {
      if (m.type === 'sp') {
        for (const f of m.sheetFlags) out.push({ m, f, where: '', jump: `${m.name}` });
        for (const p of m.params) for (const f of p.flags) out.push({ m, f, p, where: `${p.sub.kode} · no.${p.no}`, jump: `${m.name}|p|${p.row}|${f.unit || ''}` });
      } else if (m.type === 'table') {
        for (const r of m.rows) for (const f of (r.flags || [])) out.push({ m, f, r, where: `baris ${r.row + 1}`, jump: `${m.name}|r|${r.row}` });
      } else if (m.type === 'lead') {
        for (const f of m.flags) out.push({ m, f, where: `baris ${f.row + 1}`, jump: `${m.name}` });
      }
    }
    if (state.combo) for (const f of state.combo.flags) out.push({ m: state.byName[f.sheet] || { name: f.sheet }, f, where: '', jump: f.sheet });
    return out;
  }

  function renderFindings(el) {
    const ui = state.ui.findings || (state.ui.findings = { level: 'warn', code: '', sheet: '' });
    const all = allFindings();
    const codes = [...new Set(all.map(x => x.f.code))].sort();
    const order = { error: 0, warn: 1, info: 2 };
    const list = all.filter(x => (ui.level === 'all' || (ui.level === 'warn' ? order[x.f.level] <= 1 : x.f.level === ui.level)) && (!ui.code || x.f.code === ui.code) && (!ui.sheet || x.m.name === ui.sheet));
    let h = `<h1>Semua temuan otomatis</h1><p class="muted">Temuan adalah petunjuk untuk diperiksa, bukan kesimpulan. Klik baris untuk membuka parameter/baris terkait.</p>
      <div class="toolbar">
        <label>Tingkat <select id="fLevel">
          <option value="warn">Galat + perlu perhatian</option><option value="all">Semua</option><option value="error">Galat</option><option value="info">Info</option></select></label>
        <label>Jenis <select id="fCode"><option value="">Semua jenis</option>${codes.map(c => `<option>${esc(c)}</option>`).join('')}</select></label>
        <label>Sheet <select id="fSheet"><option value="">Semua sheet</option>${state.models.map(m => `<option>${esc(m.name)}</option>`).join('')}</select></label>
        <span class="spacer"></span><span class="muted">${list.length} temuan</span>
      </div>`;
    if (!list.length) h += `<div class="empty-note">Tidak ada temuan untuk filter ini.</div>`;
    else {
      h += `<table class="grid"><thead><tr><th>Tingkat</th><th>Sheet</th><th>Lokasi</th><th>Jenis</th><th>Keterangan</th><th>Sel</th></tr></thead><tbody>`;
      for (const x of list.slice(0, 1500))
        h += `<tr class="click" data-jump="${esc(x.jump)}"><td><span class="b ${x.f.level}">${LEVEL_LABEL[x.f.level]}</span></td><td>${esc(x.m.name)}</td><td>${esc(x.where)}</td><td class="small">${esc(x.f.code)}</td><td>${esc(x.f.msg)}</td><td class="cell">${esc(x.f.cell)}</td></tr>`;
      h += `</tbody></table>`;
    }
    el.innerHTML = h;
    $('#fLevel').value = ui.level; $('#fCode').value = ui.code; $('#fSheet').value = ui.sheet;
    $('#fLevel').onchange = e => { ui.level = e.target.value; renderFindings(el); };
    $('#fCode').onchange = e => { ui.code = e.target.value; renderFindings(el); };
    $('#fSheet').onchange = e => { ui.sheet = e.target.value; renderFindings(el); };
    $$('tr[data-jump]', el).forEach(tr => tr.onclick = () => jump(tr.dataset.jump));
  }

  function jump(spec) {
    const [name, kind, row, unit] = spec.split('|');
    const m = state.byName[name];
    if (!m) return;
    if (kind === 'p') {
      const ui = spUI(m);
      ui.sub = 'all'; ui.filter = 'all'; ui.q = '';
      if (unit) ui.unit = unit;
      go({ kind: 'sheet', name });
      const card = document.getElementById(`p-${cssId(name)}-${row}`);
      if (card) { card.scrollIntoView({ block: 'start' }); card.classList.add('flash'); setTimeout(() => card.classList.remove('flash'), 1600); }
    } else if (kind === 'r') {
      const ui = tableUI(m); ui.filter = 'all'; ui.q = '';
      const idx = m.rows.findIndex(r => r.row === +row);
      ui.page = Math.floor(idx / ui.per);
      go({ kind: 'sheet', name });
      openRowDrawer(m, m.rows[idx]);
    } else go({ kind: 'sheet', name });
  }
  const cssId = s => s.replace(/[^A-Za-z0-9]/g, '_');

  // ---------- Struktur & Proses ----------
  function spUI(m) {
    if (state.ui[m.name]) return state.ui[m.name];
    // tahap fokus bawaan: tahap terakhir yang sudah ada grade-nya
    let stage = 'PM';
    for (const s of STAGES) if (m.params.some(p => p.units.some(u => u.stages[s] && LEVELS.includes(u.stages[s].grade)))) stage = s;
    return (state.ui[m.name] = { sub: 'all', unit: (m.activeUnits[0] || m.units[0]).id, stage, filter: 'all', q: '', compare: false, detail: false });
  }

  function renderSP(el, m) {
    const ui = spUI(m);
    const stagesAvail = STAGES.filter(s => m.units.some(u => u.stages[s]));
    const unitSel = m.activeUnits.length > 1 || m.units.length > 1;
    let h = `<h1>${esc(m.name)} <span class="muted small">Struktur &amp; Proses</span></h1>
      <div class="muted">${m.params.length} parameter dalam ${m.subs.length} subunsur · unit terisi: ${m.activeUnits.map(u => esc(u.label)).join(', ') || '–'}</div>`;
    if (m.sheetFlags.length) h += `<ul class="flags" style="margin-top:10px">${m.sheetFlags.map(f => `<li class="${f.level}">${esc(f.msg)} <span class="cell">${esc(f.cell)}</span></li>`).join('')}</ul>`;
    h += `<div class="toolbar">
      <label>Subunsur <select id="spSub"><option value="all">Semua subunsur</option>${m.subs.map((s, i) => `<option value="${i}">${esc(s.kode)} ${esc(s.nama)}</option>`).join('')}</select></label>
      ${unitSel ? `<label>Satker <select id="spUnit">${m.units.map(u => `<option value="${u.id}">${esc(u.label)}${m.activeUnits.includes(u) ? '' : ' (kosong)'}</option>`).join('')}</select></label>` : ''}
      <label>Tahap <span class="seg" id="spStage">${stagesAvail.map(s => `<button type="button" data-s="${s}" class="${ui.stage === s ? 'on' : ''}">${STAGE_LABEL[s]}</button>`).join('')}</span></label>
      <label>Tampilkan <select id="spFilter">
        <option value="all">Semua parameter</option><option value="flag">Ada temuan (perhatian/galat)</option><option value="anyflag">Ada temuan apa pun</option>
        <option value="unrev">Belum direviu</option><option value="klar">Perlu klarifikasi / tidak sesuai</option></select></label>
      <input type="search" id="spQ" placeholder="Cari teks parameter/uraian…" value="${esc(ui.q)}">
      <label title="Tampilkan penjelasan & cara pengujian di setiap level"><input type="checkbox" id="spDetail" ${ui.detail ? 'checked' : ''}> Penjelasan</label>
      ${m.units.some(u => !u.sharedUraian) && stagesAvail.length > 1 ? `<label title="Tampilkan uraian semua tahap berdampingan"><input type="checkbox" id="spCmp" ${ui.compare ? 'checked' : ''}> Bandingkan tahap</label>` : ''}
    </div><div id="spCards"></div>`;
    el.innerHTML = h;
    $('#spSub').value = ui.sub;
    if ($('#spUnit')) { $('#spUnit').value = ui.unit; $('#spUnit').onchange = e => { ui.unit = e.target.value; drawCards(); }; }
    $('#spFilter').value = ui.filter;
    $('#spSub').onchange = e => { ui.sub = e.target.value; drawCards(); };
    $('#spFilter').onchange = e => { ui.filter = e.target.value; drawCards(); };
    $$('#spStage button').forEach(b => b.onclick = () => { ui.stage = b.dataset.s; $$('#spStage button').forEach(x => x.classList.toggle('on', x === b)); drawCards(); });
    let t; $('#spQ').oninput = e => { clearTimeout(t); t = setTimeout(() => { ui.q = e.target.value.trim(); drawCards(); }, 200); };
    $('#spDetail').onchange = e => { ui.detail = e.target.checked; drawCards(); };
    if ($('#spCmp')) $('#spCmp').onchange = e => { ui.compare = e.target.checked; drawCards(); };

    function drawCards() {
      const box = $('#spCards');
      const unit = m.units.find(u => u.id === ui.unit) || m.units[0];
      const q = ui.q.toLowerCase();
      let html = '', shown = 0, lastSub = null;
      m.subs.forEach((sub, si) => {
        if (ui.sub !== 'all' && +ui.sub !== si) return;
        for (const p of sub.params) {
          const pu = p.units.find(x => x.unit.id === unit.id);
          const fl = p.flags.filter(f => !f.unit || f.unit === unit.id);
          const r = rev(spKey(p, unit.id));
          if (ui.filter === 'flag' && !fl.some(f => f.level !== 'info')) continue;
          if (ui.filter === 'anyflag' && !fl.length) continue;
          if (ui.filter === 'unrev' && r.status) continue;
          if (ui.filter === 'klar' && !(r.status === 'klarifikasi' || r.status === 'tidak')) continue;
          if (q) {
            const hay = [p.uraian, ...p.levels.map(l => l.kriteria)].concat(pu ? Object.values(pu.stages).flatMap(st => st.uraian.map(x => x.text).concat([st.aoiU, st.sebabU])) : []).join(' ').toLowerCase();
            if (!hay.includes(q)) continue;
          }
          if (lastSub !== sub) {
            const sc = STAGES.filter(s => sub.score[s]).map(s => `${STAGE_LABEL[s]} ${fmt(sub.score[s].value)}${sub.score[s].complete ? '' : '*'}`).join(' · ');
            html += `<div class="subhead"><b>${esc(sub.kode)} ${esc(sub.nama)}</b><span class="sc">${sc ? 'skor ' + sc : ''}</span></div>`;
            lastSub = sub;
          }
          html += cardHtml(m, p, pu, unit, fl, ui);
          shown++;
        }
      });
      box.innerHTML = shown ? html : `<div class="empty-note">Tidak ada parameter yang cocok dengan filter.</div>`;
      bindCards(box, m);
    }
    drawCards();
  }

  function gradeChip(g, s) {
    return LEVELS.includes(g) ? `<span class="grade ${s}">${g}</span>` : `<span class="grade empty">${g ? esc(g) : '–'}</span>`;
  }

  function cardHtml(m, p, pu, unit, flags, ui) {
    const r = rev(spKey(p, unit.id));
    const stages = pu ? STAGES.filter(s => pu.stages[s]) : [];
    const focus = pu && pu.stages[ui.stage] ? ui.stage : stages[0];
    const fst = pu && pu.stages[focus];
    const gi = fst && LEVELS.includes(fst.grade) ? LEVELS.indexOf(fst.grade) : 99;
    const q = ui.q;
    let h = `<article class="pcard" id="p-${cssId(m.name)}-${p.row}" data-row="${p.row}" data-unit="${unit.id}">
      <header><div>
        <div class="kode">${esc(p.sub.kode)} · Parameter ${esc(p.no)} ${p.kodeParam.map(k => `<span class="tag">${esc(k)}</span>`).join(' ')} <span class="cell muted">${esc(m.name)}!${P.addr(p.row, 3)}</span> <span class="st-dot ${r.status || ''}" title="Status reviu"></span></div>
        <h3>${hl(p.uraian, q)}</h3></div>
        <div class="grades">${stages.map(s => `<div class="g">${m.activeUnits.length > 1 ? esc(unit.label) + '<br>' : ''}${STAGE_LABEL[s]}${gradeChip(pu.stages[s].grade, s)}</div>`).join('')}
          ${m.activeUnits.length > 1 ? `<div class="g kes">Kesimpulan<br><span style="display:flex;gap:4px">${STAGES.filter(s => p.kes[s]).map(s => `<span title="${STAGE_LABEL[s]} (hitung ulang)">${gradeChip(p.kes[s], s)}</span>`).join('') || '–'}</span></div>` : ''}
        </div>
      </header>`;
    // tangga level
    const cmpStages = ui.compare && pu && !unit.sharedUraian ? stages : [focus];
    h += `<div class="pbody"><div class="pmain"><table class="ladder"><thead><tr><th>Level</th><th>Kriteria</th>${cmpStages.map(s => `<th>Uraian hasil pengujian${unit.sharedUraian ? '' : ' · ' + STAGE_LABEL[s]}</th>`).join('')}</tr></thead><tbody>`;
    p.levels.forEach((l, i) => {
      const claim = LEVELS.indexOf(l.lvl) >= gi;
      const marks = stages.filter(s => pu.stages[s].grade === l.lvl).map(s => `<span class="stg ${s}" title="Grade ${STAGE_LABEL[s]}">${STAGE_LABEL[s]}</span>`).join('');
      h += `<tr class="${claim ? 'claim' : ''}"><td class="lv"><span class="lvb">${l.lvl}</span><div class="marks">${marks}</div></td>
        <td class="kr">${hl(l.kriteria, q)}${ui.detail && (l.penjelasan || l.cara) ? `<div class="pj">${hl(l.penjelasan, q)}${l.cara ? `\nCara uji: ${esc(l.cara)}` : ''}</div>` : ''}</td>`;
      for (const s of cmpStages) {
        const u = pu && pu.stages[s] ? pu.stages[s].uraian[i] : null;
        const txt = u ? u.text : '';
        h += `<td class="ur">${!txt ? '<span class="ph">kosong</span>' : P.isPlaceholder(txt) ? `<span class="ph">template: ${esc(txt.replace(/\s+/g, ' '))}</span>` : hl(txt, q)}</td>`;
      }
      h += `</tr>`;
    });
    h += `</tbody></table></div><aside class="pside">`;
    // AoI & penyebab
    if (fst) {
      const aoiReal = P.isReal(fst.aoiU) || P.isReal(fst.aoiK), sebabReal = P.isReal(fst.sebabU) || P.isReal(fst.sebabK);
      h += `<div><h4>AoI · ${STAGE_LABEL[focus]}</h4><div class="aoi">${aoiReal ? `${fst.aoiK ? `<span class="k">${esc(fst.aoiK)}</span>\n` : ''}${hl(P.isReal(fst.aoiU) ? fst.aoiU : '', q)}` : '<span class="muted">kosong</span>'}</div></div>
        <div><h4>Penyebab · ${STAGE_LABEL[focus]}</h4><div class="aoi">${sebabReal ? `${fst.sebabK ? `<span class="k">${esc(fst.sebabK)}</span>\n` : ''}${hl(P.isReal(fst.sebabU) ? fst.sebabU : '', q)}` : '<span class="muted">kosong</span>'}</div></div>`;
    }
    if (flags.length) h += `<div><h4>Temuan otomatis</h4><ul class="flags">${flags.map(f => `<li class="${f.level}">${esc(f.msg)}${f.cell ? `<span class="cell">${esc(f.cell)}</span>` : ''}</li>`).join('')}</ul></div>`;
    else if (pu && pu.active) h += `<div class="muted small">Tidak ada temuan otomatis.</div>`;
    // reviu
    h += `<div class="review"><h4>Reviu evaluator</h4>
      <span class="seg rv-status">${STATUS.map(([v, lab]) => `<button type="button" data-v="${v}" class="${(r.status || '') === v ? 'on ' + v : ''}">${lab}</button>`).join('')}</span>
      <label class="small">Grade menurut evaluator <select class="rv-grade"><option value="">–</option>${LEVELS.map(g => `<option ${r.grade === g ? 'selected' : ''}>${g}</option>`).join('')}</select></label>
      <textarea class="rv-note" placeholder="Catatan / hal yang perlu diklarifikasi…">${esc(r.note || '')}</textarea></div>`;
    h += `</aside></div></article>`;
    return h;
  }

  function bindCards(box, m) {
    $$('.pcard', box).forEach(card => {
      const p = m.params.find(x => x.row === +card.dataset.row);
      const key = spKey(p, card.dataset.unit);
      $$('.rv-status button', card).forEach(b => b.onclick = () => {
        setRev(key, { status: b.dataset.v });
        $$('.rv-status button', card).forEach(x => { x.className = x === b ? 'on ' + b.dataset.v : ''; });
        $('.st-dot', card).className = 'st-dot ' + b.dataset.v;
        renderNav();
      });
      $('.rv-grade', card).onchange = e => setRev(key, { grade: e.target.value });
      let t; $('.rv-note', card).oninput = e => { clearTimeout(t); t = setTimeout(() => setRev(key, { note: e.target.value }), 300); };
    });
  }

  // ---------- tabel KK ----------
  function tableUI(m) {
    return state.ui[m.name] || (state.ui[m.name] = { filter: 'all', q: '', page: 0, per: 100, stages: Object.fromEntries(STAGES.map(s => [s, true])) });
  }
  function renderTable(el, m) {
    const ui = tableUI(m);
    const q = ui.q.toLowerCase();
    const rows = m.rows.filter(r => {
      if (ui.filter === 'flag' && !(r.flags || []).length) return false;
      if (ui.filter === 'unrev' && (r.section || rev(rowKey(m, r)).status)) return false;
      if (q && !r.vals.join(' ').toLowerCase().includes(q)) return false;
      return true;
    });
    const pages = Math.max(1, Math.ceil(rows.length / ui.per));
    if (ui.page >= pages) ui.page = pages - 1;
    const slice = rows.slice(ui.page * ui.per, ui.page * ui.per + ui.per);
    const cols = m.cols.map((c, i) => ({ c, i })).filter(({ c }) => !c.stage || ui.stages[c.stage]);
    // indeks kolom pembanding tahap sebelumnya untuk sorot perbedaan
    const prevIdx = {};
    for (const [, o] of m.compareKeys) {
      const pres = STAGES.filter(s => o[s] != null);
      for (let k = 1; k < pres.length; k++) prevIdx[o[pres[k]]] = o[pres[k - 1]];
    }
    let h = `<h1>${esc(m.name)}</h1><div class="muted">${m.rows.filter(r => !r.section).length} baris data · ${m.cols.length} kolom${m.stages.length ? ' · tahap: ' + m.stages.map(s => STAGE_LABEL[s]).join(', ') : ''}. Klik baris untuk melihat detail vertikal &amp; memberi catatan.</div>
      <div class="toolbar">
        <label>Tampilkan <select id="tFilter"><option value="all">Semua baris</option><option value="flag">Ada temuan</option><option value="unrev">Belum direviu</option></select></label>
        ${m.stages.map(s => `<label><input type="checkbox" data-s="${s}" class="tStage" ${ui.stages[s] ? 'checked' : ''}> <span class="stg ${s}">${STAGE_LABEL[s]}</span></label>`).join('')}
        <input type="search" id="tQ" placeholder="Cari…" value="${esc(ui.q)}">
        <span class="spacer"></span>
        <span class="muted">${rows.length} baris</span>
        ${pages > 1 ? `<button class="btn sm" id="tPrev" ${ui.page ? '' : 'disabled'}>‹</button><span>${ui.page + 1}/${pages}</span><button class="btn sm" id="tNext" ${ui.page < pages - 1 ? '' : 'disabled'}>›</button>` : ''}
      </div>`;
    if (!m.rows.length) h += `<div class="empty-note">Tabel ini belum berisi data.</div>`;
    else {
      h += `<div class="scroll-x"><table class="grid wide"><thead><tr><th></th>${cols.map(({ c }) => `<th class="${c.stage || ''}" title="${esc(c.label)} (${c.name})">${c.stage ? `<span class="stg ${c.stage}">${STAGE_LABEL[c.stage]}</span><br>` : ''}${esc(c.label)}</th>`).join('')}</tr></thead><tbody>`;
      for (const r of slice) {
        if (r.section) { h += `<tr class="sec"><td></td><td colspan="${cols.length}">${esc(r.vals.find(v => v))}</td></tr>`; continue; }
        const fc = { error: 0, warn: 0, info: 0 }; (r.flags || []).forEach(f => fc[f.level]++);
        const st = rev(rowKey(m, r)).status || '';
        h += `<tr class="click" data-row="${r.row}"><td style="white-space:nowrap"><span class="st-dot ${st}"></span> ${flagBadges(fc)}${fc.info ? `<span class="b info">${fc.info}</span>` : ''}</td>`;
        for (const { c, i } of cols) {
          const v = r.vals[i];
          const d = prevIdx[i] != null && v && r.vals[prevIdx[i]] && v.toLowerCase() !== r.vals[prevIdx[i]].toLowerCase();
          h += `<td class="${d ? 'diff' : ''}"><div class="clip">${hl(v, ui.q)}</div></td>`;
        }
        h += `</tr>`;
      }
      h += `</tbody></table></div>`;
    }
    if (m.notes.length) h += `<details style="margin-top:14px"><summary>Catatan / petunjuk / simpulan di bawah tabel (${m.notes.length} baris)</summary><div class="infolist">${m.notes.map(n => `<div><span class="cell">${n.row + 1}</span>${esc(n.text)}</div>`).join('')}</div></details>`;
    el.innerHTML = h;
    $('#tFilter').value = ui.filter;
    $('#tFilter').onchange = e => { ui.filter = e.target.value; ui.page = 0; renderTable(el, m); };
    $$('.tStage', el).forEach(cb => cb.onchange = () => { ui.stages[cb.dataset.s] = cb.checked; renderTable(el, m); });
    let t; $('#tQ').oninput = e => { clearTimeout(t); t = setTimeout(() => { ui.q = e.target.value.trim(); ui.page = 0; renderTable(el, m); $('#tQ').focus(); }, 300); };
    if ($('#tPrev')) { $('#tPrev').onclick = () => { ui.page--; renderTable(el, m); }; $('#tNext').onclick = () => { ui.page++; renderTable(el, m); }; }
    $$('tr.click', el).forEach(tr => tr.onclick = () => openRowDrawer(m, m.rows.find(r => r.row === +tr.dataset.row)));
  }

  function closeDrawer() { const d = $('#drawer'); if (d) d.remove(); }
  function openRowDrawer(m, r) {
    closeDrawer();
    if (!r) return;
    const key = rowKey(m, r), rv = rev(key);
    const ident = m.cols.map((c, i) => ({ c, v: r.vals[i] })).filter(x => !x.c.stage && x.v);
    // kelompokkan kolom bertahap per label
    const groups = new Map();
    m.cols.forEach((c, i) => { if (!c.stage) return; if (!groups.has(c.key)) groups.set(c.key, { label: c.label.replace(/\b(PM|PK|EVALUASI)\b/gi, '').trim(), v: {} }); groups.get(c.key).v[c.stage] = { val: r.vals[i], cell: P.addr(r.row, c.c) }; });
    const st = m.stages;
    let h = `<button class="btn sm close" id="dClose">Tutup ✕</button><h2 style="margin-top:0">${esc(m.name)} · baris ${r.row + 1}</h2>
      <div class="kv">${ident.map(x => `<div class="k">${esc(x.c.label)}</div><div>${esc(x.v)}</div>`).join('')}</div>`;
    if (groups.size) {
      h += `<h3>Isian per tahap</h3><table class="grid cmp"><thead><tr><th>Kolom</th>${st.map(s => `<th class="${s}">${STAGE_LABEL[s]}</th>`).join('')}</tr></thead><tbody>`;
      for (const g of groups.values()) {
        if (!st.some(s => g.v[s] && g.v[s].val)) continue;
        h += `<tr><td class="muted">${esc(g.label)}</td>${st.map((s, k) => {
          const x = g.v[s], prev = k ? g.v[st[k - 1]] : null;
          const d = x && prev && x.val && prev.val && x.val.toLowerCase() !== prev.val.toLowerCase();
          return `<td class="${d ? 'diff' : ''}" style="white-space:pre-line">${x ? esc(x.val) : ''}</td>`;
        }).join('')}</tr>`;
      }
      h += `</tbody></table>`;
    }
    if ((r.flags || []).length) h += `<h3>Temuan otomatis</h3><ul class="flags">${r.flags.map(f => `<li class="${f.level}">${esc(f.msg)} <span class="cell">${esc(f.cell)}</span></li>`).join('')}</ul>`;
    h += `<div class="review" style="margin-top:14px"><h4>Reviu evaluator</h4>
      <span class="seg rv-status">${STATUS.map(([v, lab]) => `<button type="button" data-v="${v}" class="${(rv.status || '') === v ? 'on ' + v : ''}">${lab}</button>`).join('')}</span>
      <textarea class="rv-note" placeholder="Catatan…">${esc(rv.note || '')}</textarea></div>`;
    const d = document.createElement('aside');
    d.className = 'drawer'; d.id = 'drawer'; d.innerHTML = h;
    document.body.appendChild(d);
    $('#dClose').onclick = closeDrawer;
    $$('.rv-status button', d).forEach(b => b.onclick = () => {
      setRev(key, { status: b.dataset.v });
      $$('.rv-status button', d).forEach(x => { x.className = x === b ? 'on ' + b.dataset.v : ''; });
      const dot = document.querySelector(`tr[data-row="${r.row}"] .st-dot`); if (dot) dot.className = 'st-dot ' + b.dataset.v;
    });
    let t; $('.rv-note', d).oninput = e => { clearTimeout(t); t = setTimeout(() => setRev(key, { note: e.target.value }), 300); };
  }

  // ---------- KKLEAD_SPIP ----------
  function renderLead(el, m) {
    const stg = STAGES.filter(s => m.stageCols[s] != null);
    let h = `<h1>${esc(m.name)}</h1><p class="muted">Nilai yang dihasilkan rumus tidak tersimpan di berkas ini, jadi ditampilkan sebagai "rumus". Nilai yang diketik manual ditandai. Bandingkan dengan hasil hitung ulang di Ringkasan.</p>`;
    if (m.flags.length) h += `<ul class="flags">${m.flags.slice(0, 6).map(f => `<li class="${f.level}">${esc(f.msg)} <span class="cell">${esc(f.cell)}</span></li>`).join('')}${m.flags.length > 6 ? `<li class="${m.flags[0].level}">…dan ${m.flags.length - 6} sel lain dengan pola yang sama.</li>` : ''}</ul>`;
    const combo = state.combo;
    h += `<table class="grid" style="margin-top:12px"><thead><tr><th>Komponen / subunsur</th>${stg.map(s => `<th class="${s}">Skor ${STAGE_LABEL[s]}</th>`).join('')}${combo && combo.rows.length ? stg.map(s => `<th class="${s}">Hitung ulang ${STAGE_LABEL[s]}</th>`).join('') : ''}</tr></thead><tbody>`;
    for (const r of m.rows) {
      if (r.section) { h += `<tr class="sec"><td colspan="${1 + stg.length * 2}">${esc(r.label)}</td></tr>`; continue; }
      const cr = combo && combo.rows.find(x => x.kode && x.kode === r.kode);
      h += `<tr><td>${esc(r.label)}</td>${stg.map(s => {
        const v = r.vals[s];
        if (!v) return '<td></td>';
        return `<td class="num">${v.formula ? `<span class="muted" title="=${esc(v.f)}">rumus</span>` : v.value != null ? `${fmt(v.value, 3)} <span class="b warn">manual</span>` : ''}</td>`;
      }).join('')}${combo && combo.rows.length ? stg.map(s => `<td class="num">${cr ? fmt(cr.avg[s]) : ''}</td>`).join('') : ''}</tr>`;
    }
    h += `</tbody></table>`;
    el.innerHTML = h;
  }

  function renderInfo(el, m) {
    let h = `<h1>${esc(m.name)}</h1>`;
    if (m.error) h += `<div class="flags"><li class="error">Sheet tidak dapat dianalisis otomatis; ditampilkan sebagai teks. (${esc(m.error.split('\n')[0])})</li></div>`;
    h += m.lines.length ? `<div class="infolist">${m.lines.slice(0, 3000).map(l => `<div><span class="cell">${l.row + 1}</span>${esc(l.cells.join('  ·  '))}</div>`).join('')}</div>` : `<div class="empty-note">Sheet kosong (atau hanya berisi rumus tanpa nilai tersimpan).</div>`;
    el.innerHTML = h;
  }

  // ---------- ekspor ----------
  function exportXlsx() {
    const wb = XLSX.utils.book_new();
    const statusLabel = Object.fromEntries(STATUS);
    // 1. reviu parameter S&P
    const rows1 = [];
    for (const m of state.models.filter(x => x.type === 'sp')) for (const p of m.params) for (const u of m.activeUnits) {
      const pu = p.units.find(x => x.unit.id === u.id);
      const r = rev(spKey(p, u.id));
      const fl = p.flags.filter(f => !f.unit || f.unit === u.id);
      const row = { Sheet: m.name, 'Kode subunsur': p.sub.kode, Subunsur: p.sub.nama, 'No': p.no, Parameter: p.uraian, 'Satker/asesor': u.label };
      for (const s of STAGES) row['Grade ' + STAGE_LABEL[s]] = pu && pu.stages[s] ? pu.stages[s].grade : '';
      for (const s of STAGES) row['Kesimpulan ' + STAGE_LABEL[s] + ' (hitung ulang)'] = p.kes[s] || '';
      Object.assign(row, {
        'Grade menurut evaluator': r.grade || '', 'Status reviu': statusLabel[r.status || ''], 'Catatan evaluator': r.note || '',
        'Temuan perlu perhatian': fl.filter(f => f.level !== 'info').length, 'Ringkasan temuan': fl.filter(f => f.level !== 'info').map(f => '• ' + f.msg).join('\n'),
        'Sel': `${m.name}!${P.addr(p.row, 3)}`,
      });
      rows1.push(row);
    }
    if (rows1.length) XLSX.utils.book_append_sheet(wb, withWidths(XLSX.utils.json_to_sheet(rows1), [10, 8, 30, 5, 50, 16, 8, 8, 8, 10, 10, 10, 10, 14, 50, 10, 60, 14]), 'Reviu Struktur Proses');
    // 2. reviu baris tabel
    const rows2 = [];
    for (const m of state.models.filter(x => x.type === 'table')) for (const r of m.rows) {
      const rv = state.reviews[rowKey(m, r)];
      if (!rv) continue;
      const ident = m.cols.map((c, i) => c.stage ? '' : r.vals[i]).filter(Boolean).join(' | ');
      rows2.push({ Sheet: m.name, Baris: r.row + 1, Uraian: ident, 'Status reviu': statusLabel[rv.status || ''], 'Catatan evaluator': rv.note || '' });
    }
    if (rows2.length) XLSX.utils.book_append_sheet(wb, withWidths(XLSX.utils.json_to_sheet(rows2), [14, 7, 80, 14, 60]), 'Reviu Tabel KK');
    // 3. semua temuan
    const rows3 = allFindings().map(x => {
      let key = '';
      if (x.p) key = spKey(x.p, x.f.unit || (x.m.activeUnits[0] || {}).id);
      else if (x.r) key = rowKey(x.m, x.r);
      const r = key ? rev(key) : {};
      return { Tingkat: LEVEL_LABEL[x.f.level], Sheet: x.m.name, Lokasi: x.where, Parameter: x.p ? x.p.uraian : '', Jenis: x.f.code, Keterangan: x.f.msg, Sel: x.f.cell, 'Status reviu': statusLabel[r.status || ''], 'Catatan evaluator': r.note || '' };
    });
    XLSX.utils.book_append_sheet(wb, withWidths(XLSX.utils.json_to_sheet(rows3.length ? rows3 : [{ Keterangan: 'Tidak ada temuan' }]), [14, 12, 14, 40, 20, 80, 10, 12, 40]), 'Temuan Otomatis');
    // 4. rekap skor
    if (state.combo && state.combo.rows.length) {
      const rows4 = state.combo.rows.map(r => {
        const o = { Kode: r.kode, Subunsur: r.nama };
        for (const sh of state.combo.sheets) for (const s of STAGES) { const v = r.per[sh] && r.per[sh][s]; o[`${sh} ${STAGE_LABEL[s]}`] = v ? +v.value.toFixed(3) : ''; }
        for (const s of STAGES) o['Gabungan ' + STAGE_LABEL[s]] = r.avg[s] != null ? +r.avg[s].toFixed(3) : '';
        return o;
      });
      XLSX.utils.book_append_sheet(wb, withWidths(XLSX.utils.json_to_sheet(rows4), [6, 40]), 'Rekap Skor S&P');
    }
    const base = state.fileName.replace(/\.[^.]+$/, '');
    const d = new Date(); const ds = `${d.getFullYear()}${String(d.getMonth() + 1).padStart(2, '0')}${String(d.getDate()).padStart(2, '0')}`;
    XLSX.writeFile(wb, `Reviu_${base}_${ds}.xlsx`);
  }
  function withWidths(ws, w) { ws['!cols'] = w.map(x => ({ wch: x })); return ws; }

  function saveSession() {
    const blob = new Blob([JSON.stringify({ app: 'spip-evaluator', version: 1, file: state.fileName, saved: new Date().toISOString(), reviews: state.reviews }, null, 1)], { type: 'application/json' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = `Sesi_reviu_${state.fileName.replace(/\.[^.]+$/, '')}.json`;
    a.click();
    setTimeout(() => URL.revokeObjectURL(a.href), 1000);
  }
  async function loadSession(file) {
    try {
      const j = JSON.parse(await file.text());
      if (!j.reviews) throw new Error('bukan berkas sesi');
      if (j.file && j.file !== state.fileName && !confirm(`Sesi ini dibuat untuk "${j.file}", berkas terbuka "${state.fileName}". Tetap gabungkan?`)) return;
      Object.assign(state.reviews, j.reviews);
      saveReviews();
      toast(`${Object.keys(j.reviews).length} catatan dimuat.`);
      go(state.view, { keepScroll: true });
    } catch (e) { alert('Berkas sesi tidak valid: ' + e.message); }
  }

  // ---------- event ----------
  const fileInput = $('#fileInput');
  $('#btnOpen').onclick = () => fileInput.click();
  $('#btnOpen2').onclick = () => fileInput.click();
  fileInput.onchange = () => { openFile(fileInput.files[0]); fileInput.value = ''; };
  $('#btnSheets').onclick = showSheetDialog;
  $('#btnExport').onclick = exportXlsx;
  $('#btnSaveSession').onclick = () => { $('#menuMore').open = false; saveSession(); };
  $('#btnLoadSession').onclick = () => { $('#menuMore').open = false; $('#sessionInput').click(); };
  $('#sessionInput').onchange = e => { if (e.target.files[0]) loadSession(e.target.files[0]); e.target.value = ''; };
  $('#btnClearSession').onclick = () => {
    $('#menuMore').open = false;
    if (!confirm('Hapus semua status & catatan reviu untuk berkas ini dari peramban?')) return;
    state.reviews = {}; saveReviews(); go(state.view, { keepScroll: true });
  };
  $('#sheetDialog').addEventListener('close', () => {
    if ($('#sheetDialog').returnValue !== 'ok') return;
    loadSelected($$('#sheetList input:checked').map(x => x.value));
  });
  const drop = $('#drop');
  ['dragenter', 'dragover'].forEach(ev => document.addEventListener(ev, e => { e.preventDefault(); drop.classList.add('over'); }));
  ['dragleave', 'drop'].forEach(ev => document.addEventListener(ev, e => { e.preventDefault(); if (ev === 'drop' || e.target === document.documentElement) drop.classList.remove('over'); }));
  document.addEventListener('drop', e => { const f = e.dataTransfer.files[0]; if (f) openFile(f); });
  document.addEventListener('keydown', e => { if (e.key === 'Escape') closeDrawer(); });

  // untuk pengujian/otomasi: muat dari ArrayBuffer
  window.SpipApp = {
    loadBuffer(buf, name, sheets) {
      state.buf = buf; state.fileName = name;
      const wb = XLSX.read(buf, { bookSheets: true });
      state.sheetNames = wb.SheetNames; state.hidden = {};
      $('#fileLabel').textContent = name;
      return loadSelected(sheets || wb.SheetNames);
    },
    state,
  };
})();
