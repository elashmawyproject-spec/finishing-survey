/* Finishing survey app — data stays on this device (localStorage) and is exported as an Excel BOQ workbook. */
(function () {
  'use strict';
  const { num, txt, quantities, exportXlsx, buildRows } = globalThis.BOQ;
  const LISTS = globalThis.BOQ_LAYOUT.lists;
  const SEED = globalThis.BOQ_SEED;
  const KEY = 'boq-survey-v2', OLD_KEY = 'boq-survey-v1';
  const LEVELS = ['بدروم', 'أرضي', 'أول', 'ثاني', 'ملحق علوي', 'سطح', 'خارجي'];
  const app = document.getElementById('app');
  const titleEl = document.getElementById('title');
  const backEl = document.getElementById('back');

  // ------------------------------------------------------------ state (several projects, all kept on this device)
  const newId = () => Date.now().toString(36) + Math.random().toString(36).slice(2, 7);
  let DB = null;
  let storageOk = true;
  try { DB = JSON.parse(localStorage.getItem(KEY) || 'null'); } catch (e) { DB = null; }
  if (!DB || !Array.isArray(DB.projects)) {
    DB = { v: 2, projects: [], cur: null };
    try { const old = JSON.parse(localStorage.getItem(OLD_KEY) || 'null'); if (old && Array.isArray(old.spaces)) { old.id = newId(); DB.projects.push(old); } } catch (e) { /* no older data */ }
  }
  let P = null;
  function setCur(id) {
    P = DB.projects.find(p => p.id === id) || null;
    DB.cur = P ? P.id : null;
    if (P) { P.custom = P.custom || {}; P.spaces = P.spaces || []; }
  }
  setCur(DB.cur);
  let saveTimer = null;
  function save(now) {
    clearTimeout(saveTimer);
    const run = () => {
      try { localStorage.setItem(KEY, JSON.stringify(DB)); storageOk = true; }
      catch (e) { if (storageOk) toast('تعذر الحفظ على الجهاز — صدّر نسخة احتياطية الآن'); storageOk = false; }
    };
    if (now) run(); else saveTimer = setTimeout(run, 400);
  }
  window.addEventListener('pagehide', () => save(true));

  // ------------------------------------------------------------ helpers
  const esc = s => String(s === null || s === undefined ? '' : s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const f2 = n => (Math.round(n * 100) / 100).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  function toast(msg) {
    const t = document.getElementById('toast');
    t.textContent = msg; t.hidden = false;
    clearTimeout(toast.t); toast.t = setTimeout(() => { t.hidden = true; }, 3200);
  }
  function getPath(obj, path) { return path.split('.').reduce((o, k) => (o === undefined || o === null) ? undefined : o[k], obj); }
  function setPath(obj, path, val) {
    const ks = path.split('.');
    let o = obj;
    for (let i = 0; i < ks.length - 1; i++) { if (o[ks[i]] === undefined || o[ks[i]] === null) o[ks[i]] = /^\d+$/.test(ks[i + 1]) ? [] : {}; o = o[ks[i]]; }
    o[ks[ks.length - 1]] = val;
  }
  function matItems(list) {
    const base = LISTS[list].items;
    const extra = ((P.custom[list]) || []).map(c => c.name).filter(n => !base.includes(n));
    return base.concat(extra);
  }
  function hasData(sp) {
    const q = quantities(sp);
    return q.floor > 0 || q.perim > 0 || (sp.kind === 'stair' && num(sp.stair && sp.stair.steps) > 0);
  }
  const statusOf = sp => sp.done ? ['تم', 'ok'] : hasData(sp) ? ['جارٍ', 'mid'] : ['لم يبدأ', 'none'];

  // ------------------------------------------------------------ field builders
  const inp = (path, val, ph, cls) => `<input type="text" data-p="${path}" value="${esc(val)}" placeholder="${esc(ph || '')}" class="${cls || ''}" autocomplete="off">`;
  const numInp = (path, val, ph) => `<input type="text" inputmode="decimal" data-p="${path}" value="${esc(val)}" placeholder="${esc(ph || '')}" class="n" autocomplete="off">`;
  function sel(path, val, list, opts) {
    opts = opts || {};
    let o = `<option value="">${esc(opts.blank || '— اختر —')}</option>`;
    (opts.pre || []).forEach(([v, t]) => { o += `<option value="${esc(v)}"${v === val ? ' selected' : ''}>${esc(t)}</option>`; });
    const items = Array.isArray(list) ? list : matItems(list);
    if (val && val !== '-' && !items.includes(val)) items.push(val);
    items.forEach(v => { o += `<option value="${esc(v)}"${v === val ? ' selected' : ''}>${esc(v)}</option>`; });
    if (!Array.isArray(list) && LISTS[list].ucol) o += '<option value="__new__">＋ خامة جديدة…</option>';
    return `<select data-p="${path}" data-list="${Array.isArray(list) ? '' : list}">${o}</select>`;
  }
  const chk = (path, on, label) => `<label class="chk"><input type="checkbox" data-p="${path}"${on ? ' checked' : ''}> ${esc(label)}</label>`;
  const fld = (label, html, cls) => `<label class="fld ${cls || ''}"><span>${esc(label)}</span>${html}</label>`;

  // ------------------------------------------------------------ views
  const progress = p => [p.spaces.filter(x => x.done).length, p.spaces.length];

  function viewProjects() {
    titleEl.textContent = 'حصر التشطيبات — المشاريع';
    backEl.hidden = true;
    let h = '<section class="card"><h2>المشاريع</h2>';
    if (!DB.projects.length) h += '<p class="muted">لا توجد مشاريع بعد. ابدأ بمشروع جديد أو استورد ملف مشروع.</p>';
    DB.projects.forEach(p => {
      const [d, t] = progress(p);
      h += `<a class="row" href="#/p" data-open="${esc(p.id)}"><span><b>${esc(p.name || 'بدون اسم')}</b><br><span class="muted">${esc(p.location || '')}</span></span><span class="muted">${d} / ${t}</span></a>`;
    });
    h += `<button class="btn primary" data-act="newProject">＋ مشروع جديد</button>
      <button class="btn" data-act="restore">استيراد مشروع من ملف</button>
      <p class="muted">كل المشاريع محفوظة على هذا الجهاز فقط. احفظ نسخة احتياطية لكل مشروع بعد كل يوم عمل.</p>
      <input type="file" id="restoreFile" accept=".json,application/json" hidden></section>`;
    app.innerHTML = h;
  }

  function viewHome() {
    titleEl.textContent = P.name || 'مشروع بدون اسم';
    backEl.hidden = false; backEl.href = '#/';
    const total = P.spaces.length, done = P.spaces.filter(s => s.done).length, mid = P.spaces.filter(s => !s.done && hasData(s)).length;
    const levels = LEVELS.filter(l => P.spaces.some(s => s.level === l)).concat([...new Set(P.spaces.map(s => s.level))].filter(l => !LEVELS.includes(l)));
    let h = `<details class="card"${total ? '' : ' open'}><summary>بيانات المشروع</summary><div class="grid2">
      ${fld('اسم المشروع', inp('@name', P.name))}${fld('المالك / العميل', inp('@owner', P.owner))}
      ${fld('الموقع', inp('@location', P.location))}${fld('القائم بالحصر', inp('@surveyor', P.surveyor))}
      ${fld('تاريخ الحصر', `<input type="date" data-p="@date" value="${esc(P.date)}">`)}</div>
      ${fld('ملاحظات عامة', `<textarea data-p="@notes" rows="2">${esc(P.notes)}</textarea>`, 'mt')}</details>
      <section class="card"><h2>الأدوار والفراغات</h2>
      <div class="prog"><div style="width:${total ? Math.round(done / total * 100) : 0}%"></div></div>
      <p class="muted">تم ${done} من ${total} فراغًا${mid ? ` · جارٍ ${mid}` : ''}</p>`;
    levels.forEach(l => {
      const sp = P.spaces.filter(s => s.level === l);
      h += `<a class="row" href="#/floor/${encodeURIComponent(l)}"><b>${esc(l)}</b><span class="muted">${sp.filter(s => s.done).length} / ${sp.length}</span></a>`;
    });
    h += `<div class="acts"><button class="btn" data-act="addSpace">＋ فراغ جديد</button><a class="btn" href="#/add">＋ عدة فراغات مرة واحدة</a></div></section>
      <section class="card"><h2>التصدير والحفظ</h2>
      <button class="btn primary" data-act="export">تصدير ملف Excel</button>
      <button class="btn" data-act="backup">حفظ نسخة احتياطية / إرسال المشروع (ملف)</button>
      <button class="btn danger" data-act="delProject">حذف هذا المشروع من الجهاز</button>
      <p class="muted">البيانات محفوظة على هذا الجهاز فقط. صدّر Excel أو نسخة احتياطية بعد كل يوم عمل.</p></section>`;
    app.innerHTML = h;
  }

  function viewAdd() {
    titleEl.textContent = 'إضافة عدة فراغات';
    backEl.hidden = false; backEl.href = '#/p';
    app.innerHTML = `<section class="card"><h2>إضافة عدة فراغات</h2>
      <div class="grid2">${fld('الدور', `<select id="bulkLevel">${LEVELS.map(l => `<option>${esc(l)}</option>`).join('')}</select>`)}
      ${fld('عدد الحوائط المبدئي', '<input type="text" inputmode="numeric" id="bulkWalls" value="4" class="n">')}</div>
      ${fld('سطر لكل فراغ: الكود ثم الاسم', '<textarea id="bulkText" rows="10" class="ltr" placeholder="GF-01 ENTRANCE&#10;GF-02 MAIN SALON&#10;GF-ST-01 STAIR"></textarea>', 'mt')}
      <p class="muted">الفراغ الذي يبدأ اسمه بكلمة «سلم» أو «درج» أو STAIR يُسجَّل كسلم.</p>
      <button class="btn primary" data-act="bulkAdd">إضافة</button></section>`;
  }

  function viewFloor(level) {
    titleEl.textContent = 'الدور: ' + level;
    backEl.hidden = false; backEl.href = '#/p';
    let h = '<section class="card">';
    P.spaces.forEach((s, i) => {
      if (s.level !== level) return;
      const [t, c] = statusOf(s);
      h += `<a class="row" href="#/space/${i}"><span><b dir="ltr">${esc(s.code)}</b> <span dir="ltr">${esc(s.name)}</span></span><span class="chip ${c}">${t}</span></a>`;
    });
    h += `<button class="btn ghost" data-act="addSpace" data-level="${esc(level)}">＋ فراغ جديد في هذا الدور</button></section>`;
    app.innerHTML = h;
  }

  function partRows(base, parts, withMat) {
    let h = '';
    parts.forEach((f, i) => {
      h += `<div class="item"><div class="item-h"><b>جزء ${i + 1}</b><button class="x" data-act="del" data-arr="${base}" data-i="${i}" aria-label="حذف">✕</button></div>
        <div class="grid2">${withMat ? fld('الخامة', sel(`${base}.${i}.mat`, txt(f.mat), 'FloorMat')) : ''}${fld('الوصف', inp(`${base}.${i}.desc`, f.desc, 'اختياري'))}</div>
        <div class="grid4">${fld('الطول', numInp(`${base}.${i}.len`, f.len))}${fld('العرض', numInp(`${base}.${i}.wid`, f.wid))}
        ${fld('التكرار', numInp(`${base}.${i}.cnt`, f.cnt, '1'))}${fld('أو مساحة مباشرة', numInp(`${base}.${i}.dir`, f.dir))}</div></div>`;
    });
    return h;
  }

  function wallCard(sp, w, i) {
    const base = `walls.${i}`;
    let h = `<div class="item wall"><div class="item-h"><b>ح${i + 1}</b><button class="x" data-act="del" data-arr="walls" data-i="${i}" aria-label="حذف الحائط">✕</button></div>
      <div class="grid3">${fld('الطول (م)', numInp(`${base}.len`, w.len))}
      ${fld('خامة مختلفة', sel(`${base}.mat`, txt(w.mat), 'WallMat', { blank: 'حسب الفراغ', pre: [['-', 'بدون تشطيب']] }))}
      ${fld('ارتفاع مختلف', numInp(`${base}.hgt`, w.hgt, 'حسب الفراغ'))}</div>`;
    (w.ops || []).forEach((o, k) => {
      const ob = `${base}.ops.${k}`;
      const isWin = /نافذة|شباك/.test(txt(o.typ));
      const skDef = !(isWin && sp.skirt && sp.skirt.underWin);
      h += `<div class="sub"><div class="item-h"><span>فتحة ${k + 1}</span><button class="x" data-act="del" data-arr="${base}.ops" data-i="${k}" aria-label="حذف">✕</button></div>
        <div class="grid4">${fld('النوع', sel(`${ob}.typ`, txt(o.typ), 'OpenTypes'))}${fld('العرض', numInp(`${ob}.wid`, o.wid))}
        ${fld('الارتفاع', numInp(`${ob}.hgt`, o.hgt))}${fld('العدد', numInp(`${ob}.cnt`, o.cnt, '1'))}</div>
        <div class="chks"><span class="muted">تُخصم من:</span>${chk(`${ob}.fin`, o.fin !== false, 'التشطيب')}${chk(`${ob}.pl`, o.pl !== false, 'اللياسة')}
        ${chk(`${ob}.sk`, o.sk === undefined ? skDef : o.sk, 'الوزرة')}</div></div>`;
    });
    (w.feats || []).forEach((f, k) => {
      const fb = `${base}.feats.${k}`;
      const lm = f.unit === 'lm';
      h += `<div class="sub feat"><div class="item-h"><span>جزء بخامة أخرى ${k + 1}</span><button class="x" data-act="del" data-arr="${base}.feats" data-i="${k}" aria-label="حذف">✕</button></div>
        <div class="grid2">${fld('القياس', `<select data-p="${fb}.unit" data-rerender="1"><option value="m2"${lm ? '' : ' selected'}>م² (مسطح)</option><option value="lm"${lm ? ' selected' : ''}>م.ط (طولي)</option></select>`)}
        ${fld('الخامة', sel(`${fb}.mat`, txt(f.mat), lm ? 'LinMat' : 'WallMat'))}</div>
        <div class="grid3">${fld('الطول', numInp(`${fb}.len`, f.len))}${lm ? '' : fld('الارتفاع', numInp(`${fb}.hgt`, f.hgt))}${fld('العدد', numInp(`${fb}.cnt`, f.cnt, '1'))}</div>
        ${lm ? '' : `<div class="chks">${chk(`${fb}.ded`, f.ded !== false, 'يُخصم من تشطيب الحائط الأساسي (لا يوجد تحته)')}</div>`}</div>`;
    });
    h += `<div class="acts"><button class="btn sm" data-act="addOp" data-i="${i}">＋ فتحة</button><button class="btn sm" data-act="addFeat" data-i="${i}">＋ جزء بخامة أخرى</button></div></div>`;
    return h;
  }

  function summaryHtml(sp) {
    const q = quantities(sp);
    const line = (a, b) => `<div><span>${a}</span><b dir="ltr">${b}</b></div>`;
    let h = line('الأرضية (م²)', f2(q.floor));
    h += line('محيط الحوائط (م)', f2(q.perim));
    h += line('تشطيب الحوائط: إجمالي − خصم = صافي', `${f2(q.wallGross)} − ${f2(q.wallDed)} = ${f2(q.wallGross - q.wallDed)}`);
    if (q.plGross > 0) h += line('اللياسة: إجمالي − خصم = صافي', `${f2(q.plGross)} − ${f2(q.plDed)} = ${f2(q.plGross - q.plDed)}`);
    if (sp.skirt && sp.skirt.on !== false && q.perim > 0) h += line('الوزرة (م.ط)', `${f2(q.perim)} − ${f2(q.skDed)} = ${f2(q.perim - q.skDed)}`);
    h += line('عدد الفتحات', String(q.openings));
    return h;
  }

  function viewSpace(i) {
    const sp = P.spaces[i];
    if (!sp) { location.hash = '#/p'; return; }
    titleEl.textContent = sp.code + ' — ' + sp.name;
    backEl.hidden = false; backEl.href = '#/floor/' + encodeURIComponent(sp.level);
    sp.floors = sp.floors || []; sp.walls = sp.walls || []; sp.ceil = sp.ceil || { same: true }; sp.wall = sp.wall || {};
    sp.plaster = sp.plaster || { on: true }; sp.skirt = sp.skirt || { on: true }; sp.corners = sp.corners || {}; sp.stair = sp.stair || {};
    const stair = sp.kind === 'stair';
    let h = `<details class="card"><summary>بيانات الفراغ</summary><div class="grid2">
      ${fld('الكود', inp('code', sp.code, '', 'ltr'))}${fld('الاسم', inp('name', sp.name, '', 'ltr'))}
      ${fld('الدور', sel('level', sp.level, LEVELS))}${fld('النوع', `<select data-p="kind" data-rerender="1"><option value="room"${stair ? '' : ' selected'}>غرفة / فراغ</option><option value="stair"${stair ? ' selected' : ''}>سلم</option></select>`)}
      </div><button class="btn danger sm" data-act="delSpace">حذف هذا الفراغ</button></details>`;

    if (stair) {
      h += `<section class="card"><h2>الدرج</h2><div class="grid4">
        ${fld('عدد الدرجات', numInp('stair.steps', sp.stair.steps))}${fld('عرض الدرجة', numInp('stair.width', sp.stair.width))}
        ${fld('النائمة', numInp('stair.tread', sp.stair.tread))}${fld('القائمة', numInp('stair.riser', sp.stair.riser))}</div>
        <div class="grid2">${fld('خامة الدرج', inp('stair.mat', sp.stair.mat, 'رخام / جرانيت …'))}${fld('وزرة الدرج (م.ط)', numInp('stair.skirtLen', sp.stair.skirtLen))}</div>
        <p class="muted">النوائم والقوائم تُحسب: عرض الدرجة × عدد الدرجات (م.ط). البسطات تُدخل في الأرضية.</p></section>`;
    }

    h += `<section class="card"><h2>${stair ? 'البسطات (أرضية)' : 'الأرضية'}</h2>${partRows('floors', sp.floors, true)}
      <button class="btn sm" data-act="addFloor">＋ جزء / خامة أخرى</button>
      <div class="grid2 mt">${fld('فاصل ستانلس بالأرضية (م.ط)', numInp('floorStrip', sp.floorStrip))}</div></section>`;

    h += `<section class="card"><h2>السقف</h2><div class="grid2">${fld('نوع السقف', sel('ceil.mat', txt(sp.ceil.mat), 'CeilMat', { blank: 'لا يُحصر' }))}
      ${fld('ارتفاعه من الأرض', numInp('ceil.hgt', sp.ceil.hgt))}</div>
      <div class="chks">${chk('ceil.same', sp.ceil.same !== false, 'نفس أبعاد الأرضية')}</div>
      ${sp.ceil.same === false ? partRows('ceil.parts', sp.ceil.parts || [], false) + '<button class="btn sm" data-act="addCeil">＋ جزء</button>' : ''}</section>`;

    h += `<section class="card"><h2>تشطيب الحوائط</h2>
      ${P.spaces.some((s, k) => k !== i && txt(s.wall && s.wall.mat)) && !txt(sp.wall.mat) ? '<button class="btn sm ghost" data-act="copyLast">نسخ إعدادات التشطيب من آخر فراغ</button>' : ''}
      <div class="grid2">${fld('خامة التشطيب', sel('wall.mat', txt(sp.wall.mat), 'WallMat'))}${fld('ارتفاع التشطيب', numInp('wall.hgt', sp.wall.hgt))}</div>
      <div class="chks">${chk('plaster.on', sp.plaster.on !== false, 'توجد لياسة')}</div>
      <div class="grid2">${fld('نوع اللياسة', sel('plaster.mat', txt(sp.plaster.mat), 'PlMat'))}${fld('ارتفاع اللياسة', numInp('plaster.hgt', sp.plaster.hgt))}</div>
      <div class="chks">${chk('skirt.on', sp.skirt.on !== false, 'توجد وزرة')}${chk('skirt.underWin', !!sp.skirt.underWin, 'الوزرة موجودة أسفل النوافذ')}</div>
      <div class="grid2">${fld('خامة الوزرة', sel('skirt.mat', txt(sp.skirt.mat), 'LinMat'))}</div></section>`;

    h += `<section class="card"><h2>الحوائط والفتحات</h2>${sp.walls.map((w, k) => wallCard(sp, w, k)).join('')}
      <button class="btn" data-act="addWall">＋ حائط</button>
      <div class="grid2 mt">${fld('زوايا ستانلس: العدد', numInp('corners.cnt', sp.corners.cnt))}${fld('ارتفاع الزاوية', numInp('corners.hgt', sp.corners.hgt))}</div></section>`;

    h += `<section class="card"><h2>ملاحظات</h2><textarea data-p="note" rows="3">${esc(sp.note)}</textarea></section>
      <section class="card sum"><h2>ملخص الفراغ</h2><div id="sum">${summaryHtml(sp)}</div></section>
      <section class="card"><label class="done"><input type="checkbox" data-p="done"${sp.done ? ' checked' : ''}> تم رفع هذا الفراغ</label>
      <div class="acts">${i > 0 ? `<a class="btn" href="#/space/${i - 1}">→ السابق</a>` : ''}${i < P.spaces.length - 1 ? `<a class="btn primary" href="#/space/${i + 1}">التالي ←</a>` : ''}</div></section>`;
    app.innerHTML = h;
  }

  // ------------------------------------------------------------ routing
  let cur = { view: 'home' };
  function route(keepScroll) {
    const y = window.scrollY;
    const h = location.hash;
    const m = h.match(/^#\/(floor|space)\/(.+)$/);
    if (!P || h === '' || h === '#/' || h === '#') { cur = { view: 'projects' }; viewProjects(); }
    else if (m && m[1] === 'floor') { cur = { view: 'floor', level: decodeURIComponent(m[2]) }; viewFloor(cur.level); }
    else if (m && m[1] === 'space') { cur = { view: 'space', i: Number(m[2]) }; viewSpace(cur.i); }
    else if (h === '#/add') { cur = { view: 'add' }; viewAdd(); }
    else { cur = { view: 'home' }; viewHome(); }
    window.scrollTo(0, keepScroll ? y : 0);
  }
  window.addEventListener('hashchange', () => route(false));
  const curSpace = () => cur.view === 'space' ? P.spaces[cur.i] : null;

  // ------------------------------------------------------------ input handling
  function onField(el) {
    const path = el.dataset.p;
    if (!path) return;
    let val = el.type === 'checkbox' ? el.checked : el.value;
    const target = path[0] === '@' ? P : curSpace();
    if (!target) return;
    const p = path.replace(/^@/, '');
    if (el.tagName === 'SELECT' && val === '__new__') {
      const list = el.dataset.list;
      const name = (prompt('اسم الخامة الجديدة:') || '').trim().replace(/[*?]/g, '×');
      if (!name) { el.value = txt(getPath(target, p)); return; }
      const unit = list === 'LinMat' ? 'م.ط' : 'م²';
      P.custom[list] = P.custom[list] || [];
      if (!matItems(list).includes(name)) P.custom[list].push({ name, unit });
      val = name;
      setPath(target, p, val); save(); route(true);
      return;
    }
    setPath(target, p, val);
    save();
    if (el.dataset.rerender || (el.type === 'checkbox' && /^(ceil\.same|skirt\.underWin)$/.test(p))) { route(true); return; }
    const s = document.getElementById('sum');
    if (s && target !== P) s.innerHTML = summaryHtml(target);
    if (target !== P && (p === 'code' || p === 'name')) titleEl.textContent = target.code + ' — ' + target.name;
    if (target === P && p === 'name') titleEl.textContent = P.name || 'مشروع بدون اسم';
  }
  app.addEventListener('input', e => { if (e.target.matches('input[type=text],input[type=date],textarea')) onField(e.target); });
  app.addEventListener('click', e => { const a = e.target.closest('[data-open]'); if (a) { setCur(a.dataset.open); save(); } }, true);
  app.addEventListener('change', e => { if (e.target.matches('select,input[type=checkbox]')) onField(e.target); });

  // ------------------------------------------------------------ actions
  function download(bytes, name, type) {
    const blob = new Blob([bytes], { type });
    const file = new File([blob], name, { type });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob); a.download = name;
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(a.href), 60000);
    return file;
  }
  const stamp = () => new Date().toISOString().slice(0, 10);
  const safeName = () => (txt(P.name) || 'Project').replace(/[\\/:*?"<>|]+/g, '_').replace(/\s+/g, '_');

  const ACT = {
    addFloor: sp => sp.floors.push({ mat: txt((sp.floors[sp.floors.length - 1] || {}).mat) }),
    addCeil: sp => { sp.ceil.parts = sp.ceil.parts || []; sp.ceil.parts.push({}); },
    addWall: sp => sp.walls.push({ len: '', ops: [], feats: [] }),
    addOp: (sp, d) => { const w = sp.walls[d.i]; w.ops = w.ops || []; w.ops.push({ typ: 'باب', wid: '', hgt: '', cnt: '' }); },
    addFeat: (sp, d) => { const w = sp.walls[d.i]; w.feats = w.feats || []; w.feats.push({ unit: 'm2', mat: '', len: '', hgt: '', cnt: '', ded: true }); },
    del: (sp, d) => {
      const arr = getPath(sp, d.arr);
      const it = arr[d.i];
      const filled = it && Object.values(it).some(v => (typeof v === 'string' && v.trim()) || (Array.isArray(v) && v.length));
      if (filled && !confirm('حذف هذا العنصر؟')) return false;
      arr.splice(d.i, 1);
    },
    copyLast: sp => {
      const src = [...P.spaces].reverse().find(s => s !== sp && txt(s.wall && s.wall.mat) && s.done) || [...P.spaces].reverse().find(s => s !== sp && txt(s.wall && s.wall.mat));
      if (!src) return false;
      sp.wall = { ...src.wall }; sp.plaster = { ...src.plaster }; sp.skirt = { ...src.skirt };
      if (!txt(sp.ceil.mat)) sp.ceil = { ...sp.ceil, mat: src.ceil.mat, hgt: src.ceil.hgt };
      toast('تم نسخ الإعدادات من ' + src.code);
    },
  };

  async function doExport(btn) {
    const rows = buildRows(P);
    if (!rows.FL.length && !rows.WL.length && !rows.LN.length) { toast('لا توجد قياسات للتصدير بعد'); return; }
    btn.disabled = true; btn.textContent = 'جارٍ إنشاء الملف…';
    try {
      const r = await exportXlsx(P);
      const name = `${safeName()}_BOQ_${stamp()}.xlsx`;
      const file = download(r.bytes, name, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
      toast(`تم إنشاء ${name}`);
      if (navigator.canShare && navigator.canShare({ files: [file] }) && confirm('تم تنزيل الملف. هل تريد مشاركته الآن؟')) {
        try { await navigator.share({ files: [file], title: name }); } catch (e) { /* cancelled */ }
      }
    } catch (e) { alert('تعذر التصدير: ' + e.message); }
    btn.disabled = false; btn.textContent = 'تصدير ملف Excel';
  }

  document.addEventListener('click', e => {
    const b = e.target.closest('[data-act]');
    if (!b) return;
    const act = b.dataset.act, d = { ...b.dataset, i: Number(b.dataset.i) };
    if (act === 'export') return void doExport(b);
    if (act === 'backup') { const { id, ...copy } = P; download(new TextEncoder().encode(JSON.stringify(copy, null, 1)), `${safeName()}_backup_${stamp()}.json`, 'application/json'); return; }
    if (act === 'restore') { document.getElementById('restoreFile').click(); return; }
    if (act === 'newProject') {
      const name = (prompt('اسم المشروع الجديد:') || '').trim();
      if (!name) return;
      const p = SEED.empty(); p.id = newId(); p.name = name; p.date = stamp();
      DB.projects.push(p); setCur(p.id); save(true); location.hash = '#/p'; return;
    }
    if (act === 'delProject') {
      if (!confirm(`سيتم حذف المشروع «${P.name}» وكل قياساته من هذا الجهاز. هل حفظت نسخة احتياطية؟`)) return;
      if (!confirm('تأكيد أخير: حذف المشروع نهائيًا من هذا الجهاز؟')) return;
      DB.projects = DB.projects.filter(p => p !== P); setCur(null); save(true); location.hash = '#/'; route(false); return;
    }
    if (act === 'bulkAdd') {
      const level = document.getElementById('bulkLevel').value;
      const nw = Math.max(0, Math.min(20, Math.round(num(document.getElementById('bulkWalls').value) ?? 4)));
      let added = 0, skipped = 0;
      document.getElementById('bulkText').value.split(/\r?\n/).forEach(line => {
        const m = line.trim().match(/^([^\s|,،]+)[\s|,،]*(.*)$/);
        if (!m) return;
        const code = m[1], name = m[2].trim();
        if (P.spaces.some(s => s.code === code)) { skipped++; return; }
        const stair = /^(درج|سلم|stair)/i.test(name);
        let at = -1;
        P.spaces.forEach((s, k) => { if (s.level === level) at = k; });
        P.spaces.splice(at < 0 ? P.spaces.length : at + 1, 0, SEED.blankSpace(code, name, level, stair ? 'stair' : 'room', stair ? 0 : nw));
        added++;
      });
      save(true); toast(`أضيف ${added} فراغًا${skipped ? ` — تم تخطي ${skipped} لتكرار الكود` : ''}`);
      if (added) location.hash = '#/floor/' + encodeURIComponent(level);
      return;
    }
    if (act === 'addSpace') {
      const code = (prompt('كود الفراغ (مثال: GF-16):') || '').trim();
      if (!code) return;
      if (P.spaces.some(s => s.code === code)) { toast('هذا الكود مستخدم من قبل'); return; }
      const name = (prompt('اسم الفراغ:') || '').trim();
      const level = b.dataset.level || (prompt('الدور (' + LEVELS.join(' / ') + '):', 'أرضي') || 'أرضي').trim();
      const isStair = /^(درج|سلم|stair)/i.test(name);
      const sp = SEED.blankSpace(code, name, level, isStair ? 'stair' : 'room', isStair ? 0 : 4);
      let at = -1;
      P.spaces.forEach((s, k) => { if (s.level === level) at = k; });
      P.spaces.splice(at < 0 ? P.spaces.length : at + 1, 0, sp);
      save(true); location.hash = '#/space/' + P.spaces.indexOf(sp); return;
    }
    if (act === 'delSpace') {
      const sp = curSpace();
      if (!sp || !confirm(`حذف الفراغ ${sp.code} وكل قياساته؟`)) return;
      P.spaces.splice(cur.i, 1); save(true); location.hash = '#/floor/' + encodeURIComponent(sp.level); return;
    }
    if (!P) return;
    const sp = curSpace();
    if (sp && ACT[act]) { if (ACT[act](sp, d) !== false) { save(); route(true); } }
  });

  document.addEventListener('change', e => {
    if (e.target.id !== 'restoreFile' || !e.target.files[0]) return;
    const fr = new FileReader();
    fr.onload = () => {
      try {
        const data = JSON.parse(fr.result);
        if (!data || !Array.isArray(data.spaces)) throw new Error('ملف غير صالح');
        if (!confirm(`استيراد «${data.name || 'بدون اسم'}» (${data.spaces.length} فراغًا) كمشروع جديد؟`)) return;
        data.id = newId(); if (DB.projects.some(p => p.name === data.name)) data.name = (data.name || 'مشروع') + ' (مستورد)';
        DB.projects.push(data); setCur(data.id); save(true); location.hash = '#/p'; route(false); toast('تم استيراد المشروع');
      } catch (err) { alert('تعذر قراءة الملف: ' + err.message); }
    };
    fr.readAsText(e.target.files[0]);
    e.target.value = '';
  });

  // ------------------------------------------------------------ install banner + offline
  let deferred = null;
  const bar = document.getElementById('install');
  const dismissed = () => { try { return localStorage.getItem(KEY + '-install') === 'later'; } catch (e) { return false; } };
  window.addEventListener('beforeinstallprompt', e => { e.preventDefault(); deferred = e; if (!dismissed()) bar.hidden = false; });
  document.getElementById('installBtn').addEventListener('click', async () => { bar.hidden = true; if (deferred) { deferred.prompt(); deferred = null; } });
  document.getElementById('installLater').addEventListener('click', () => { bar.hidden = true; try { localStorage.setItem(KEY + '-install', 'later'); } catch (e) { /* ignore */ } });
  const ios = /iphone|ipad|ipod/i.test(navigator.userAgent) && !window.navigator.standalone;
  if (ios && !dismissed()) { bar.hidden = false; document.getElementById('installMsg').textContent = 'للتثبيت على الآيفون: زر المشاركة ثم «إضافة إلى الشاشة الرئيسية»'; document.getElementById('installBtn').hidden = true; }
  if ('serviceWorker' in navigator && location.protocol !== 'file:') navigator.serviceWorker.register('sw.js').catch(() => {});

  route(false);
})();
