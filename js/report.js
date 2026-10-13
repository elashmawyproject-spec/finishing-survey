/* On-screen reports (per floor, stairs). Quantities are computed from the same rows the Excel export writes,
   with the same rules as the workbook formulas, so the report and the exported file agree. */
(function () {
  'use strict';
  const { buildRows, materials, reviewCount, txt, num } = globalThis.BOQ;
  const YES = 'نعم';
  const PRIMARY = { FL: 'م²', WL: 'م²', CL: 'م²', PL: 'م²', LN: 'م.ط' };
  const LIST = { FL: 'FloorMat', WL: 'WallMat', CL: 'CeilMat', PL: 'PlMat', LN: 'LinMat' };
  const CAT = { FL: 'الأرضيات', WL: 'تشطيبات الحوائط', CL: 'الأسقف', PL: 'اللياسة والمحارة', LN: 'الوزرات والبنود الطولية', OT: 'بنود أخرى' };
  const ORDER = ['FL', 'WL', 'CL', 'PL', 'LN', 'OT'];

  const esc = s => String(s === null || s === undefined ? '' : s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const f2 = n => (Math.round((n || 0) * 100) / 100).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  const one = v => (v === null || v === undefined) ? 1 : v;
  const nz = v => (v === null || v === undefined) ? 0 : v;

  /* every takeoff record with its gross / deduction / net: [{cat, sp, mat, unit, gross, ded, net, k}] */
  function compute(pr) {
    const R = buildRows(pr), M = materials(pr);
    const unitOf = (sheet, name) => { const m = M[LIST[sheet]].find(x => x.name === name); return (m && txt(m.unit)) || PRIMARY[sheet]; };
    const rows = [];
    const add = (sheet, r, gross, ded) => {
      const unit = unitOf(sheet, r.mat);
      rows.push({ cat: unit === PRIMARY[sheet] ? sheet : 'OT', sp: r.sp, mat: r.mat, unit, gross, ded, net: gross - ded, k: r._k || '' });
    };
    const opsOn = new Map(), opsIn = new Map(), recs = new Map();
    const push = (m, k, v) => { if (!m.has(k)) m.set(k, []); m.get(k).push(v); };
    R.OP.forEach(o => { push(opsOn, o.sp + '|' + o.wall, o); push(opsIn, o.sp, o); });
    R.WL.forEach(r => recs.set(r.sp + '|' + r.wall, (recs.get(r.sp + '|' + r.wall) || 0) + 1));
    const area = o => (o.wid !== null && o.hgt !== null) ? o.wid * o.hgt * one(o.cnt) : 0;
    const flat = r => (r.dir !== null && r.dir !== undefined) ? r.dir * one(r.cnt) : nz(r.len) * nz(r.wid) * one(r.cnt);

    R.FL.forEach(r => add('FL', r, flat(r), 0));
    R.CL.forEach(r => add('CL', r, flat(r), 0));
    R.WL.forEach(r => {
      const key = r.sp + '|' + r.wall;
      const gross = r.meth ? nz(r.dir) * one(r.cnt) : nz(r.len) * nz(r.hgt) * one(r.cnt);
      const auto = (opsOn.get(key) || []).filter(o => o.dfin === YES && (recs.get(key) === 1 || o.tgt === r.mat)).reduce((a, o) => a + area(o), 0);
      add('WL', r, gross, auto + nz(r.man));
    });
    R.PL.forEach(r => add('PL', r, nz(r.len) * nz(r.hgt), (opsOn.get(r.sp + '|' + r.wall) || []).filter(o => o.dpl === YES).reduce((a, o) => a + area(o), 0)));
    R.LN.forEach(r => {
      const auto = r.autoq === YES ? (opsIn.get(r.sp) || []).filter(o => o.dsk === YES).reduce((a, o) => a + nz(o.wid) * one(o.cnt), 0) : 0;
      add('LN', r, nz(r.dir) * one(r.cnt), auto + nz(r.man));
    });
    return rows;
  }

  const total = (rows, f) => rows.filter(f).reduce((a, r) => ({ gross: a.gross + r.gross, ded: a.ded + r.ded, net: a.net + r.net }), { gross: 0, ded: 0, net: 0 });
  /* "رخام" for one material, "رخام 8.28 + دهانات 13.98" for several */
  function matText(rows, withUnit) {
    const by = new Map();
    rows.forEach(r => by.set(r.mat, (by.get(r.mat) || 0) + r.net));
    const list = [...by].filter(([, v]) => Math.abs(v) > 1e-9 || by.size === 1);
    if (list.length === 1 && !withUnit) return list[0][0];
    return list.map(([m, v]) => `${m} ${f2(v)}${withUnit ? ' ' + rows.find(r => r.mat === m).unit : ''}`).join(' + ');
  }

  function head(pr, title, sub, cells) {
    return `<h1>${esc(title)}</h1><div class="rep-sub">${esc(sub)}</div>
      <div class="rep-meta">${cells.map(([k, v]) => `<span>${esc(k)}</span><b>${esc(v)}</b>`).join('')}</div>`;
  }
  const foot = () => '<div class="rep-sign"><span>إعداد</span><span>مراجعة</span><span>اعتماد</span></div>';
  const table = (heads, body, cls) => `<div class="rep-x"><table class="${cls || ''}"><thead><tr>${heads.map(h => `<th>${h}</th>`).join('')}</tr></thead><tbody>${body}</tbody></table></div>`;
  const tr = (cells, cls) => `<tr class="${cls || ''}">${cells.map(c => `<td${typeof c === 'number' ? ' class="n"' : ''}>${typeof c === 'number' ? f2(c) : esc(c)}</td>`).join('')}</tr>`;

  function floorHtml(pr, level) {
    const spaces = (pr.spaces || []).filter(s => s.level === level && txt(s.code));
    const codes = new Set(spaces.map(s => s.code));
    const rows = compute(pr).filter(r => codes.has(r.sp));
    const review = spaces.reduce((a, s) => a + reviewCount(s), 0);
    const done = spaces.length > 0 && spaces.every(s => s.done) && review === 0;
    let h = head(pr, 'تقرير حصر كميات التشطيبات', 'الدور: ' + level, [
      ['المشروع', pr.name], ['الموقع', pr.location], ['القائم بالحصر', pr.surveyor],
      ['تاريخ الحصر', pr.date], ['عدد الفراغات', String(spaces.length)], ['حالة التقرير', done ? 'نهائي' : 'مبدئي']]);

    // 1) floor totals per item (one line per unit, units are never added together)
    const groups = [];
    ORDER.forEach(cat => [...new Set(rows.filter(r => r.cat === cat).map(r => r.unit))].forEach(unit => groups.push({ cat, unit, ...total(rows, r => r.cat === cat && r.unit === unit) })));
    h += '<h2>أولًا: إجماليات الدور حسب البند</h2>' + table(['البند', 'الوحدة', 'الإجمالي', 'الخصومات', 'الصافي'],
      groups.map(g => tr([CAT[g.cat], g.unit, g.gross, g.ded, g.net])).join('') || tr(['لا توجد كميات بعد', '', '', '', '']), 'tot');

    // 2) per item and material
    const mats = [];
    ORDER.forEach(cat => [...new Set(rows.filter(r => r.cat === cat).map(r => r.mat))].forEach(mat => {
      const mine = rows.filter(r => r.cat === cat && r.mat === mat);
      mats.push(tr([CAT[cat], mat, mine[0].unit, ...Object.values(total(mine, () => true))]));
    }));
    h += '<h2>ثانيًا: كميات الدور حسب البند والخامة</h2>' + table(['البند', 'الخامة / النوع', 'الوحدة', 'الإجمالي', 'الخصومات', 'الصافي'], mats.join(''), 'mat');

    // 3) spaces: net quantities with the material next to each quantity
    const ceil = !pr.noCeil && rows.some(r => r.cat === 'CL'), other = rows.some(r => r.cat === 'OT');
    const heads = ['كود الفراغ', 'اسم الفراغ', 'الأرضيات<br>(م²)', 'خامة الأرضية', 'الحوائط<br>(م²)', 'خامة الحوائط']
      .concat(ceil ? ['الأسقف<br>(م²)', 'نوع السقف'] : [], ['اللياسة<br>(م²)', 'بنود طولية<br>(م.ط)', 'وصف البنود الطولية'], other ? ['بنود أخرى'] : []);
    const sum = { FL: 0, WL: 0, CL: 0, PL: 0, LN: 0 };
    const body = spaces.map(s => {
      const of = cat => rows.filter(r => r.sp === s.code && r.cat === cat);
      const net = cat => { const v = of(cat).reduce((a, r) => a + r.net, 0); sum[cat] += v; return v; };
      return tr([s.code, s.name, net('FL'), matText(of('FL')), net('WL'), matText(of('WL'))]
        .concat(ceil ? [net('CL'), matText(of('CL'))] : [], [net('PL'), net('LN'), matText(of('LN'))], other ? [matText(of('OT'), true)] : []), reviewCount(s) ? 'rv' : '');
    }).join('');
    h += '<h2>ثالثًا: تفصيل فراغات الدور (الكميات الصافية مع الخامة)</h2>' + table(heads, body +
      tr(['إجمالي الدور', '', sum.FL, '', sum.WL, ''].concat(ceil ? [sum.CL, ''] : [], [sum.PL, sum.LN, ''], other ? [''] : []), 'sum'), 'det');

    h += `<p class="rep-note">${review ? `بنود تحتاج مراجعة في هذا الدور: ${review} (فراغاتها مظللة). ` : ''}الكميات محسوبة تلقائيًا من القياسات المدخلة، وهي نفسها التي تخرج في ملف Excel. الوحدات المختلفة لا تُجمع معًا.</p>`;
    return h + foot();
  }

  function stairsHtml(pr) {
    const flights = (pr.spaces || []).filter(s => s.kind === 'stair' && txt(s.code));
    const rows = compute(pr);
    let h = head(pr, 'تقرير حصر السلالم', [pr.name, pr.location].filter(Boolean).join('   —   '), []);
    if (!flights.length) return h + '<p class="rep-note">لا توجد سلالم مسجلة في هذا المشروع.</p>';
    const groups = [...new Set(flights.map(s => txt(s.stair && s.stair.group) || 'بدون رقم'))];
    const cols = [['tread', 'LN'], ['riser', 'LN'], ['stsk', 'LN'], ['', 'FL'], ['', 'WL'], ['', 'PL']];
    groups.forEach((g, gi) => {
      const mine = flights.filter(s => (txt(s.stair && s.stair.group) || 'بدون رقم') === g);
      const sum = cols.map(() => 0);
      const body = mine.map(s => {
        const v = cols.map(([k, cat], i) => { const x = rows.filter(r => r.sp === s.code && r.cat === cat && (!k || r.k === k)).reduce((a, r) => a + r.net, 0); sum[i] += x; return x; });
        return tr([s.code, `${s.name} — ${s.level}`, ...v]);
      }).join('');
      const st = mine.find(s => num(s.stair && s.stair.width) !== null) || mine[0];
      const d = st.stair || {};
      h += `<h2>${gi + 1}) السلم ${esc(g)}${txt(d.mat) ? ' — ' + esc(d.mat) : ''}</h2>` +
        table(['كود القلبة', 'القلبة', 'النوائم<br>(م.ط)', 'القوائم<br>(م.ط)', 'وزرة الدرج<br>(م.ط)', 'البسطات والأرضية<br>(م²)', 'حوائط بيت السلم<br>(م²)', 'اللياسة<br>(م²)'],
          body + tr(['إجمالي السلم', '', ...sum], 'sum'), 'det') +
        `<p class="rep-note">${[num(d.width) !== null ? `عرض الدرجة ${f2(num(d.width))} م` : '', num(d.tread) !== null ? `النائمة ${f2(num(d.tread))}` : '', num(d.riser) !== null ? `القائمة ${f2(num(d.riser))}` : '',
          mine.some(s => s.stair && s.stair.fullWalls) ? 'حوائط بيت السلم مسجلة مرة واحدة بكامل ارتفاع المبنى على أول قلبة' : ''].filter(Boolean).join(' · ')}</p>`;
    });
    h += '<p class="rep-note">النوائم والقوائم بالمتر الطولي = عرض الدرجة × عدد الدرجات. كل قلبة داخلة أيضًا في إجمالي دورها.</p>';
    return h + foot();
  }

  /* per-space totals in the layout of the workbook's «ملخص الكميات حسب الفراغ» sheet (used by the export test) */
  function bySpace(pr) {
    const rows = compute(pr), out = {};
    (pr.spaces || []).forEach(s => {
      const of = cat => total(rows, r => r.sp === s.code && r.cat === cat);
      out[s.code] = { floor: of('FL').net, wall: of('WL'), ceil: of('CL').net, plaster: of('PL'), linear: of('LN'),
        count: rows.filter(r => r.sp === s.code && r.unit === 'عدد').reduce((a, r) => a + r.net, 0) };
    });
    return out;
  }

  globalThis.BOQ_REPORT = { compute, floorHtml, stairsHtml, bySpace };
})();
