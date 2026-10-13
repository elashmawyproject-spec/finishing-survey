/* Builds the Excel workbook: fills the input cells of the embedded BOQ template, formulas stay as they are. */
(function () {
  'use strict';
  const L = globalThis.BOQ_LAYOUT;
  const YES = 'نعم', NO = 'لا';

  // ---------------------------------------------------------------- numbers
  function num(v) {
    if (v === null || v === undefined) return null;
    if (typeof v === 'number') return isFinite(v) ? v : null;
    let s = String(v).trim();
    if (!s) return null;
    s = s.replace(/[٠-٩]/g, d => '٠١٢٣٤٥٦٧٨٩'.indexOf(d)).replace(/[٫,،]/g, '.').replace(/\s/g, '');
    const n = Number(s);
    return isFinite(n) ? n : null;
  }
  const txt = v => (v === null || v === undefined) ? '' : String(v).trim();

  // ---------------------------------------------------------------- zip (deflate via the platform streams)
  const CRC = (() => {
    const t = new Uint32Array(256);
    for (let n = 0; n < 256; n++) { let c = n; for (let k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1; t[n] = c >>> 0; }
    return t;
  })();
  function crc32(u8) { let c = 0xFFFFFFFF; for (let i = 0; i < u8.length; i++) c = CRC[(c ^ u8[i]) & 0xFF] ^ (c >>> 8); return (c ^ 0xFFFFFFFF) >>> 0; }
  async function pipe(u8, stream) { return new Uint8Array(await new Response(new Blob([u8]).stream().pipeThrough(stream)).arrayBuffer()); }
  const inflate = u8 => pipe(u8, new DecompressionStream('deflate-raw'));
  const deflate = u8 => pipe(u8, new CompressionStream('deflate-raw'));

  async function unzip(u8) {
    const dv = new DataView(u8.buffer, u8.byteOffset, u8.byteLength);
    let e = u8.length - 22;
    while (e >= 0 && dv.getUint32(e, true) !== 0x06054b50) e--;
    if (e < 0) throw new Error('ملف القالب غير صالح');
    const count = dv.getUint16(e + 10, true);
    let p = dv.getUint32(e + 16, true);
    const dec = new TextDecoder();
    const files = new Map();
    for (let i = 0; i < count; i++) {
      const method = dv.getUint16(p + 10, true), csize = dv.getUint32(p + 20, true);
      const nlen = dv.getUint16(p + 28, true), xlen = dv.getUint16(p + 30, true), clen = dv.getUint16(p + 32, true);
      const off = dv.getUint32(p + 42, true);
      const name = dec.decode(u8.subarray(p + 46, p + 46 + nlen));
      const start = off + 30 + dv.getUint16(off + 26, true) + dv.getUint16(off + 28, true);
      const raw = u8.subarray(start, start + csize);
      files.set(name, method === 0 ? raw.slice() : await inflate(raw));
      p += 46 + nlen + xlen + clen;
    }
    return files;
  }

  async function zip(files) {
    const enc = new TextEncoder();
    const chunks = [], central = [];
    let offset = 0;
    for (const [name, data] of files) {
      const nm = enc.encode(name), comp = await deflate(data), crc = crc32(data);
      const h = new DataView(new ArrayBuffer(30));
      h.setUint32(0, 0x04034b50, true); h.setUint16(4, 20, true); h.setUint16(6, 0x0800, true); h.setUint16(8, 8, true);
      h.setUint16(10, 0, true); h.setUint16(12, 0x21, true);
      h.setUint32(14, crc, true); h.setUint32(18, comp.length, true); h.setUint32(22, data.length, true); h.setUint16(26, nm.length, true);
      chunks.push(new Uint8Array(h.buffer), nm, comp);
      const c = new DataView(new ArrayBuffer(46));
      c.setUint32(0, 0x02014b50, true); c.setUint16(4, 20, true); c.setUint16(6, 20, true); c.setUint16(8, 0x0800, true); c.setUint16(10, 8, true);
      c.setUint16(12, 0, true); c.setUint16(14, 0x21, true);
      c.setUint32(16, crc, true); c.setUint32(20, comp.length, true); c.setUint32(24, data.length, true); c.setUint16(28, nm.length, true);
      c.setUint32(42, offset, true);
      central.push(new Uint8Array(c.buffer), nm);
      offset += 30 + nm.length + comp.length;
    }
    const csize = central.reduce((a, b) => a + b.length, 0);
    const end = new DataView(new ArrayBuffer(22));
    end.setUint32(0, 0x06054b50, true); end.setUint16(8, files.size, true); end.setUint16(10, files.size, true);
    end.setUint32(12, csize, true); end.setUint32(16, offset, true);
    const all = [...chunks, ...central, new Uint8Array(end.buffer)];
    const out = new Uint8Array(all.reduce((a, b) => a + b.length, 0));
    let o = 0;
    for (const a of all) { out.set(a, o); o += a.length; }
    return out;
  }

  // ---------------------------------------------------------------- sheet xml patching
  const xmlEsc = s => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '');
  const CELL = /<c r="([A-Z]+[0-9]+)"([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g;

  /* values: Map(ref -> number|string). Cached formula results are dropped so nothing stale is shown before recalculation. */
  function patchSheet(xml, values) {
    const pending = new Set(values.keys());
    const out = xml.replace(CELL, (m, ref, attrs, inner) => {
      if (values.has(ref)) {
        if (inner && inner.indexOf('<f') >= 0) throw new Error('محاولة الكتابة فوق معادلة: ' + ref);
        pending.delete(ref);
        const style = (attrs.match(/\ss="\d+"/) || [''])[0];
        const v = values.get(ref);
        if (v === null) return `<c r="${ref}"${style}/>`;
        if (typeof v === 'number') return `<c r="${ref}"${style}><v>${v}</v></c>`;
        return `<c r="${ref}"${style} t="inlineStr"><is><t xml:space="preserve">${xmlEsc(v)}</t></is></c>`;
      }
      if (inner && inner.indexOf('<f') >= 0) {
        return `<c r="${ref}"${attrs.replace(/\st="[^"]*"/, '')}>${inner.replace(/<v(?:\s[^>]*)?>[\s\S]*?<\/v>|<v\/>/, '')}</c>`;
      }
      return m;
    });
    for (const ref of pending) if (values.get(ref) === null) pending.delete(ref);
    if (pending.size) throw new Error('خلايا غير موجودة في القالب: ' + [...pending].slice(0, 5).join(', '));
    return out;
  }

  // ---------------------------------------------------------------- model -> rows
  /* plaster height of a space: its own value, else the one set for its floor */
  const plasterH = (pr, sp) => num(sp.plaster && sp.plaster.hgt) ?? num(pr && pr.levels && pr.levels[sp.level] && pr.levels[sp.level].plHgt);
  /* «يحتاج مراجعة» + site note of one entered item, as workbook columns */
  const flag = (...items) => {
    const rev = items.some(x => x && x.rev), note = items.map(x => txt(x && x.note)).filter(Boolean).join(' | ');
    return { ...(rev ? { rev: YES } : {}), ...(note ? { note } : {}) };
  };
  const addNote = (row, extra) => { if (extra) row.note = row.note ? row.note + ' | ' + extra : extra; return row; };
  function reviewCount(sp) {
    let n = 0;
    const c = x => { if (x && x.rev) n++; };
    (sp.floors || []).forEach(c); ((sp.ceil && sp.ceil.parts) || []).forEach(c);
    (sp.walls || []).forEach(w => { c(w); (w.ops || []).forEach(c); (w.feats || []).forEach(c); });
    return n;
  }
  function wallRecords(sp) {
    /* per wall: finish records grouped by material (one record per wall + material, as the workbook requires) */
    const res = [];
    (sp.walls || []).forEach((w, i) => {
      const code = 'W' + (i + 1);
      const len = num(w.len);
      const mainMat = txt(w.mat) || txt(sp.wall && sp.wall.mat);
      const mainH = num(w.hgt) ?? num(sp.wall && sp.wall.hgt);
      const groups = new Map();
      let manual = 0;
      const hasMain = !!mainMat && mainMat !== '-' && len !== null;
      if (hasMain) groups.set(mainMat, [{ len, hgt: mainH, cnt: null, main: true, src: w }]);
      (w.feats || []).forEach(f => {
        if (f.unit === 'lm') return;
        const fl = num(f.len), fh = num(f.hgt), fm = txt(f.mat);
        if (!fm || fl === null) return;
        if (!groups.has(fm)) groups.set(fm, []);
        groups.get(fm).push({ len: fl, hgt: fh, cnt: num(f.cnt), src: f });
        if (f.ded !== false && fh !== null) manual += fl * fh * (num(f.cnt) ?? 1);
      });
      res.push({ code, len, mainMat: hasMain ? mainMat : '', groups, manual: hasMain ? manual : 0, wall: w });
    });
    return res;
  }

  function buildRows(pr) {
    const R = { SP: [], FL: [], WL: [], OP: [], LN: [], CL: [], PL: [] };
    for (const sp of pr.spaces || []) {
      const code = txt(sp.code);
      if (!code) continue;
      R.SP.push({ code, nm: txt(sp.name), lvl: txt(sp.level), note: txt(sp.note) });

      const parts = (sp.floors || []).filter(f => num(f.dir) !== null || (num(f.len) !== null && num(f.wid) !== null));
      parts.forEach((f, i) => {
        if (!txt(f.mat)) return;
        R.FL.push({ sp: code, mat: txt(f.mat), desc: txt(f.desc) || 'جزء ' + (i + 1), len: num(f.len), wid: num(f.wid), cnt: num(f.cnt), dir: num(f.dir), ...flag(f) });
      });

      const ce = sp.ceil || {};
      if (!pr.noCeil && txt(ce.mat)) {
        const cparts = ce.same === false ? (ce.parts || []).filter(f => num(f.dir) !== null || (num(f.len) !== null && num(f.wid) !== null)) : parts;
        cparts.forEach((f, i) => R.CL.push(addNote({ sp: code, mat: txt(ce.mat), spec: txt(ce.spec), desc: (txt(f.desc) || 'جزء') + ' ' + (i + 1),
          len: num(f.len), wid: num(f.wid), cnt: num(f.cnt), dir: num(f.dir), note: num(ce.hgt) !== null ? 'ارتفاع السقف ' + num(ce.hgt) + ' م' : '' },
          ce.same === false ? flag(f).note : '')));
        if (ce.same === false) cparts.forEach((f, i) => { if (f.rev) R.CL[R.CL.length - cparts.length + i].rev = YES; });
      }

      const recs = wallRecords(sp);
      const plOn = sp.plaster && sp.plaster.on !== false && txt(sp.plaster.mat) && plasterH(pr, sp) !== null;
      const skOn = sp.skirt && sp.skirt.on !== false && txt(sp.skirt.mat);
      let perim = 0;
      recs.forEach(rec => {
        if (rec.len !== null) perim += rec.len;
        for (const [mat, list] of rec.groups) {
          const row = { sp: code, wall: rec.code, mat };
          if (list.length === 1 && list[0].hgt !== null) {
            row.len = list[0].len; row.hgt = list[0].hgt; row.cnt = list[0].cnt;
          } else {
            row.meth = 'يدوي';
            row.dir = list.reduce((a, x) => a + x.len * (x.hgt ?? 0) * (x.cnt ?? 1), 0);
            row.note = 'مساحة مجمّعة من التطبيق (أكثر من جزء بنفس الخامة)';
            if (list.some(x => x.hgt === null)) { row.dir = null; row.len = list[0].len; }
          }
          if (mat === rec.mainMat && rec.manual > 0) { row.man = Math.round(rec.manual * 1e4) / 1e4; row.note = (row.note ? row.note + ' | ' : '') + 'الخصم اليدوي = أجزاء بخامة أخرى على نفس الحائط'; }
          const fl = flag(...list.map(x => x.src));
          if (fl.rev) row.rev = YES;
          R.WL.push(addNote(row, fl.note));
        }
        if (plOn && rec.len !== null) R.PL.push({ sp: code, wall: rec.code, mat: txt(sp.plaster.mat), len: rec.len, hgt: plasterH(pr, sp), ...flag(rec.wall) });
        (rec.wall.ops || []).forEach((o, k) => {
          if (num(o.wid) === null && num(o.hgt) === null) return;
          const isWin = /نافذة|شباك/.test(txt(o.typ));
          R.OP.push({ code: `${code}-${rec.code}-${k + 1}`, sp: code, wall: rec.code, typ: txt(o.typ) || 'باب', wid: num(o.wid), hgt: num(o.hgt), cnt: num(o.cnt),
            dfin: (o.fin !== false && rec.mainMat) ? YES : NO, tgt: (o.fin !== false && rec.mainMat && rec.groups.size > 1) ? rec.mainMat : '',
            dpl: (o.pl !== false && plOn) ? YES : NO,
            dsk: ((o.sk === undefined ? !(isWin && sp.skirt && sp.skirt.underWin) : o.sk) && skOn) ? YES : NO, ...flag(o) });
        });
        (rec.wall.feats || []).forEach((f, k) => {
          if (f.unit !== 'lm' || !txt(f.mat) || num(f.len) === null) return;
          R.LN.push({ sp: code, mat: txt(f.mat), desc: `على الحائط ${rec.code} (${k + 1})`, dir: num(f.len), cnt: num(f.cnt), ...flag(f) });
        });
      });
      if (skOn && perim > 0) R.LN.push({ sp: code, mat: txt(sp.skirt.mat), desc: 'وزرة — محيط الفراغ', dir: Math.round(perim * 1e4) / 1e4, autoq: YES });
      if (num(sp.floorStrip) > 0) R.LN.push({ sp: code, mat: txt(sp.floorStripMat) || matName(pr, 'LinMat', 'شرائح ستانلس ستيل'), desc: 'فاصل بين خامات الأرضية', dir: num(sp.floorStrip) });
      const co = sp.corners || {};
      if (num(co.cnt) > 0 && num(co.hgt) > 0) R.LN.push({ sp: code, mat: matName(pr, 'LinMat', 'زوايا ستانلس للحوائط'), desc: 'زوايا الحوائط', dir: num(co.hgt), cnt: num(co.cnt) });
      const st = sp.stair || {};
      if (sp.kind === 'stair' && num(st.steps) > 0 && num(st.width) > 0) {
        const d = t => [t, txt(st.mat), num(st.tread) !== null ? 'نائمة ' + num(st.tread) : '', num(st.riser) !== null ? 'قائمة ' + num(st.riser) : ''].filter(Boolean).join(' — ');
        R.LN.push({ sp: code, mat: matName(pr, 'LinMat', 'نائمة درج'), desc: d('نوائم'), dir: num(st.width), cnt: num(st.steps) });
        R.LN.push({ sp: code, mat: matName(pr, 'LinMat', 'قائمة درج'), desc: d('قوائم'), dir: num(st.width), cnt: num(st.steps) });
      }
      if (sp.kind === 'stair' && num(st.skirtLen) > 0) R.LN.push({ sp: code, mat: matName(pr, 'LinMat', 'وزرة درج'), desc: 'وزرة الدرج', dir: num(st.skirtLen) });
    }
    return R;
  }

  /* Material lists of a project: [{name, unit, base}] per list. `base` is the template's original name, kept through renames
     so the items the export itself writes (stair treads, stainless strips…) follow a corrected name. */
  const MAT_LISTS = ['FloorMat', 'WallMat', 'LinMat', 'CeilMat', 'PlMat', 'OpenTypes'];
  const SYSTEM = { LinMat: ['شرائح ستانلس ستيل', 'زوايا ستانلس للحوائط', 'نائمة درج', 'قائمة درج', 'وزرة درج'] };
  const DEF_UNIT = { FloorMat: 'م²', WallMat: 'م²', LinMat: 'م.ط', CeilMat: 'م²', PlMat: 'م²' };
  function materials(pr) {
    const out = {};
    for (const list of MAT_LISTS) {
      const def = L.lists[list];
      if (pr.mats && Array.isArray(pr.mats[list])) { out[list] = pr.mats[list]; continue; }
      out[list] = def.items.map((name, i) => ({ name, unit: (def.units && def.units[i]) || '', base: name }));
      ((pr.custom && pr.custom[list]) || []).forEach(c => {
        if (txt(c.name) && !out[list].some(m => m.name === txt(c.name))) out[list].push({ name: txt(c.name), unit: txt(c.unit) });
      });
    }
    return out;
  }
  const matName = (pr, list, base) => (materials(pr)[list].find(m => m.base === base) || { name: base }).name;
  /* cells to write in the lists sheet; null = clear. Untouched lists are left exactly as the template has them. */
  function listCells(pr) {
    const cells = [];
    const mats = materials(pr);
    for (const list of MAT_LISTS) {
      const def = L.lists[list], eff = mats[list].filter(m => txt(m.name));
      const same = eff.length === def.items.length && eff.every((m, i) => m.name === def.items[i] && (!def.ucol || txt(m.unit) === ((def.units && def.units[i]) || '')));
      if (same) continue;
      if (eff.length > L.listRN - L.listR0 + 1) throw new Error(`قائمة «${list}» بها ${eff.length} خامة والقالب يسع ${L.listRN - L.listR0 + 1} فقط. احذف ما لا تحتاجه من شاشة الخامات.`);
      for (let r = L.listR0; r <= L.listRN; r++) {
        const m = eff[r - L.listR0];
        cells.push([def.col + r, m ? txt(m.name) : null]);
        if (def.ucol) cells.push([def.ucol + r, m && txt(m.unit) ? txt(m.unit) : null]);
      }
    }
    return cells;
  }

  // ---------------------------------------------------------------- export
  async function exportXlsx(pr) {
    const rows = buildRows(pr);
    for (const k of Object.keys(rows)) {
      if (rows[k].length > L.cap[k]) throw new Error(`عدد السجلات في «${L.sheets[k]}» (${rows[k].length}) أكبر من سعة القالب (${L.cap[k]}).`);
    }
    const bin = Uint8Array.from(atob(globalThis.BOQ_TEMPLATE_B64), c => c.charCodeAt(0));
    const files = await unzip(bin);
    const dec = new TextDecoder(), enc = new TextEncoder();
    const wbXml = dec.decode(files.get('xl/workbook.xml'));
    const rels = dec.decode(files.get('xl/_rels/workbook.xml.rels'));
    const pathOf = name => {
      const m = [...wbXml.matchAll(/<sheet [^>]*>/g)].map(x => x[0]).find(s => s.includes(`name="${xmlEsc(name)}"`));
      if (!m) throw new Error('ورقة غير موجودة: ' + name);
      const rid = m.match(/r:id="([^"]+)"/)[1];
      const rel = [...rels.matchAll(/<Relationship [^>]*>/g)].map(x => x[0]).find(s => s.includes(`Id="${rid}"`));
      return 'xl/' + rel.match(/Target="([^"]+)"/)[1].replace(/^\/xl\//, '');
    };
    const maps = {};
    const put = (sheetKey, ref, v) => {
      if (v === null || v === undefined || v === '') return;
      (maps[sheetKey] = maps[sheetKey] || new Map()).set(ref, v);
    };
    for (const k of Object.keys(rows)) {
      rows[k].forEach((row, i) => { for (const [f, v] of Object.entries(row)) { if (!L.cols[k][f]) throw new Error('عمود غير معروف ' + f); put(k, L.cols[k][f] + (L.r0 + i), v); } });
    }
    for (const [ref, v] of listCells(pr)) (maps.LS = maps.LS || new Map()).set(ref, v);
    for (const [f, ref] of Object.entries(L.pd)) put('PD', ref, f === 'method' ? 'رفع من الطبيعة' : txt(pr[f]));
    // every worksheet goes through patchSheet so cached formula values are cleared everywhere
    const byPath = new Map(Object.entries(maps).map(([k, m]) => [pathOf(L.sheets[k]), m]));
    for (const [name, data] of files) {
      if (!/^xl\/worksheets\/sheet\d+\.xml$/.test(name)) continue;
      files.set(name, enc.encode(patchSheet(dec.decode(data), byPath.get(name) || new Map())));
    }
    files.set('xl/workbook.xml', enc.encode(wbXml.replace(/<calcPr([^>]*?)\/>/, (m, a) => `<calcPr${a.replace(/\sfullCalcOnLoad="[^"]*"/, '')} fullCalcOnLoad="1"/>`)));
    files.delete('xl/calcChain.xml');
    if (files.has('[Content_Types].xml')) {
      files.set('[Content_Types].xml', enc.encode(dec.decode(files.get('[Content_Types].xml')).replace(/<Override PartName="\/xl\/calcChain\.xml"[^>]*\/>/, '')));
    }
    files.set('xl/_rels/workbook.xml.rels', enc.encode(rels.replace(/<Relationship [^>]*calcChain[^>]*\/>/, '')));
    return { bytes: await zip(files), counts: Object.fromEntries(Object.entries(rows).map(([k, v]) => [k, v.length])) };
  }

  // ---------------------------------------------------------------- quick quantities for the on-screen summary
  function quantities(sp, pr) {
    const q = { floor: 0, wallGross: 0, wallDed: 0, plGross: 0, plDed: 0, perim: 0, skDed: 0, openings: 0 };
    (sp.floors || []).forEach(f => { const c = num(f.cnt) ?? 1; q.floor += num(f.dir) !== null ? num(f.dir) * c : (num(f.len) ?? 0) * (num(f.wid) ?? 0) * c; });
    const plOn = sp.plaster && sp.plaster.on !== false && plasterH(pr, sp) !== null;
    const skOn = sp.skirt && sp.skirt.on !== false;
    wallRecords(sp).forEach(rec => {
      if (rec.len === null) return;
      q.perim += rec.len;
      const main = rec.mainMat ? rec.groups.get(rec.mainMat)[0] : null;
      if (main && main.hgt !== null) { q.wallGross += main.len * main.hgt; q.wallDed += rec.manual; }
      if (plOn) q.plGross += rec.len * plasterH(pr, sp);
      (rec.wall.ops || []).forEach(o => {
        const w = num(o.wid), h = num(o.hgt), c = num(o.cnt) ?? 1;
        if (w === null) return;
        q.openings += c;
        const a = h !== null ? w * h * c : 0;
        const isWin = /نافذة|شباك/.test(txt(o.typ));
        if (o.fin !== false && main) q.wallDed += a;
        if (o.pl !== false && plOn) q.plDed += a;
        if ((o.sk === undefined ? !(isWin && sp.skirt && sp.skirt.underWin) : o.sk) && skOn) q.skDed += w * c;
      });
    });
    return q;
  }

  globalThis.BOQ = { num, txt, buildRows, exportXlsx, quantities, wallRecords, reviewCount, unzip, zip, materials, matName, MAT_LISTS, SYSTEM, DEF_UNIT, listCap: L.listRN - L.listR0 + 1 };
})();
