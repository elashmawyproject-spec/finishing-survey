// Node test: export a made-up project through the app's export code; then tools/recalc.ps1 (Excel) and tools/check_export.py.
const path = require('path'), fs = require('fs');
const app = process.argv[2];
require(path.join(app, 'js', 'layout.js'));
require(path.join(app, 'js', 'template.js'));
require(path.join(app, 'js', 'export.js'));
const win = { typ: 'نافذة', wid: '0.80', hgt: '1.50', cnt: '2' };
// invented measurements only: this repository is public, never paste real site data here
const project = {
  name: 'Test Project', location: 'موقع تجريبي', surveyor: 'اختبار', date: '2026-10-11',
  levels: { 'ثاني': { plHgt: '3' } },   // floor-wide plaster height: T-01 inherits it, T-02 overrides with its own
  custom: { WallMat: [{ name: 'كسوة ألمنيوم', unit: 'م²' }], LinMat: [{ name: 'نقطة إنارة', unit: 'عدد' }] },
  spaces: [
    { code: 'T-01', name: 'ROOM A', level: 'ثاني', kind: 'room',
      floors: [{ mat: 'سيراميك أرضيات', len: '4.2', wid: '3.5' }, { mat: 'سيراميك أرضيات', len: '1.2', wid: '2' }],
      ceil: { mat: 'سقف جبسي', hgt: '2.8' }, wall: { mat: 'دهانات', hgt: '2.8' }, plaster: { on: true, mat: 'لياسة أسمنتية — حوائط داخلية', hgt: '' },
      skirt: { on: true, mat: 'وزرات سيراميك', underWin: false },
      walls: [{ len: '4.2', ops: [win] }, { len: '3.5' }, { len: '3', ops: [{ typ: 'باب', wid: '0.9', hgt: '2.2' }] }, { len: '5.5' },
              { len: '1.2', ops: [{ typ: 'باب', wid: '1', hgt: '2.2' }] }, { len: '2' }] },
    { code: 'T-02', name: 'BATH A', level: 'ثاني', kind: 'room',
      floors: [{ mat: 'سيراميك أرضيات', len: '1.5', wid: '1.8' }, { mat: 'رخام', len: '1.5', wid: '0.1' }, { mat: 'موزايكو', len: '1', wid: '1.5' }],
      wall: { mat: 'سيراميك حوائط الحمامات', hgt: '2.8' }, plaster: { on: true, mat: 'لياسة أسمنتية — حوائط داخلية', hgt: '3.1' }, skirt: { on: false },
      walls: [{ len: '8.6', ops: [{ typ: 'باب', wid: '0.8', hgt: '2.2' }] }] },
    // features: marble panel (m2, deducted from paint) + marble strip (lm) + stainless + a window with skirting under it
    { code: 'GF-02', name: 'MAIN SALON', level: 'أرضي', kind: 'room',
      floors: [{ mat: 'رخام', dir: '40', note: 'مساحة من اللوحة' }], floorStrip: '3.2',
      extras: [{ mat: 'رخام كعنصر طولي', qty: '(0.70 + 2.52 + 2.52) × 3', desc: 'رخام حلق باب' }, { mat: 'نقطة إنارة', qty: '3', cnt: '2' }, { mat: 'رخام كعنصر طولي', qty: 'abc' }], corners: { cnt: '4', hgt: '3' },
      wall: { mat: 'دهانات', hgt: '3' }, plaster: { on: true, mat: 'لياسة أسمنتية — حوائط داخلية', hgt: '3.2' }, skirt: { on: true, mat: 'وزرات رخام', underWin: true },
      walls: [{ len: '6', rev: true, note: 'الطول يحتاج تأكيدًا', ops: [{ typ: 'نافذة', wid: '2', hgt: '1.5', rev: true, note: 'شباك مغلق بجبس بورد؟' }],
                feats: [{ unit: 'm2', mat: 'رخام حوائط', len: '2', hgt: '3' }, { unit: 'lm', mat: 'رخام كعنصر طولي', len: '6' }, { unit: 'm2', mat: 'كسوة ألمنيوم', len: '1', hgt: '1', ded: false }] },
              { len: '5', ops: [{ typ: 'باب', wid: '1', hgt: '2.2' }] }, { len: '2', noSk: true, ops: [{ typ: 'باب', wid: '0.9', hgt: '2.2' }] }] },
    { code: 'GF-ST-01', name: 'STAIR ST-01', level: 'أرضي', kind: 'stair', stair: { steps: '20', width: '1.2', tread: '0.3', riser: '0.17', mat: 'رخام', skirtLen: '9' },
      floors: [{ mat: 'رخام', desc: 'بسطة', len: '1.2', wid: '2.5' }], walls: [] },
  ],
};
// second file: the same project after the materials screen edited the lists (rename, delete, new item, changed unit)
function editedLists() {
  const p = JSON.parse(JSON.stringify(project));
  p.mats = globalThis.BOQ.materials(p);
  const rename = (list, from, to) => {
    p.mats[list].find(m => m.name === from).name = to;
    const fix = o => { if (o && o.mat === from) o.mat = to; };
    p.spaces.forEach(sp => { (sp.floors || []).forEach(fix); fix(sp.wall); fix(sp.skirt); fix(sp.ceil); fix(sp.plaster); (sp.walls || []).forEach(w => (w.feats || []).forEach(fix)); });
  };
  rename('FloorMat', 'سيراميك أرضيات', 'سيراميك 60×60');
  rename('WallMat', 'دهانات', 'دهان بلاستيك');
  rename('LinMat', 'نائمة درج', 'نائمة درج رخام');
  rename('LinMat', 'شرائح ستانلس ستيل', 'فاصل ستانلس');
  p.mats.FloorMat = p.mats.FloorMat.filter(m => !['تيرازو', 'موكيت'].includes(m.name));
  p.mats.FloorMat.push({ name: 'بلاط أسمنتي 20×20', unit: 'م²' });
  p.mats.OpenTypes.push({ name: 'شباك مغلق' });
  p.mats.CeilMat.find(m => m.name === 'بلاطات معلقة').unit = 'عدد';
  p.noCeil = true;   // ceilings not required: no ceiling rows at all
  return p;
}
(async () => {
  const out = process.argv[3];
  for (const [file, pr] of [[out, project], [out.replace(/\.xlsx$/, '') + '_lists.xlsx', editedLists()]]) {
    const r = await globalThis.BOQ.exportXlsx(pr);
    fs.writeFileSync(file, Buffer.from(r.bytes));
    console.log('written', path.basename(file), r.bytes.length, JSON.stringify(r.counts));
  }
})().catch(e => { console.error(e); process.exit(1); });
