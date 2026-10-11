/* Blank records for new projects and spaces. */
(function () {
  'use strict';
  const PLASTER = 'لياسة أسمنتية — حوائط داخلية';
  function blankSpace(code, name, level, kind, nWalls) {
    return {
      code, name, level, kind: kind || 'room', done: false, note: '',
      floors: [{}], floorStrip: '',
      ceil: { mat: '', hgt: '', same: true },
      wall: { mat: '', hgt: '' },
      plaster: { on: true, mat: PLASTER, hgt: '' },
      skirt: { on: kind !== 'stair', mat: '', underWin: false },
      walls: Array.from({ length: nWalls === undefined ? 4 : nWalls }, () => ({ len: '', ops: [], feats: [] })),
      corners: { cnt: '', hgt: '' },
      stair: { steps: '', width: '', tread: '', riser: '', mat: '', skirtLen: '' },
    };
  }

  globalThis.BOQ_SEED = { blankSpace, empty: () => ({ name: '', owner: '', location: '', date: '', surveyor: '', notes: '', custom: {}, spaces: [] }) };
})();
