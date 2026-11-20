const assert = require("node:assert/strict");
const test = require("node:test");
const G = require("../src/geometry.js");

const close = (a, b, m) => assert.ok(Math.abs(a - b) < 1e-6, `${m || ""} ${a} != ${b}`);
const ref = { left: 100, top: 100, width: 200, height: 100 }; // master: x100-300, y100-200
const S = { left: 0, top: 0, width: 960, height: 540 };

test("align terhadap master", () => {
  const r = { left: 10, top: 10, width: 50, height: 20 };
  assert.equal(G.align(r, ref, "left").left, 100);
  assert.equal(G.align(r, ref, "right").left, 250);
  assert.equal(G.align(r, ref, "center").left, 175);
  assert.equal(G.align(r, ref, "top").top, 100);
  assert.equal(G.align(r, ref, "bottom").top, 180);
  assert.equal(G.align(r, ref, "middle").top, 140);
});

test("align terhadap slide", () => {
  const r = { left: 10, top: 10, width: 60, height: 40 };
  assert.equal(G.align(r, S, "center").left, 450);
  assert.equal(G.align(r, S, "bottom").top, 500);
});

test("dock: hanya objek di sisi asal yang bergerak, lalu menempel", () => {
  const right = { left: 500, top: 0, width: 50, height: 50 };
  const left = { left: 0, top: 0, width: 50, height: 50 };
  assert.equal(G.dock(right, ref, "left").left, 300);   // menempel di sisi kanan master
  assert.equal(G.dock(left, ref, "left").left, 0);      // sudah di kiri: tidak berubah
  assert.equal(G.dock(left, ref, "right").left, 50);    // kanan: berhenti menempel di kiri master (100-50)
  const below = { left: 0, top: 400, width: 10, height: 30 };
  assert.equal(G.dock(below, ref, "up").top, 200);
  const above = { left: 0, top: 0, width: 10, height: 30 };
  assert.equal(G.dock(above, ref, "down").top, 70);
});

test("dock mode slide = ke tepi slide", () => {
  const r = { left: 100, top: 100, width: 50, height: 50 };
  assert.equal(G.dock(r, S, "left", true).left, 0);
  assert.equal(G.dock(r, S, "right", true).left, 910);
  assert.equal(G.dock(r, S, "up", true).top, 0);
  assert.equal(G.dock(r, S, "down", true).top, 490);
});

test("stretch: sisi ditarik, sisi seberang tetap", () => {
  const r = { left: 400, top: 300, width: 100, height: 50 };
  const a = G.stretch(r, ref, "left");
  assert.equal(a.left, 100); assert.equal(a.left + a.width, 500);
  const r2 = { left: 150, top: 300, width: 50, height: 50 };
  const b = G.stretch(r2, ref, "right");
  assert.equal(b.left, 150); assert.equal(b.left + b.width, 300);
});

test("stretch right: lebar tidak pernah di bawah minimum", () => {
  const r = { left: 400, top: 300, width: 100, height: 50 };
  assert.ok(G.stretch(r, ref, "right").width >= 1);
});

test("stretch up/down", () => {
  const r = { left: 0, top: 300, width: 10, height: 50 };
  const u = G.stretch(r, ref, "up");
  assert.equal(u.top, 100); assert.equal(u.top + u.height, 350);
  const r2 = { left: 0, top: 10, width: 10, height: 50 };
  const d = G.stretch(r2, ref, "down");
  assert.equal(d.top, 10); assert.equal(d.top + d.height, 200);
});

test("fill: menutup celah ke sisi master", () => {
  const r = { left: 400, top: 120, width: 100, height: 40 }; // kanan master (master.right=300)
  const f = G.fill(r, ref, "left");
  assert.equal(f.left, 300); assert.equal(f.left + f.width, 500);
  const l = { left: 0, top: 120, width: 50, height: 40 };    // kiri master (master.left=100)
  const g = G.fill(l, ref, "right");
  assert.equal(g.left, 0); assert.equal(g.left + g.width, 100);
  const below = { left: 0, top: 400, width: 10, height: 30 };
  const u = G.fill(below, ref, "up");
  assert.equal(u.top, 200); assert.equal(u.top + u.height, 430);
  const above = { left: 0, top: 0, width: 10, height: 30 };
  const d = G.fill(above, ref, "down");
  assert.equal(d.top, 0); assert.equal(d.top + d.height, 100);
  // objek di sisi yang salah tidak berubah
  assert.deepEqual(G.fill(l, ref, "left"), l);
});

test("distribute: objek terluar tetap, jarak sama", () => {
  const rs = [
    { left: 0, top: 0, width: 100, height: 10 },
    { left: 150, top: 0, width: 50, height: 10 },
    { left: 400, top: 0, width: 100, height: 10 },
  ];
  const o = G.distribute(rs, "h");
  assert.equal(o[0].left, 0); assert.equal(o[2].left, 400);
  const g1 = o[1].left - (o[0].left + o[0].width);
  const g2 = o[2].left - (o[1].left + o[1].width);
  close(g1, g2);
});

test("distribute tidak bergantung urutan input", () => {
  const rs = [
    { left: 400, top: 0, width: 100, height: 10 },
    { left: 0, top: 0, width: 100, height: 10 },
    { left: 150, top: 0, width: 50, height: 10 },
  ];
  const o = G.distribute(rs, "h");
  assert.equal(o[1].left, 0); assert.equal(o[0].left, 400);
  close(o[2].left - 100, 400 - (o[2].left + 50));
});

test("distribute dalam slide (Ctrl)", () => {
  const rs = [
    { left: 300, top: 0, width: 100, height: 10 },
    { left: 500, top: 0, width: 100, height: 10 },
  ];
  const o = G.distribute(rs, "h", { w: 960, h: 540 });
  assert.equal(o[0].left, 0); assert.equal(o[1].left + o[1].width, 960);
});

test("distribute: validasi jumlah objek", () => {
  assert.throws(() => G.distribute([{ left: 0, top: 0, width: 1, height: 1 }, { left: 5, top: 0, width: 1, height: 1 }], "h"));
});

test("slice 2x2 dengan jarak", () => {
  const cells = G.sliceOrMultiply({ left: 0, top: 0, width: 220, height: 220 }, { mode: "slice", rows: 2, cols: 2, rowGap: 20, colGap: 20 });
  assert.equal(cells.length, 4);
  assert.equal(cells[0].width, 100); assert.equal(cells[3].left, 120); assert.equal(cells[3].top, 120);
});

test("multiply menjaga ukuran asli", () => {
  const cells = G.sliceOrMultiply({ left: 10, top: 10, width: 50, height: 30 }, { mode: "multiply", rows: 2, cols: 3, rowGap: 5, colGap: 10 });
  assert.equal(cells.length, 6);
  assert.deepEqual(cells[0], { left: 10, top: 10, width: 50, height: 30 });
  assert.equal(cells[5].left, 10 + 2 * 60); assert.equal(cells[5].top, 10 + 35);
});

test("slice: jarak terlalu besar ditolak", () => {
  assert.throws(() => G.sliceOrMultiply({ left: 0, top: 0, width: 10, height: 10 }, { mode: "slice", rows: 2, cols: 2, rowGap: 50, colGap: 50 }));
});

test("matrix auto 4 objek -> 2x2, urut per baris", () => {
  const rs = [0, 1, 2, 3].map((i) => ({ left: i * 7, top: i * 3, width: 100, height: 50 }));
  const o = G.matrix(rs, { mode: "auto", rowGap: 10, colGap: 20, sizing: "fit", align: [1, 1] });
  assert.equal(o[0].left, 0); assert.equal(o[1].left, 120); assert.equal(o[2].left, 0); assert.equal(o[3].left, 120);
  assert.equal(o[2].top, 60);
});

test("matrix kolom=3 dari 5 objek -> 2 baris", () => {
  const rs = Array.from({ length: 5 }, () => ({ left: 0, top: 0, width: 10, height: 10 }));
  const o = G.matrix(rs, { mode: "cols", cols: 3, rowGap: 0, colGap: 0, sizing: "fit", align: [0, 0] });
  assert.equal(o[3].top, 10); assert.equal(o[3].left, 0); assert.equal(o[4].left, 10);
});

test("matrix: objek kecil di sel (tengah) dan sizing equal", () => {
  const rs = [{ left: 0, top: 0, width: 100, height: 100 }, { left: 0, top: 0, width: 50, height: 50 }];
  const o = G.matrix(rs, { mode: "rows", rows: 1, rowGap: 0, colGap: 0, sizing: "equal", align: [1, 1] });
  // cols=2, sel sama (100x100); objek kecil dipusatkan di sel ke-2
  assert.equal(o[1].left, 100 + 25); assert.equal(o[1].top, 25);
});

test("radius <-> adjustment", () => {
  close(G.adjToRadius(0.1, 200, 100), 10);
  close(G.radiusToAdj(10, 200, 100), 0.1);
  close(G.radiusToAdj(9999, 200, 100), 0.5); // dibatasi 0.5
});

test("satuan", () => {
  close(G.toPt(1, "in"), 72); close(G.toPt(2.54, "cm"), 72); close(G.fromPt(72, "in"), 1);
});

test("splitText", () => {
  assert.deepEqual(G.splitText("Halo dunia ini", 10), ["Halo dunia", "ini"]);
  assert.throws(() => G.splitText("abc", 0));
  assert.throws(() => G.splitText("abc", 3));
});

test("pickMaster", () => {
  const it = [{ width: 10, height: 10 }, { width: 50, height: 50 }, { width: 5, height: 5 }];
  assert.equal(G.pickMaster(it, "last"), 2);
  assert.equal(G.pickMaster(it, "first"), 0);
  assert.equal(G.pickMaster(it, "largest"), 1);
});
