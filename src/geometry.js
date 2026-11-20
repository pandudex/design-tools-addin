/* Logika geometri murni (tanpa Office.js) supaya bisa diuji di Node.
 * Rect = { left, top, width, height } dalam point.
 * "ref" = rect Master, atau rect slide ketika hanya 1 objek dipilih / Ctrl ditahan. */
(function (root) {
  "use strict";

  const EPS = 0.01;
  const MIN_SIZE = 1;

  const R = (r) => r.left + r.width;
  const B = (r) => r.top + r.height;
  const copy = (r) => ({ left: r.left, top: r.top, width: r.width, height: r.height });

  /* ---------- Align (ke tepi/tengah ref) ---------- */
  function align(r, ref, dir) {
    const o = copy(r);
    if (dir === "left") o.left = ref.left;
    else if (dir === "right") o.left = R(ref) - r.width;
    else if (dir === "center") o.left = ref.left + (ref.width - r.width) / 2;
    else if (dir === "top") o.top = ref.top;
    else if (dir === "bottom") o.top = B(ref) - r.height;
    else if (dir === "middle") o.top = ref.top + (ref.height - r.height) / 2;
    else throw new Error("align: arah tidak dikenal " + dir);
    return o;
  }

  /* ---------- Dock (geser sampai menempel pada Master) ----------
   * Hanya objek yang berada di sisi "asal" yang bergerak:
   * dock left menggeser objek yang ada di KANAN master ke kiri sampai menempel, dst.
   * Mode slide: objek dipindah ke tepi slide. */
  function dock(r, ref, dir, slideMode) {
    if (slideMode) {
      return align(r, ref, { left: "left", right: "right", up: "top", down: "bottom" }[dir]);
    }
    const o = copy(r);
    if (dir === "left" && r.left > R(ref) + EPS) o.left = R(ref);
    else if (dir === "right" && R(r) < ref.left - EPS) o.left = ref.left - r.width;
    else if (dir === "up" && r.top > B(ref) + EPS) o.top = B(ref);
    else if (dir === "down" && B(r) < ref.top - EPS) o.top = ref.top - r.height;
    return o;
  }

  /* ---------- Stretch (tarik satu sisi ke sisi yang sama milik ref) ---------- */
  function stretch(r, ref, dir) {
    const o = copy(r);
    if (dir === "left") { o.left = ref.left; o.width = R(r) - ref.left; }
    else if (dir === "right") { o.width = R(ref) - r.left; }
    else if (dir === "up") { o.top = ref.top; o.height = B(r) - ref.top; }
    else if (dir === "down") { o.height = B(ref) - r.top; }
    else throw new Error("stretch: arah tidak dikenal " + dir);
    o.width = Math.max(MIN_SIZE, o.width);
    o.height = Math.max(MIN_SIZE, o.height);
    return o;
  }

  /* ---------- Fill (tutup celah menuju sisi master) ----------
   * fill left: sisi kiri objek (yang ada di kanan master) ditarik ke sisi KANAN master.
   * Mode slide: sama dengan stretch ke tepi slide. */
  function fill(r, ref, dir, slideMode) {
    if (slideMode) return stretch(r, ref, dir);
    const o = copy(r);
    if (dir === "left" && r.left > R(ref) + EPS) { o.left = R(ref); o.width = R(r) - R(ref); }
    else if (dir === "right" && R(r) < ref.left - EPS) { o.width = ref.left - r.left; }
    else if (dir === "up" && r.top > B(ref) + EPS) { o.top = B(ref); o.height = B(r) - B(ref); }
    else if (dir === "down" && B(r) < ref.top - EPS) { o.height = ref.top - r.top; }
    return o;
  }

  /* ---------- Distribute ----------
   * Tanpa slide: objek terluar tetap, jarak antar objek disamakan.
   * Dengan slide (Ctrl): disebar mengisi seluruh area slide. */
  function distribute(rects, axis, slideSize) {
    const horiz = axis === "h";
    const pos = (r) => (horiz ? r.left : r.top);
    const size = (r) => (horiz ? r.width : r.height);
    const need = slideSize ? 2 : 3;
    if (rects.length < need) throw new Error(`Pilih minimal ${need} objek.`);

    const order = rects.map((r, i) => i).sort((a, b) => pos(rects[a]) - pos(rects[b]));
    const total = order.reduce((s, i) => s + size(rects[i]), 0);
    let start, end;
    if (slideSize) { start = 0; end = horiz ? slideSize.w : slideSize.h; }
    else {
      start = pos(rects[order[0]]);
      const last = rects[order[order.length - 1]];
      end = pos(last) + size(last);
    }
    const gap = (end - start - total) / (order.length - 1);
    const out = rects.map(copy);
    let p = start;
    for (const i of order) {
      if (horiz) out[i].left = p; else out[i].top = p;
      p += size(rects[i]) + gap;
    }
    return out;
  }

  /* ---------- Satuan ---------- */
  const UNIT_PT = { pt: 1, in: 72, cm: 72 / 2.54 };
  const toPt = (v, unit) => v * (UNIT_PT[unit] || 1);
  const fromPt = (v, unit) => v / (UNIT_PT[unit] || 1);

  /* ---------- Slice / Multiply ---------- */
  function sliceOrMultiply(r, o) {
    const { mode, rows, cols, rowGap, colGap } = o;
    if (!(rows >= 1 && cols >= 1)) throw new Error("Baris dan kolom minimal 1.");
    let cw, ch;
    if (mode === "slice") {
      cw = (r.width - (cols - 1) * colGap) / cols;
      ch = (r.height - (rows - 1) * rowGap) / rows;
      if (cw <= 0 || ch <= 0) throw new Error("Jarak terlalu besar untuk ukuran shape.");
    } else { cw = r.width; ch = r.height; }
    const out = [];
    for (let i = 0; i < rows; i++)
      for (let j = 0; j < cols; j++)
        out.push({ left: r.left + j * (cw + colGap), top: r.top + i * (ch + rowGap), width: cw, height: ch });
    return out;
  }

  /* ---------- Align in Matrix ----------
   * Diisi per baris dari atas ke bawah mengikuti urutan array (= urutan seleksi). */
  function matrix(rects, o) {
    const n = rects.length;
    if (n < 2) throw new Error("Pilih minimal 2 objek.");
    let rows, cols;
    if (o.mode === "rows") { rows = Math.max(1, o.rows); cols = Math.ceil(n / rows); }
    else if (o.mode === "cols") { cols = Math.max(1, o.cols); rows = Math.ceil(n / cols); }
    else { cols = Math.ceil(Math.sqrt(n)); rows = Math.ceil(n / cols); }

    const originX = Math.min(...rects.map((r) => r.left));
    const originY = Math.min(...rects.map((r) => r.top));
    const colW = Array(cols).fill(0), rowH = Array(rows).fill(0);
    rects.forEach((r, k) => {
      const c = k % cols, rw = Math.floor(k / cols);
      colW[c] = Math.max(colW[c], r.width);
      rowH[rw] = Math.max(rowH[rw], r.height);
    });
    if (o.sizing === "equal") {
      colW.fill(Math.max(...colW));
      rowH.fill(Math.max(...rowH));
    }
    const [av, ah] = o.align || [1, 1]; // 0=awal, 1=tengah, 2=akhir
    const xs = [], ys = [];
    let x = originX; for (let c = 0; c < cols; c++) { xs.push(x); x += colW[c] + o.colGap; }
    let y = originY; for (let rw = 0; rw < rows; rw++) { ys.push(y); y += rowH[rw] + o.rowGap; }

    return rects.map((r, k) => {
      const c = k % cols, rw = Math.floor(k / cols);
      return {
        left: xs[c] + ((colW[c] - r.width) * ah) / 2,
        top: ys[rw] + ((rowH[rw] - r.height) * av) / 2,
        width: r.width, height: r.height,
      };
    });
  }

  /* ---------- Rounded rectangle radius ----------
   * adjustment[0] diasumsikan pecahan 0..0.5 dari sisi terpendek. */
  const adjToRadius = (adj, w, h) => adj * Math.min(w, h);
  const radiusToAdj = (radius, w, h) => Math.max(0, Math.min(0.5, radius / Math.min(w, h)));

  /* ---------- Split teks ---------- */
  function splitText(text, pos) {
    if (!(pos > 0 && pos < text.length)) throw new Error("Taruh kursor di tengah teks dulu.");
    return [text.slice(0, pos).replace(/\s+$/, ""), text.slice(pos).replace(/^\s+/, "")];
  }

  /* ---------- Pilih Master ---------- */
  function pickMaster(items, mode) {
    if (mode === "first") return 0;
    if (mode === "largest") {
      let best = 0;
      items.forEach((s, i) => { if (s.width * s.height > items[best].width * items[best].height) best = i; });
      return best;
    }
    return items.length - 1; // default sesuai PDF: objek yang dipilih terakhir
  }

  const api = {
    align, dock, stretch, fill, distribute, sliceOrMultiply, matrix,
    adjToRadius, radiusToAdj, splitText, pickMaster, toPt, fromPt,
  };
  if (typeof module !== "undefined" && module.exports) module.exports = api;
  else root.Geo = api;
})(typeof self !== "undefined" ? self : this);
