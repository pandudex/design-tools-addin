/* Design Tools - glue Office.js (PowerPointApi 1.5+; Adjustments & ukuran slide butuh 1.10)
 * Logika geometri ada di geometry.js (diuji terpisah). */

const TEXT_TYPES = ["GeometricShape", "TextBox", "Placeholder", "Callout", "Freeform"];
const DEFAULT_SLIDE = { w: 960, h: 540 };
const EPS = 0.01;
const $ = (id) => document.getElementById(id);
const masterMode = () => document.querySelector('input[name="masterMode"]:checked').value;

function setStatus(msg, isErr = false) {
  const el = $("status");
  el.textContent = msg;
  el.className = isErr ? "err" : "";
}

/* ---------- akses Office.js ---------- */

async function readSelection(context, min = 1) {
  const shapes = context.presentation.getSelectedShapes();
  shapes.load("items/id,items/name,items/type,items/left,items/top,items/width,items/height");
  await context.sync();
  if (shapes.items.length < min) {
    throw new Error(min === 1 ? "Pilih minimal 1 objek dulu." : `Pilih minimal ${min} objek dulu.`);
  }
  return shapes.items;
}

async function slideSize(context) {
  try {
    const ps = context.presentation.pageSetup;
    ps.load("slideWidth,slideHeight");
    await context.sync();
    return { w: ps.slideWidth, h: ps.slideHeight };
  } catch (e) {
    return DEFAULT_SLIDE; // butuh PowerPointApi 1.10
  }
}

const rectOf = (s) => ({ left: s.left, top: s.top, width: s.width, height: s.height });

function applyRect(shape, r) {
  if (Math.abs(shape.left - r.left) > EPS) shape.left = r.left;
  if (Math.abs(shape.top - r.top) > EPS) shape.top = r.top;
  if (Math.abs(shape.width - r.width) > EPS) shape.width = r.width;
  if (Math.abs(shape.height - r.height) > EPS) shape.height = r.height;
}

/* Master = objek yang dipilih terakhir (sesuai spesifikasi). Pengaturan di panel bisa mengubahnya,
 * karena urutan seleksi dari PowerPoint tidak dijamin oleh dokumentasi API. */
async function withMaster(ctrl, fn) {
  let n = 0;
  await PowerPoint.run(async (context) => {
    const items = await readSelection(context);
    const slideMode = items.length === 1 || ctrl;
    const mIdx = Geo.pickMaster(items, masterMode());
    let ref;
    if (slideMode) {
      const s = await slideSize(context);
      ref = { left: 0, top: 0, width: s.w, height: s.h };
    } else ref = rectOf(items[mIdx]);
    items.forEach((s, i) => {
      if (!slideMode && i === mIdx) return; // master tidak ikut bergerak
      applyRect(s, fn(rectOf(s), ref, slideMode));
      n++;
    });
    await context.sync();
  });
  return n;
}

/* ---------- aksi: tepi ---------- */

const ACTIONS = {
  async align(v, ctrl) {
    const dir = { left: "left", center: "center", right: "right", top: "top", middle: "middle", bottom: "bottom" }[v];
    const n = await withMaster(ctrl, (r, ref) => Geo.align(r, ref, dir));
    setStatus(`Align ${v}: ${n} objek.`);
  },
  async dock(v, ctrl) {
    const n = await withMaster(ctrl, (r, ref, slide) => Geo.dock(r, ref, v, slide));
    setStatus(`Dock ${v}: ${n} objek.`);
  },
  async stretch(v, ctrl) {
    const n = await withMaster(ctrl, (r, ref) => Geo.stretch(r, ref, v));
    setStatus(`Stretch ${v}: ${n} objek.`);
  },
  async fill(v, ctrl) {
    const n = await withMaster(ctrl, (r, ref, slide) => Geo.fill(r, ref, v, slide));
    setStatus(`Fill ${v}: ${n} objek.`);
  },

  async distribute(v, ctrl) {
    await PowerPoint.run(async (context) => {
      const items = await readSelection(context, ctrl ? 2 : 3);
      const size = ctrl ? await slideSize(context) : null;
      const out = Geo.distribute(items.map(rectOf), v, size);
      items.forEach((s, i) => applyRect(s, out[i]));
      await context.sync();
    });
    setStatus(ctrl ? "Disebar di seluruh slide." : "Jarak antar objek disamakan.");
  },

  /* ---------- Slice / Multiply ---------- */
  async sliceMultiply() {
    const unit = $("smUnit").value;
    const opts = {
      mode: document.querySelector('input[name="smMode"]:checked').value,
      rows: parseInt($("smRows").value, 10),
      cols: parseInt($("smCols").value, 10),
      rowGap: Geo.toPt(parseFloat($("smRowGap").value) || 0, unit),
      colGap: Geo.toPt(parseFloat($("smColGap").value) || 0, unit),
    };
    let made = 0;
    await PowerPoint.run(async (context) => {
      const items = await readSelection(context);
      if (items.length !== 1) throw new Error("Pilih tepat 1 shape untuk Slice/Multiply.");
      const src = items[0];
      const cells = Geo.sliceOrMultiply(rectOf(src), opts);

      // API tidak punya "duplicate shape", jadi salinan dibuat ulang sebagai shape baru.
      src.fill.load("type,foregroundColor");
      src.lineFormat.load("visible,color,weight");
      const roundish = /round/i.test(src.name);
      let adj = null;
      if (roundish) { src.adjustments.load("count"); await context.sync(); if (src.adjustments.count > 0) adj = src.adjustments.get(0); }
      const slide = src.getParentSlide();
      await context.sync();

      const first = opts.mode === "multiply" ? 1 : 0; // multiply: sel pertama = shape asli
      for (let i = first; i < cells.length; i++) {
        const c = cells[i];
        const sh = slide.shapes.addGeometricShape(roundish ? "RoundRectangle" : "Rectangle", c);
        sh.name = `${src.name} ${opts.mode === "slice" ? "slice" : "copy"} ${i + 1}`;
        if (src.fill.type === "Solid") sh.fill.setSolidColor(src.fill.foregroundColor);
        if (src.lineFormat.visible) { sh.lineFormat.color = src.lineFormat.color; sh.lineFormat.weight = src.lineFormat.weight; }
        else sh.lineFormat.visible = false;
        if (adj) sh.adjustments.set(0, adj.value);
        made++;
      }
      if (opts.mode === "multiply") applyRect(src, cells[0]);
      else src.delete(); // slice: shape asli diganti potongan-potongannya
      await context.sync();
    });
    setStatus(`${opts.mode === "slice" ? "Slice" : "Multiply"} selesai: ${made} shape baru. Catatan: salinan berupa persegi/persegi rounded dengan warna isi & garis yang sama; teks dan efek tidak ikut tersalin.`);
  },

  /* ---------- Align in Matrix ---------- */
  async matrix() {
    const unit = $("mxUnit").value;
    const [av, ah] = document.querySelector('input[name="mxAlign"]:checked').value.split(",").map(Number);
    const opts = {
      mode: document.querySelector('input[name="mxMode"]:checked').value,
      rows: parseInt($("mxRows").value, 10) || 1,
      cols: parseInt($("mxCols").value, 10) || 1,
      rowGap: Geo.toPt(parseFloat($("mxRowGap").value) || 0, unit),
      colGap: Geo.toPt(parseFloat($("mxColGap").value) || 0, unit),
      sizing: document.querySelector('input[name="mxSize"]:checked').value,
      align: [av, ah],
    };
    await PowerPoint.run(async (context) => {
      const items = await readSelection(context, 2);
      const out = Geo.matrix(items.map(rectOf), opts);
      items.forEach((s, i) => applyRect(s, out[i]));
      await context.sync();
    });
    setStatus("Objek disusun dalam matriks (urut sesuai urutan seleksi dari PowerPoint).");
  },

  /* ---------- Text box ---------- */
  async splitText() {
    await PowerPoint.run(async (context) => {
      const items = await readSelection(context);
      if (items.length !== 1) throw new Error("Pilih tepat 1 text box, taruh kursor di titik pemisah.");
      const src = items[0];
      const sel = context.presentation.getSelectedTextRangeOrNullObject();
      sel.load("isNullObject,start");
      const tr = src.textFrame.textRange;
      tr.load("text");
      tr.font.load("name,size,color,bold,italic");
      const slide = src.getParentSlide();
      await context.sync();
      if (sel.isNullObject) throw new Error("Taruh kursor di dalam teks dulu.");

      const [head, tail] = Geo.splitText(tr.text, sel.start);
      const gap = 8;
      const box = slide.shapes.addTextBox(tail, { left: src.left, top: src.top + src.height + gap, width: src.width, height: src.height });
      box.name = src.name + " (2)";
      const f = box.textFrame.textRange.font;
      if (tr.font.name) f.name = tr.font.name;
      if (tr.font.size) f.size = tr.font.size;
      if (tr.font.color) f.color = tr.font.color;
      if (tr.font.bold !== null) f.bold = tr.font.bold;
      if (tr.font.italic !== null) f.italic = tr.font.italic;
      tr.text = head;
      await context.sync();
    });
    setStatus("Text box dipisah jadi dua.");
  },

  async mergeText() {
    await PowerPoint.run(async (context) => {
      const items = await readSelection(context, 2);
      const texts = items.map((s) => { s.textFrame.textRange.load("text"); return s; });
      await context.sync();
      const merged = texts.map((s) => s.textFrame.textRange.text).join("\n");
      const first = items[0];
      const bottom = Math.max(...items.map((s) => s.top + s.height));
      first.textFrame.textRange.text = merged;
      first.height = Math.max(first.height, bottom - first.top);
      items.slice(1).forEach((s) => s.delete());
      await context.sync();
    });
    setStatus("Text box digabung (urut sesuai urutan seleksi).");
  },

  /* ---------- Rounded rectangle radius ---------- */
  async radius() {
    const mode = document.querySelector('input[name="rrMode"]:checked').value;
    const unit = $("rrUnit").value;
    const force = $("rrForce").checked;
    let n = 0;
    await PowerPoint.run(async (context) => {
      const items = await readSelection(context);
      const mIdx = Geo.pickMaster(items, masterMode());
      const cand = items.filter((s) => TEXT_TYPES.includes(s.type) && (force || /round/i.test(s.name)));
      if (!cand.length) throw new Error('Tidak ada rounded rectangle terpilih (nama shape harus mengandung "Round"). Centang "abaikan cek nama" jika perlu.');

      let radiusPt;
      if (mode === "master") {
        const m = items[mIdx];
        m.adjustments.load("count");
        await context.sync();
        if (!m.adjustments.count) throw new Error("Master tidak punya nilai radius.");
        const a = m.adjustments.get(0);
        await context.sync();
        radiusPt = Geo.adjToRadius(a.value, m.width, m.height);
      } else {
        radiusPt = Geo.toPt(parseFloat($("rrValue").value) || 0, unit);
      }
      cand.forEach((s) => s.adjustments.load("count"));
      await context.sync();
      for (const s of cand) {
        if (!s.adjustments.count) continue;
        s.adjustments.set(0, Geo.radiusToAdj(radiusPt, s.width, s.height));
        n++;
      }
      await context.sync();
    });
    setStatus(`Radius diterapkan ke ${n} shape.`);
  },
};

/* ---------- UI ---------- */

function syncRadiusUI() {
  $("rrValue").disabled = $("rrUnit").disabled = document.querySelector('input[name="rrMode"]:checked').value === "master";
}
function syncMatrixUI() {
  const m = document.querySelector('input[name="mxMode"]:checked').value;
  $("mxRows").disabled = m !== "rows";
  $("mxCols").disabled = m !== "cols";
}

/* ---------- lembar pengaturan ---------- */

function openSheet(id) {
  closeSheets();
  $("scrim").hidden = false;
  $(id).hidden = false;
}
function closeSheets() {
  $("scrim").hidden = true;
  document.querySelectorAll(".sheet").forEach((el) => (el.hidden = true));
}

/* ---------- jumlah objek terpilih ---------- */

async function updateCount() {
  try {
    await PowerPoint.run(async (context) => {
      const sel = context.presentation.getSelectedShapes();
      sel.load("items/id");
      await context.sync();
      const n = sel.items.length;
      $("selCount").textContent = n ? `${n} objek dipilih` : "Pilih objek di slide";
    });
  } catch (e) { /* abaikan: hanya info tambahan */ }
}

Office.onReady((info) => {
  if (info.host !== Office.HostType.PowerPoint) {
    setStatus("Add-in ini hanya untuk PowerPoint.", true);
    return;
  }
  setStatus("Siap. Pilih objek lalu klik tombol.");
  syncRadiusUI(); syncMatrixUI();
  document.querySelectorAll('input[name="rrMode"]').forEach((e) => e.addEventListener("change", syncRadiusUI));
  document.querySelectorAll('input[name="mxMode"]').forEach((e) => e.addEventListener("change", syncMatrixUI));

  updateCount();
  try {
    Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, updateCount);
  } catch (e) { /* versi lama: hitungan tidak diperbarui otomatis */ }

  document.addEventListener("keydown", (ev) => { if (ev.key === "Escape") closeSheets(); });

  document.body.addEventListener("click", async (ev) => {
    const opener = ev.target.closest("[data-open]");
    if (opener) { openSheet(opener.dataset.open); return; }
    if (ev.target.closest("[data-close]") || ev.target === $("scrim")) { closeSheets(); return; }

    const btn = ev.target.closest("button[data-act]");
    if (!btn) return;
    const fn = ACTIONS[btn.dataset.act];
    if (!fn) return;
    btn.disabled = true;
    try {
      await fn(btn.dataset.v, ev.ctrlKey || ev.metaKey);
      if (btn.closest(".sheet")) closeSheets(); // tutup lembar hanya jika berhasil
    } catch (e) {
      setStatus(e.message || String(e), true);
    } finally {
      btn.disabled = false;
    }
  });
});
