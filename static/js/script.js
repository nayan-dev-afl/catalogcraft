/**
 * CatalogCraft — script.js (v3 — Full Redesign)
 *
 * CHANGES:
 * 1. Template text ALWAYS fits fixed box — font scales down + text wraps, no overflow ever.
 * 2. Excel flow: upload → dropdown of all items → select → auto-fills form + triggers live preview.
 * 3. Live preview in right column of Tab 3 — updates every time form changes or item selected.
 * 4. Empty space used: 3-col layout (Excel | Form | Live Preview), entries panel in Tab 4.
 * 5. Output tab shows entry list + filename alongside page thumbnails.
 */

/* ── STATE ─────────────────────────────────────────────────────────────────── */
const STATE = {
  productImages: [], logoDataURL: null,
  logoPos: { x: 50, y: 50 }, logoSize: 20, logoOpacity: 100,
  appliedImages: [], templates: { 1: null, 2: null, 3: null },
  excelRows: [], excelHeaders: [],
  savedEntries: [], editingEntryId: null,
  livePreviewTimer: null,
};

/* ── CATEGORY FIELDS ───────────────────────────────────────────────────────── */
const CATEGORY_FIELDS = {
  'SAREE': [
    { key: 'productName',  label: 'Name of Product *',     placeholder: 'e.g. RT-FASHION FAIR', type: 'text' },
    { key: 'rate',         label: 'Rate (Price) *',        placeholder: 'e.g. 1200',            type: 'text' },
    { key: 'pieces',       label: 'Number of Pieces',      placeholder: 'e.g. 05',              type: 'text' },
    { key: 'length',       label: 'Product Length',        placeholder: 'e.g. 6.30 MTR APX',    type: 'text' },
    { key: 'fabric',       label: 'Product Fabric',        placeholder: 'e.g. P*C MOSS',        type: 'text' },
    { key: 'work',         label: 'Product Work',          placeholder: 'e.g. EMBROIDERY',       type: 'text' },
    { key: 'salesPackage', label: 'Product Sales Package', placeholder: 'e.g. SAREE WITH UNSTITCHED BLOUSE', type: 'text', span: 2 },
    { key: 'packing',      label: 'Packing Style',         placeholder: 'e.g. POUCH',            type: 'text' },
  ],
  'LEHENGA': [
    { key: 'productName',  label: 'Name of Product *',     placeholder: 'e.g. PBG-35309',       type: 'text' },
    { key: 'rate',         label: 'Rate (Price) *',        placeholder: 'e.g. 3500',            type: 'text' },
    { key: 'pieces',       label: 'Number of Pieces',      placeholder: 'e.g. 01',              type: 'text' },
    { key: 'fabric',       label: 'Product Fabric',        placeholder: 'e.g. VELVET',           type: 'text' },
    { key: 'work',         label: 'Product Pattern',       placeholder: 'e.g. EMBROIDERY',       type: 'text' },
    { key: 'length',       label: 'Product Style',         placeholder: 'e.g. LEHENGA',          type: 'text' },
    { key: 'packing',      label: 'Packing Style',         placeholder: 'e.g. BAG',              type: 'text' },
    { key: 'salesPackage', label: 'Product Sales Package', placeholder: 'e.g. UNSTITCHED BLOUSE WITH LEHENGA AND DUPATTA', type: 'text', span: 2 },
  ],
  'SUIT': [
    { key: 'productName',  label: 'Name of Product *', placeholder: 'e.g. YF-LOTUS CRUNCHY', type: 'text' },
    { key: 'rate',         label: 'Rate (Price) *',    placeholder: 'e.g. 1500',             type: 'text' },
    { key: 'pieces',       label: 'Number of Pieces',  placeholder: 'e.g. 04',               type: 'text' },
    { key: 'work',         label: 'Product Type',      placeholder: 'e.g. UNSTITCHED',       type: 'text' },
    { key: 'fabric',       label: 'Top Fabric',        placeholder: 'e.g. CRUNCHI',          type: 'text' },
    { key: 'extraFabric1', label: 'Bottom Fabric',     placeholder: 'e.g. CRUNCHI',          type: 'text' },
    { key: 'extraFabric2', label: 'Dupatta Fabric',    placeholder: 'e.g. CRUNCHI',          type: 'text' },
    { key: 'length',       label: 'Top Length',        placeholder: 'e.g. 2.50 MTR APX',    type: 'text' },
    { key: 'extraLen1',    label: 'Bottom Length',     placeholder: 'e.g. 2.50 MTR APX',    type: 'text' },
    { key: 'extraLen2',    label: 'Dupatta Length',    placeholder: 'e.g. 2.25 MTR APX',    type: 'text' },
    { key: 'salesPackage', label: 'Sales Package',     placeholder: 'e.g. UNSTITCHED TOP WITH BOTTOM & DUPATTA', type: 'text', span: 2 },
    { key: 'packing',      label: 'Packing Style',     placeholder: 'e.g. POUCH',            type: 'text' },
  ],
  'KURTI': [
    { key: 'productName', label: 'Name of Product *', placeholder: 'e.g. KURTI-101',   type: 'text' },
    { key: 'rate',        label: 'Rate (Price) *',    placeholder: 'e.g. 450',         type: 'text' },
    { key: 'pieces',      label: 'Number of Pieces',  placeholder: 'e.g. 06',          type: 'text' },
    { key: 'fabric',      label: 'Product Fabric',    placeholder: 'e.g. RAYON',       type: 'text' },
    { key: 'length',      label: 'Product Length',    placeholder: 'e.g. 42 INCH',     type: 'text' },
    { key: 'work',        label: 'Product Work',      placeholder: 'e.g. PRINT',       type: 'text' },
    { key: 'length2',     label: 'Product Size',      placeholder: 'e.g. M-L-XL-XXL', type: 'text' },
    { key: 'packing',     label: 'Packing Style',     placeholder: 'e.g. POLYTHENE',   type: 'text' },
  ],
  'DRESS MATERIAL': [
    { key: 'productName',  label: 'Name of Product *', placeholder: 'e.g. DM-101',          type: 'text' },
    { key: 'rate',         label: 'Rate (Price) *',    placeholder: 'e.g. 800',             type: 'text' },
    { key: 'pieces',       label: 'Number of Pieces',  placeholder: 'e.g. 06',              type: 'text' },
    { key: 'fabric',       label: 'Product Fabric',    placeholder: 'e.g. COTTON',          type: 'text' },
    { key: 'length',       label: 'Product Length',    placeholder: 'e.g. 4.50 MTR APX',   type: 'text' },
    { key: 'work',         label: 'Product Work',      placeholder: 'e.g. PRINT',           type: 'text' },
    { key: 'salesPackage', label: 'Sales Package',     placeholder: 'e.g. TOP WITH BOTTOM', type: 'text', span: 2 },
    { key: 'packing',      label: 'Packing Style',     placeholder: 'e.g. POLYTHENE',       type: 'text' },
  ],
  'DUPATTA': [
    { key: 'productName', label: 'Name of Product *', placeholder: 'e.g. RRC-SCARF-2',   type: 'text' },
    { key: 'rate',        label: 'Rate (Price) *',    placeholder: 'e.g. 150',           type: 'text' },
    { key: 'pieces',      label: 'Number of Pieces',  placeholder: 'e.g. 15',            type: 'text' },
    { key: 'fabric',      label: 'Product Fabric',    placeholder: 'e.g. COTTON',        type: 'text' },
    { key: 'length',      label: 'Product Length',    placeholder: 'e.g. 2.00 MTR APX', type: 'text' },
    { key: 'packing',     label: 'Packing Style',     placeholder: 'e.g. ZIP BAG',       type: 'text' },
  ],
  'BED SHEET': [
    { key: 'productName',  label: 'Name of Product *', placeholder: 'e.g. DRF-DAISY (1+1)',           type: 'text' },
    { key: 'rate',         label: 'Rate (Price) *',    placeholder: 'e.g. 650',                       type: 'text' },
    { key: 'pieces',       label: 'Number of Pieces',  placeholder: 'e.g. 12',                        type: 'text' },
    { key: 'fabric',       label: 'Product Fabric',    placeholder: 'e.g. GLACE COTTON',              type: 'text' },
    { key: 'length',       label: 'Bed Sheet Length',  placeholder: 'e.g. 60 X 90 CM',               type: 'text' },
    { key: 'length2',      label: 'Pillow Length',     placeholder: 'e.g. 43 X 61 CM',               type: 'text' },
    { key: 'packing',      label: 'Packing Style',     placeholder: 'e.g. BED SHEET BAG',             type: 'text' },
    { key: 'salesPackage', label: 'Sales Package',     placeholder: 'e.g. BED SHEET WITH PILLOW COVER', type: 'text', span: 2 },
  ],
  'CURTAIN': [
    { key: 'productName', label: 'Name of Product *', placeholder: 'e.g. SMF-BAHURANI', type: 'text' },
    { key: 'rate',        label: 'Rate (Price) *',    placeholder: 'e.g. 350',          type: 'text' },
    { key: 'pieces',      label: 'Number of Pieces',  placeholder: 'e.g. 04',           type: 'text' },
    { key: 'fabric',      label: 'Product Fabric',    placeholder: 'e.g. NET',          type: 'text' },
    { key: 'length',      label: 'Product Length',    placeholder: 'e.g. 4*7',          type: 'text' },
    { key: 'packing',     label: 'Packing Style',     placeholder: 'e.g. POLYTHENE',    type: 'text' },
    { key: 'notes',       label: 'Notes',             placeholder: 'Any extra info...', type: 'textarea', span: 2 },
  ],
  'GOWN': [
    { key: 'productName',  label: 'Name of Product *', placeholder: 'e.g. SHP-RAGINI',            type: 'text' },
    { key: 'rate',         label: 'Rate (Price) *',    placeholder: 'e.g. 1800',                  type: 'text' },
    { key: 'pieces',       label: 'Number of Pieces',  placeholder: 'e.g. 03',                    type: 'text' },
    { key: 'fabric',       label: 'Product Fabric',    placeholder: 'e.g. RANGOLI',               type: 'text' },
    { key: 'work',         label: 'Product Pattern',   placeholder: 'e.g. SEQUENCE',              type: 'text' },
    { key: 'length',       label: 'Product Style',     placeholder: 'e.g. GOWN',                  type: 'text' },
    { key: 'salesPackage', label: 'Sales Package',     placeholder: 'e.g. LONG GOWN WITH DUPATTA', type: 'text', span: 2 },
    { key: 'packing',      label: 'Packing Style',     placeholder: 'e.g. POLYTHENE',             type: 'text' },
  ],
  'KIDS WEAR': [
    { key: 'productName', label: 'Name of Product *', placeholder: 'e.g. LAM-1013',  type: 'text' },
    { key: 'rate',        label: 'Rate (Price) *',    placeholder: 'e.g. 400',        type: 'text' },
    { key: 'pieces',      label: 'Number of Pieces',  placeholder: 'e.g. 06',         type: 'text' },
    { key: 'length2',     label: 'Product Size',      placeholder: 'e.g. 22-36',      type: 'text' },
    { key: 'packing',     label: 'Packing Style',     placeholder: 'e.g. POLYTHENE',  type: 'text' },
    { key: 'notes',       label: 'Notes',             placeholder: 'e.g. GIRLS',      type: 'text' },
  ],
  'MENS WEAR': [
    { key: 'productName', label: 'Name of Product *', placeholder: 'e.g. SNA-MSU13702', type: 'text' },
    { key: 'rate',        label: 'Rate (Price) *',    placeholder: 'e.g. 1100',          type: 'text' },
    { key: 'pieces',      label: 'Number of Pieces',  placeholder: 'e.g. 01',            type: 'text' },
    { key: 'fabric',      label: 'Product Fabric',    placeholder: 'e.g. IMPORTED',      type: 'text' },
    { key: 'length2',     label: 'Product Size',      placeholder: 'e.g. 36-38',         type: 'text' },
    { key: 'packing',     label: 'Packing Style',     placeholder: 'e.g. BAG',           type: 'text' },
    { key: 'notes',       label: 'Notes',             placeholder: 'Any notes...',        type: 'text' },
  ],
  'NIGHTY': [
    { key: 'productName', label: 'Name of Product *', placeholder: 'e.g. SNW-GUJRI EMBROIDERY', type: 'text' },
    { key: 'rate',        label: 'Rate (Price) *',    placeholder: 'e.g. 300',                  type: 'text' },
    { key: 'pieces',      label: 'Number of Pieces',  placeholder: 'e.g. 03',                   type: 'text' },
    { key: 'fabric',      label: 'Product Fabric',    placeholder: 'e.g. COTTON',               type: 'text' },
    { key: 'length2',     label: 'Product Size',      placeholder: 'e.g. XX-XXL',               type: 'text' },
    { key: 'packing',     label: 'Packing Style',     placeholder: 'e.g. POLYTHENE',             type: 'text' },
  ],
  'READYMADE BLOUSE': [
    { key: 'productName', label: 'Name of Product *', placeholder: 'e.g. NSA-SHAPEWEAR',  type: 'text' },
    { key: 'rate',        label: 'Rate (Price) *',    placeholder: 'e.g. 250',            type: 'text' },
    { key: 'pieces',      label: 'Number of Pieces',  placeholder: 'e.g. 10',             type: 'text' },
    { key: 'fabric',      label: 'Product Fabric',    placeholder: 'e.g. COTTON',         type: 'text' },
    { key: 'length2',     label: 'Product Size',      placeholder: 'e.g. S-M-L-XL',      type: 'text' },
    { key: 'length',      label: 'Product Length',    placeholder: 'e.g. 2.00 MTR APX',  type: 'text' },
    { key: 'packing',     label: 'Packing Style',     placeholder: 'e.g. ZIP BAG',        type: 'text' },
  ],
  'SHIRTING SUITING': [
    { key: 'productName',  label: 'Name of Product *', placeholder: 'e.g. MFP-AYUSH',     type: 'text' },
    { key: 'rate',         label: 'Rate (Price) *',    placeholder: 'e.g. 900',           type: 'text' },
    { key: 'pieces',       label: 'Number of Pieces',  placeholder: 'e.g. 10',            type: 'text' },
    { key: 'fabric',       label: 'Product Fabric',    placeholder: 'e.g. POLYESTER',     type: 'text' },
    { key: 'length',       label: 'Pant Length',       placeholder: 'e.g. 2.20 MTR APX', type: 'text' },
    { key: 'length2',      label: 'Shirt Length',      placeholder: 'e.g. 1.20 MTR APX', type: 'text' },
    { key: 'salesPackage', label: 'Notes',             placeholder: 'e.g. SHIRT+PENT',   type: 'text' },
    { key: 'packing',      label: 'Packing Style',     placeholder: 'e.g. BOX',           type: 'text' },
  ],
  'GIRLS TOP AND BOTTOM': [
    { key: 'productName', label: 'Name of Product *', placeholder: 'e.g. DIM-BAGGI WATER',  type: 'text' },
    { key: 'rate',        label: 'Rate (Price) *',    placeholder: 'e.g. 700',              type: 'text' },
    { key: 'pieces',      label: 'Number of Pieces',  placeholder: 'e.g. 05',               type: 'text' },
    { key: 'fabric',      label: 'Product Fabric',    placeholder: 'e.g. LYCRA',             type: 'text' },
    { key: 'length2',     label: 'Product Size',      placeholder: 'e.g. 28-30-32-34-36',  type: 'text' },
    { key: 'packing',     label: 'Packing Style',     placeholder: 'e.g. POLYTHENE',        type: 'text' },
    { key: 'notes',       label: 'Notes',             placeholder: 'Any extra info...',      type: 'text' },
  ],
  'OTHER': [
    { key: 'productName',  label: 'Name of Product *', placeholder: 'Product name',      type: 'text' },
    { key: 'rate',         label: 'Rate (Price) *',    placeholder: 'e.g. 500',          type: 'text' },
    { key: 'pieces',       label: 'Number of Pieces',  placeholder: 'e.g. 06',           type: 'text' },
    { key: 'fabric',       label: 'Product Fabric',    placeholder: 'Fabric type',        type: 'text' },
    { key: 'length',       label: 'Product Length',    placeholder: 'Length',             type: 'text' },
    { key: 'work',         label: 'Product Work',      placeholder: 'Work type',          type: 'text' },
    { key: 'salesPackage', label: 'Sales Package',     placeholder: 'Sales package',      type: 'text', span: 2 },
    { key: 'packing',      label: 'Packing Style',     placeholder: 'Packing',            type: 'text' },
    { key: 'notes',        label: 'Notes',             placeholder: 'Any extra info...',  type: 'textarea', span: 2 },
  ],
};

/* ── EXCEL COLUMN FUZZY RULES ─────────────────────────────────────────────── */
const EXCEL_COL_RULES = [
  ['product name','productName'],['item name','productName'],['name of product','productName'],
  ['no of pieces','pieces'],['num of pieces','pieces'],['number of pieces','pieces'],['no. of pieces','pieces'],
  ['product sales package','salesPackage'],['sales package','salesPackage'],
  ['packing style','packing'],['product pattern','work'],['product type','work'],['product work','work'],
  ['product style','length'],['product fabric','fabric'],['top fabric','fabric'],
  ['bottom fabric','extraFabric1'],['dupatta fabric','extraFabric2'],
  ['bed sheet length','length'],['pillow length','length2'],['top length','length'],
  ['bottom length','extraLen1'],['dupatta length','extraLen2'],
  ['pant length','length'],['shirt length','length2'],['product length','length'],['product size','length2'],
  ['category','category'],['product','productName'],['item','productName'],['name','productName'],
  ['pieces','pieces'],['pcs','pieces'],['package','salesPackage'],['packing','packing'],
  ['length','length'],['size','length2'],['fabric','fabric'],['pattern','work'],['work','work'],
  ['style','length'],['notes','notes'],['note','notes'],
  ['rate','rate'],['price','rate'],['mrp','rate'],['amount','rate'],
];

function normaliseHeader(h) {
  const n = h.toString().trim().toLowerCase();
  for (const [sub, key] of EXCEL_COL_RULES) if (n.includes(sub)) return key;
  return null;
}

/* ── INIT ─────────────────────────────────────────────────────────────────── */
document.addEventListener('DOMContentLoaded', async () => {
  initTabs(); initMedia(); initTemplates(); initSpecs(); initOutput(); initModals();
  await loadSession();
  try { await loadDefaultAssets(); setupAssetToggles(); } catch (e) { console.warn('Defaults skipped'); }
});

/* ── SESSION ──────────────────────────────────────────────────────────────── */
async function saveSession() {
  const d = { savedEntries: STATE.savedEntries, logoSize: STATE.logoSize, logoOpacity: STATE.logoOpacity, logoPos: STATE.logoPos };
  try { await fetch('/api/session', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(d) }); } catch (_) {}
  try { localStorage.setItem('catalogcraft_session', JSON.stringify(d)); } catch (_) {}
}
async function loadSession() {
  let data = null;
  try { const r = await fetch('/api/session'); if (r.ok) { const j = await r.json(); if (j && Object.keys(j).length) data = j; } } catch (_) {}
  if (!data) { try { const r = localStorage.getItem('catalogcraft_session'); if (r) data = JSON.parse(r); } catch (_) {} }
  if (!data) return;
  if (data.savedEntries) STATE.savedEntries = data.savedEntries;
  if (data.logoSize)     STATE.logoSize     = data.logoSize;
  if (data.logoOpacity)  STATE.logoOpacity  = data.logoOpacity;
  if (data.logoPos)      STATE.logoPos      = data.logoPos;
  if (STATE.savedEntries.length) renderSavedEntries();
}
document.getElementById('btn-reset-session').addEventListener('click', async () => {
  if (!confirm('Reset all data?')) return;
  try { await fetch('/api/session/reset', { method: 'POST' }); } catch (_) {}
  localStorage.removeItem('catalogcraft_session');
  location.reload();
});

/* ── TABS ─────────────────────────────────────────────────────────────────── */
function initTabs() {
  document.querySelectorAll('.tab-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
      document.querySelectorAll('.tab-panel').forEach(p => p.classList.remove('active'));
      btn.classList.add('active');
      const panel = document.getElementById(btn.dataset.tab);
      if (panel) panel.classList.add('active');
      if (btn.dataset.tab === 'tab-output') { refreshOutputSummary(); renderOutputEntries(); }
    });
  });
}

/* ── TOAST ────────────────────────────────────────────────────────────────── */
function toast(msg, type = 'info', duration = 3200) {
  const el = document.createElement('div');
  el.className = `toast ${type}`;
  el.innerHTML = `<span>${{success:'✓',error:'✕',warning:'⚠',info:'ℹ'}[type]||'ℹ'}</span> ${msg}`;
  document.getElementById('toast-container').appendChild(el);
  setTimeout(() => el.remove(), duration);
}

/* ── MODALS ───────────────────────────────────────────────────────────────── */
function openLightbox(src) { document.getElementById('lightbox-img').src = src; document.getElementById('lightbox-overlay').style.display = 'flex'; }
function initModals() {
  document.getElementById('lightbox-close').onclick = () => document.getElementById('lightbox-overlay').style.display = 'none';
  document.getElementById('lightbox-overlay').addEventListener('click', e => { if (e.target.id === 'lightbox-overlay') e.target.style.display = 'none'; });
  document.getElementById('edit-modal-close').onclick  = closeEditModal;
  document.getElementById('edit-modal-cancel').onclick = closeEditModal;
  document.getElementById('edit-modal-overlay').addEventListener('click', e => { if (e.target.id === 'edit-modal-overlay') closeEditModal(); });
  document.getElementById('edit-modal-save').onclick = saveEditModal;
}

/* ── DROP ZONE ────────────────────────────────────────────────────────────── */
function setupDropZone(dzId, inputId, cb) {
  const dz = document.getElementById(dzId); if (!dz) return;
  dz.addEventListener('dragover', e => { e.preventDefault(); dz.classList.add('dragover'); });
  dz.addEventListener('dragleave', () => dz.classList.remove('dragover'));
  dz.addEventListener('drop', e => { e.preventDefault(); dz.classList.remove('dragover'); cb(e.dataTransfer.files); });
  dz.addEventListener('click', e => { if (e.target.tagName !== 'LABEL' && e.target.tagName !== 'INPUT') document.getElementById(inputId)?.click(); });
}

/* ── TAB 1: MEDIA ─────────────────────────────────────────────────────────── */
function initMedia() {
  setupDropZone('dz-products', 'input-products', handleProductFiles);
  document.getElementById('input-products').addEventListener('change', e => handleProductFiles(e.target.files));
  document.getElementById('btn-clear-products').addEventListener('click', clearProducts);
  setupDropZone('dz-logo', 'input-logo', f => handleLogoFile(f[0]));
  document.getElementById('input-logo').addEventListener('change', e => handleLogoFile(e.target.files[0]));
  document.getElementById('btn-remove-logo').addEventListener('click', removeLogo);
  document.getElementById('logo-size').addEventListener('input', e => { STATE.logoSize = +e.target.value; document.getElementById('logo-size-val').textContent = STATE.logoSize; updateLogoOnCanvas(); });
  document.getElementById('logo-opacity').addEventListener('input', e => { STATE.logoOpacity = +e.target.value; document.getElementById('logo-opacity-val').textContent = STATE.logoOpacity; updateLogoOnCanvas(); });
  document.getElementById('position-grid').querySelectorAll('button').forEach(btn => {
    btn.addEventListener('click', () => { document.querySelectorAll('#position-grid button').forEach(b => b.classList.remove('active')); btn.classList.add('active'); applyPositionPreset(btn.dataset.pos); });
  });
  document.getElementById('btn-apply-logo-all').addEventListener('click', applyLogoToAll);
}
function handleProductFiles(files) {
  Array.from(files).forEach(file => { if (!file.type.startsWith('image/')) return; const r = new FileReader(); r.onload = e => { STATE.productImages.push({ id: Date.now() + Math.random(), dataURL: e.target.result, name: file.name }); renderProductGrid(); showCanvasPreview(); }; r.readAsDataURL(file); });
}
function renderProductGrid() {
  const grid = document.getElementById('product-grid'), bar = document.getElementById('product-count-bar');
  grid.innerHTML = '';
  STATE.productImages.forEach((img, idx) => {
    const wrap = document.createElement('div'); wrap.className = 'image-thumb-wrap';
    wrap.innerHTML = `<img src="${img.dataURL}" alt="${img.name}"/><div class="thumb-name">${img.name}</div><button class="thumb-remove" data-idx="${idx}">✕</button>`;
    wrap.querySelector('img').addEventListener('click', () => openLightbox(img.dataURL));
    wrap.querySelector('.thumb-remove').addEventListener('click', e => { e.stopPropagation(); STATE.productImages.splice(idx,1); renderProductGrid(); });
    grid.appendChild(wrap);
  });
  bar.style.display = STATE.productImages.length ? 'flex' : 'none';
  document.getElementById('product-count-label').textContent = `${STATE.productImages.length} image(s) loaded`;
}
function clearProducts() { STATE.productImages = []; STATE.appliedImages = []; renderProductGrid(); document.getElementById('applied-preview-card').style.display = 'none'; }
function handleLogoFile(file) {
  if (!file || !file.type.startsWith('image/')) { toast('Please upload an image file.', 'warning'); return; }
  const r = new FileReader();
  r.onload = e => {
    STATE.logoDataURL = e.target.result;
    document.getElementById('logo-preview-thumb').src = e.target.result;
    document.getElementById('dz-logo').style.display = 'none';
    document.getElementById('logo-controls').style.display = 'block';
    document.getElementById('logo-size').value = STATE.logoSize;
    document.getElementById('logo-opacity').value = STATE.logoOpacity;
    showCanvasPreview(); toast('Logo uploaded!', 'success');
  };
  r.readAsDataURL(file);
}
function removeLogo() { STATE.logoDataURL = null; document.getElementById('logo-preview-thumb').src = ''; document.getElementById('dz-logo').style.display = 'block'; document.getElementById('logo-controls').style.display = 'none'; document.getElementById('logo-canvas-area').style.display = 'none'; }
function showCanvasPreview() {
  if (!STATE.logoDataURL || !STATE.productImages.length) return;
  document.getElementById('logo-canvas-area').style.display = 'block';
  const bg = document.getElementById('canvas-bg-image'), logo = document.getElementById('canvas-logo');
  bg.src = STATE.productImages[0].dataURL; logo.src = STATE.logoDataURL;
  bg.onload = () => { positionLogoOnCanvas(); initLogoDrag(document.getElementById('canvas-container'), logo, bg); };
  if (bg.complete) { positionLogoOnCanvas(); initLogoDrag(document.getElementById('canvas-container'), logo, bg); }
}
function positionLogoOnCanvas() {
  const bg = document.getElementById('canvas-bg-image'), logo = document.getElementById('canvas-logo');
  if (!bg.offsetWidth) return;
  const lw = bg.offsetWidth * (STATE.logoSize / 100);
  logo.style.width = lw + 'px'; logo.style.opacity = STATE.logoOpacity / 100;
  logo.style.left = ((STATE.logoPos.x/100)*bg.offsetWidth - lw/2) + 'px';
  logo.style.top  = ((STATE.logoPos.y/100)*bg.offsetHeight - (logo.offsetHeight||lw*0.6)/2) + 'px';
}
function updateLogoOnCanvas() { positionLogoOnCanvas(); }
function applyPositionPreset(pos) {
  const p = {'top-left':{x:12,y:12},'top-center':{x:50,y:12},'top-right':{x:88,y:12},'mid-left':{x:12,y:50},'center':{x:50,y:50},'mid-right':{x:88,y:50},'bottom-left':{x:12,y:88},'bottom-center':{x:50,y:88},'bottom-right':{x:88,y:88}};
  if (p[pos]) { STATE.logoPos={...p[pos]}; positionLogoOnCanvas(); }
}
function initLogoDrag(container, logoEl, bgImg) {
  let d=false,sx,sy,ol,ot;
  const dn=e=>{d=true;const pt=e.touches?e.touches[0]:e;sx=pt.clientX;sy=pt.clientY;ol=logoEl.offsetLeft;ot=logoEl.offsetTop;logoEl.style.cursor='grabbing';e.preventDefault();};
  const mv=e=>{if(!d)return;const pt=e.touches?e.touches[0]:e;const dx=pt.clientX-sx,dy=pt.clientY-sy;const cw=bgImg.offsetWidth,ch=bgImg.offsetHeight,lw=logoEl.offsetWidth,lh=logoEl.offsetHeight;const nl=Math.max(0,Math.min(cw-lw,ol+dx)),nt=Math.max(0,Math.min(ch-lh,ot+dy));logoEl.style.left=nl+'px';logoEl.style.top=nt+'px';STATE.logoPos.x=((nl+lw/2)/cw)*100;STATE.logoPos.y=((nt+lh/2)/ch)*100;e.preventDefault();};
  const up=()=>{d=false;logoEl.style.cursor='grab';saveSession();};
  logoEl.removeEventListener('mousedown',logoEl._md);logoEl._md=dn;logoEl.addEventListener('mousedown',dn);logoEl.addEventListener('touchstart',dn,{passive:false});
  document.removeEventListener('mousemove',document._mm);document._mm=mv;document.addEventListener('mousemove',mv);document.addEventListener('touchmove',mv,{passive:false});
  document.removeEventListener('mouseup',document._mu);document._mu=up;document.addEventListener('mouseup',up);document.addEventListener('touchend',up);
}
async function applyLogoToAll() {
  if(!STATE.logoDataURL){toast('Upload a logo first.','warning');return;}
  if(!STATE.productImages.length){toast('Upload product images first.','warning');return;}
  const btn=document.getElementById('btn-apply-logo-all');btn.textContent='⏳ Processing...';btn.disabled=true;
  STATE.appliedImages=[];
  const oc=document.getElementById('offscreen-canvas'),ctx=oc.getContext('2d');
  const li=new Image();li.src=STATE.logoDataURL;await new Promise(r=>{li.onload=r;if(li.complete)r();});
  for(const prod of STATE.productImages){
    const bg=new Image();bg.src=prod.dataURL;await new Promise(r=>{bg.onload=r;if(bg.complete)r();});
    oc.width=bg.naturalWidth;oc.height=bg.naturalHeight;ctx.clearRect(0,0,oc.width,oc.height);ctx.drawImage(bg,0,0);
    const lw=oc.width*(STATE.logoSize/100),lh=li.naturalHeight*(lw/li.naturalWidth);
    const lx=(STATE.logoPos.x/100)*oc.width-lw/2,ly=(STATE.logoPos.y/100)*oc.height-lh/2;
    ctx.globalAlpha=STATE.logoOpacity/100;ctx.drawImage(li,Math.max(0,lx),Math.max(0,ly),lw,lh);ctx.globalAlpha=1;
    STATE.appliedImages.push({dataURL:oc.toDataURL('image/jpeg',0.95),name:prod.name});
  }
  const card=document.getElementById('applied-preview-card'),grid=document.getElementById('applied-grid');
  card.style.display='block';grid.innerHTML='';
  STATE.appliedImages.forEach(img=>{const w=document.createElement('div');w.className='image-thumb-wrap';w.innerHTML=`<img src="${img.dataURL}"/>`;w.querySelector('img').addEventListener('click',()=>openLightbox(img.dataURL));grid.appendChild(w);});
  btn.textContent='Apply Logo to All Images →';btn.disabled=false;
  toast(`Logo applied to ${STATE.appliedImages.length} image(s)!`,'success');
}

/* ── TAB 2: TEMPLATES ─────────────────────────────────────────────────────── */
function initTemplates(){
  [1,2,3].forEach(n=>{
    setupDropZone(`dz-tmpl${n}`,`input-tmpl${n}`,f=>handleTemplateFile(n,f));
    document.getElementById(`input-tmpl${n}`).addEventListener('change',e=>handleTemplateFile(n,e.target.files));
  });
  document.querySelectorAll('.tmpl-remove').forEach(btn=>btn.addEventListener('click',()=>removeTemplate(+btn.dataset.tmpl)));
  document.getElementById('btn-preview-tmpl1').addEventListener('click',previewTemplate1);
}
function handleTemplateFile(n,files){
  const file=Array.from(files).find(f=>f.type.startsWith('image/'));
  if(!file){toast('Please upload an image file.','warning');return;}
  const r=new FileReader();
  r.onload=e=>{STATE.templates[n]=e.target.result;document.getElementById(`tmpl${n}-img`).src=e.target.result;document.getElementById(`tmpl${n}-preview-wrap`).style.display='block';document.getElementById(`dz-tmpl${n}`).style.display='none';toast(`Template ${n} uploaded!`,'success');};
  r.readAsDataURL(file);
}
function removeTemplate(n){STATE.templates[n]=null;document.getElementById(`tmpl${n}-preview-wrap`).style.display='none';document.getElementById(`dz-tmpl${n}`).style.display='block';if(n===1)document.getElementById('tmpl1-full-preview-card').style.display='none';}
async function previewTemplate1(){
  if(!STATE.templates[1]){toast('Upload Template 1 first.','warning');return;}
  const entry=STATE.savedEntries.length?STATE.savedEntries[STATE.savedEntries.length-1]:getFormData();
  const btn=document.getElementById('btn-preview-tmpl1');btn.disabled=true;btn.textContent='⏳ Rendering...';
  try{
    const dataURL=await compositeTemplate1(entry);
    const card=document.getElementById('tmpl1-full-preview-card'),cont=document.getElementById('tmpl1-canvas-container');
    card.style.display='block';cont.innerHTML=`<img src="${dataURL}" style="max-width:100%;border-radius:8px;border:1px solid var(--clr-border)"/>`;
    card.scrollIntoView({behavior:'smooth',block:'start'});toast('Preview ready!','success');
  }catch(e){toast('Preview failed.','error');}
  finally{btn.disabled=false;btn.textContent='👁 Preview with Data';}
}

/* ── TEMPLATE 1 COMPOSITE — FIXED BOX, TEXT AUTO-FITS ────────────────────────
   The white product-details box has FIXED dimensions matching your template.
   All text wraps inside that box. Font scales down if needed. Nothing overflows.
────────────────────────────────────────────────────────────────────────────── */
async function compositeTemplate1(entry) {
  return new Promise(resolve => {
    const bgImg = new Image();
    bgImg.crossOrigin = 'Anonymous';
    bgImg.onload = () => {
      const W = bgImg.naturalWidth || 1816;
      const H = bgImg.naturalHeight || 2568;
      const canvas = document.createElement('canvas');
      canvas.width = W; canvas.height = H;
      const ctx = canvas.getContext('2d');
      ctx.drawImage(bgImg, 0, 0, W, H);

      const rows = getEntryDisplayFields(entry);
      if (!rows.length) { resolve(canvas.toDataURL('image/jpeg', 0.95)); return; }

      /* FIXED box that matches the template's white product-details area */
      const BOX_X  = W * 0.04;
      const BOX_Y  = H * 0.298;
      const BOX_W  = W * 0.92;
      const BOX_H  = H * 0.29;   // FIXED — text must fit inside this height

      const PAD    = W * 0.04;
      const labelX = BOX_X + PAD;
      const valueX = BOX_X + BOX_W * 0.47;
      const valueMaxW = BOX_X + BOX_W - valueX - PAD * 0.5;

      /* Helper: wrap text into lines */
      function wrapLines(text, font, maxW) {
        ctx.font = font;
        const words = text.split(' ');
        const lines = []; let cur = '';
        for (const w of words) {
          const t = cur ? cur + ' ' + w : w;
          if (ctx.measureText(t).width <= maxW) cur = t;
          else { if (cur) lines.push(cur); cur = w; }
        }
        if (cur) lines.push(cur);
        return lines.length ? lines : [text];
      }

      /* Dynamically find the largest font that makes ALL rows fit in BOX_H */
      let fontSize  = W * 0.023;
      const MIN_FS  = W * 0.010;
      const LINE_SP = 1.35;

      function calcTotalH(fs) {
        let total = 0;
        const rowPad = fs * 0.55;
        rows.forEach(row => {
          const lines = wrapLines(row.value, `bold ${fs}px "DM Sans",Arial,sans-serif`, valueMaxW);
          total += lines.length * fs * LINE_SP + rowPad;
        });
        return total;
      }

      /* Scale font down until content fits */
      while (fontSize > MIN_FS && calcTotalH(fontSize) > BOX_H * 0.92) {
        fontSize -= 0.5;
      }

      const rowPad = fontSize * 0.55;

      /* Draw white box */
      ctx.save(); ctx.globalAlpha = 0.93; ctx.fillStyle = '#ffffff';
      roundRect(ctx, BOX_X, BOX_Y, BOX_W, BOX_H, W * 0.012); ctx.fill();
      ctx.globalAlpha = 1; ctx.restore();

      /* Gold border */
      ctx.strokeStyle = '#c9a84c'; ctx.lineWidth = W * 0.003;
      roundRect(ctx, BOX_X, BOX_Y, BOX_W, BOX_H, W * 0.012); ctx.stroke();

      /* Clip future drawing to the box so nothing can escape */
      ctx.save();
      ctx.beginPath();
      roundRect(ctx, BOX_X + 2, BOX_Y + 2, BOX_W - 4, BOX_H - 4, W * 0.011);
      ctx.clip();

      /* Draw rows */
      let curY   = BOX_Y + fontSize * 0.8;
      const lfs  = fontSize * 0.85;

      rows.forEach((row, i) => {
        const valLines = wrapLines(row.value, `bold ${fontSize}px "DM Sans",Arial,sans-serif`, valueMaxW);
        const rowH     = valLines.length * fontSize * LINE_SP + rowPad;

        /* Alternating tint */
        if (i % 2 === 0) {
          ctx.save(); ctx.globalAlpha = 0.22; ctx.fillStyle = '#f3ede4';
          ctx.fillRect(BOX_X + 4, curY - fontSize * 0.4, BOX_W - 8, rowH);
          ctx.restore();
        }

        /* Label */
        ctx.fillStyle = '#7a1a2e';
        ctx.font = `600 ${lfs}px "DM Sans",Arial,sans-serif`;
        ctx.textAlign = 'left'; ctx.textBaseline = 'top';
        ctx.fillText(row.label + ' :-', labelX, curY);

        /* Value lines */
        ctx.fillStyle = '#1e1410';
        ctx.font = `bold ${fontSize}px "DM Sans",Arial,sans-serif`;
        valLines.forEach((line, li) => ctx.fillText(line, valueX, curY + li * fontSize * LINE_SP));

        curY += rowH;
      });

      ctx.restore(); /* remove clip */
      resolve(canvas.toDataURL('image/jpeg', 0.95));
    };
    bgImg.onerror = () => resolve(STATE.templates[1]);
    bgImg.src = STATE.templates[1];
  });
}

function roundRect(ctx, x, y, w, h, r) {
  ctx.beginPath();
  ctx.moveTo(x+r,y);ctx.lineTo(x+w-r,y);ctx.quadraticCurveTo(x+w,y,x+w,y+r);
  ctx.lineTo(x+w,y+h-r);ctx.quadraticCurveTo(x+w,y+h,x+w-r,y+h);
  ctx.lineTo(x+r,y+h);ctx.quadraticCurveTo(x,y+h,x,y+h-r);
  ctx.lineTo(x,y+r);ctx.quadraticCurveTo(x,y,x+r,y);ctx.closePath();
}

function getEntryDisplayFields(entry) {
  const cat = (entry.category||'').toUpperCase();
  const fields = CATEGORY_FIELDS[cat] || CATEGORY_FIELDS['OTHER'];
  const rows = [];
  fields.forEach(f => {
    if (f.key === 'rate') return;  // rate not shown in product details box
    const val = entry[f.key];
    if (val) rows.push({ label: f.label.replace(' *','').toUpperCase(), value: val.toString().toUpperCase() });
  });
  return rows;
}

/* ── TAB 3: CATALOG SPECS ─────────────────────────────────────────────────── */
function initSpecs() {
  setupDropZone('dz-excel', 'input-excel', handleExcelFile);
  document.getElementById('input-excel').addEventListener('change', e => handleExcelFile(e.target.files));
  document.getElementById('btn-clear-excel').addEventListener('click', clearExcel);
  document.getElementById('item-name-dropdown').addEventListener('change', onItemSelected);
  document.getElementById('btn-save-spec').addEventListener('click', saveSpecEntry);
  document.getElementById('btn-clear-form').addEventListener('click', clearSpecForm);
  document.getElementById('spec-category').addEventListener('change', () => {
    renderCategoryFields(document.getElementById('spec-category').value);
    scheduleLivePreview();
    updateFilenamePreview();
  });
}

/* ── EXCEL UPLOAD → DROPDOWN (new flow) ──────────────────────────────────── */
function handleExcelFile(files) {
  const file = Array.from(files).find(f => /\.(xlsx|xls|csv)$/i.test(f.name));
  if (!file) return toast('Upload a valid Excel/CSV file', 'warning');
  const r = new FileReader();
  r.onload = e => {
    try {
      const wb   = XLSX.read(new Uint8Array(e.target.result), { type: 'array' });
      const ws   = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(ws, { defval: '' });
      if (!rows.length) return toast('Excel appears empty', 'warning');

      STATE.excelRows    = rows;
      STATE.excelHeaders = Object.keys(rows[0]);

      /* Show status */
      document.getElementById('excel-status').style.display = 'block';
      document.getElementById('excel-file-name').textContent = file.name;
      document.getElementById('excel-row-count').textContent = `${rows.length} products`;

      /* Build dropdown */
      populateItemDropdown(rows);

      /* Build mini product list */
      renderExcelProductList(rows);

      /* Mark step 1 done */
      markStep(1);

      toast(`✓ ${rows.length} products loaded — pick one from the dropdown`, 'success', 4000);
    } catch (err) {
      toast('Excel parse failed', 'error');
    }
  };
  r.readAsArrayBuffer(file);
}

function populateItemDropdown(rows) {
  const dd = document.getElementById('item-name-dropdown');
  dd.innerHTML = '<option value="">— Choose a product —</option>';

  const nameKey = STATE.excelHeaders.find(h => {
    const n = h.toLowerCase();
    return n.includes('product name') || n.includes('item name') || n.includes('name of product');
  }) || STATE.excelHeaders.find(h => {
    const n = h.toLowerCase();
    return n.includes('product') || n.includes('item') || n.includes('name');
  }) || STATE.excelHeaders[0];

  rows.forEach((row, idx) => {
    const opt = document.createElement('option');
    opt.value       = idx;
    opt.textContent = (row[nameKey] || `Row ${idx + 1}`).toString().trim();
    dd.appendChild(opt);
  });

  document.getElementById('item-selector-wrap').style.display = 'block';
}

function renderExcelProductList(rows) {
  const wrap = document.getElementById('excel-product-list-wrap');
  const list = document.getElementById('excel-product-list');
  if (!wrap || !list) return;
  list.innerHTML = '';

  const nameKey = STATE.excelHeaders.find(h => {
    const n = h.toLowerCase();
    return n.includes('product name') || n.includes('item name') || n.includes('name');
  }) || STATE.excelHeaders[0];

  rows.forEach((row, idx) => {
    const item = document.createElement('div');
    item.className = 'excel-prod-item';
    item.textContent = (row[nameKey] || `Row ${idx+1}`).toString().trim();
    item.addEventListener('click', () => {
      document.getElementById('item-name-dropdown').value = idx;
      onItemSelected();
    });
    list.appendChild(item);
  });
  wrap.style.display = 'block';
}

function clearExcel() {
  STATE.excelRows = []; STATE.excelHeaders = [];
  document.getElementById('excel-status').style.display = 'none';
  document.getElementById('item-selector-wrap').style.display = 'none';
  const w = document.getElementById('excel-product-list-wrap');
  if (w) w.style.display = 'none';
  document.getElementById('item-name-dropdown').innerHTML = '<option value="">— Choose a product —</option>';
  document.getElementById('input-excel').value = '';
}

/* KEY FIX: select from dropdown → fill form + trigger live preview */
function onItemSelected() {
  const idx = document.getElementById('item-name-dropdown').value;
  if (idx === '') return;
  const row = STATE.excelRows[+idx];
  if (!row) return;

  /* Map columns */
  const mapped = {};
  Object.entries(row).forEach(([col, val]) => {
    const key = normaliseHeader(col);
    if (key && val !== '' && val !== null) mapped[key] = String(val).trim();
  });

  /* Set category first */
  if (mapped.category) {
    const sel   = document.getElementById('spec-category');
    const catUp = mapped.category.toUpperCase().trim();
    let found   = false;
    Array.from(sel.options).forEach(o => { if (o.value.toUpperCase() === catUp) { sel.value = o.value; found = true; } });
    if (!found) Array.from(sel.options).forEach(o => { if (catUp.includes(o.value.toUpperCase()) || o.value.toUpperCase().includes(catUp)) { sel.value = o.value; found = true; } });
    if (!found) { const no = new Option(mapped.category, mapped.category); sel.add(no); sel.value = mapped.category; }
    renderCategoryFields(sel.value);
  }

  /* Fill all fields after DOM renders */
  setTimeout(() => {
    Object.entries(mapped).forEach(([key, val]) => {
      const el = document.getElementById('dynfield-' + key);
      if (el) el.value = val;
    });
    highlightFormAsAutofilled();
    updateFilenamePreview();
    markStep(2); markStep(3);
    scheduleLivePreview(0);   // trigger immediately
  }, 30);

  toast(`✓ Loaded: ${document.getElementById('item-name-dropdown').selectedOptions[0].text}`, 'success');
}

function highlightFormAsAutofilled() {
  const g = document.getElementById('spec-form-grid');
  g.classList.add('autofill-glow');
  setTimeout(() => g.classList.remove('autofill-glow'), 2000);
}

/* ── CATEGORY FIELDS RENDERER ─────────────────────────────────────────────── */
function renderCategoryFields(category) {
  const grid = document.getElementById('spec-form-grid'); grid.innerHTML = '';
  if (!category) return;
  const fields = CATEGORY_FIELDS[category] || CATEGORY_FIELDS['OTHER'];
  fields.forEach(f => {
    const fg = document.createElement('div'); fg.className = 'field-group';
    if (f.span === 2) fg.style.gridColumn = '1/-1';
    const lbl = document.createElement('label'); lbl.className = 'field-label'; lbl.textContent = f.label; fg.appendChild(lbl);
    let inp;
    if (f.type === 'textarea') { inp = document.createElement('textarea'); inp.rows = 2; }
    else { inp = document.createElement('input'); inp.type = 'text'; }
    inp.className = 'field-input'; inp.id = 'dynfield-' + f.key; inp.placeholder = f.placeholder || '';
    /* Live preview on every keystroke */
    inp.addEventListener('input', () => { scheduleLivePreview(); updateFilenamePreview(); });
    fg.appendChild(inp); grid.appendChild(fg);
  });
  if (!fields.some(f => f.key === 'notes')) {
    const fg = document.createElement('div'); fg.className = 'field-group'; fg.style.gridColumn = '1/-1';
    fg.innerHTML = `<label class="field-label">Notes <span class="optional-badge">Optional</span></label><textarea class="field-input" id="dynfield-notes" rows="2" placeholder="Any extra information..."></textarea>`;
    grid.appendChild(fg);
  }
}

function getDynFieldData() {
  const data = {};
  document.querySelectorAll('[id^="dynfield-"]').forEach(el => { data[el.id.replace('dynfield-','')]=el.value.trim(); });
  return data;
}

function getFormData() {
  const category = document.getElementById('spec-category').value;
  const dyn = getDynFieldData();
  return { category, productName:dyn.productName||'', pieces:dyn.pieces||'', length:dyn.length||'', length2:dyn.length2||'', fabric:dyn.fabric||'', work:dyn.work||'', salesPackage:dyn.salesPackage||'', packing:dyn.packing||'', notes:dyn.notes||'', extraFabric1:dyn.extraFabric1||'', extraFabric2:dyn.extraFabric2||'', extraLen1:dyn.extraLen1||'', extraLen2:dyn.extraLen2||'', rate:dyn.rate||'' };
}

function saveSpecEntry() {
  const data = getFormData();
  if (!data.productName) { toast('Product Name is required.', 'warning'); return; }
  if (STATE.editingEntryId !== null) {
    const idx = STATE.savedEntries.findIndex(e => e.id === STATE.editingEntryId);
    if (idx >= 0) STATE.savedEntries[idx] = { ...data, id: STATE.editingEntryId };
    STATE.editingEntryId = null;
    document.getElementById('btn-save-spec').textContent = '✓ Save Entry';
  } else {
    STATE.savedEntries.push({ ...data, id: Date.now() });
  }
  renderSavedEntries(); clearSpecForm(); saveSession();
  markStep(4);
  toast('Entry saved!', 'success');
}

function clearSpecForm() {
  document.getElementById('spec-category').selectedIndex = 0;
  document.getElementById('spec-form-grid').innerHTML = '';
  STATE.editingEntryId = null;
  document.getElementById('btn-save-spec').textContent = '✓ Save Entry';
  hideLivePreview();
  document.getElementById('filename-preview-strip').style.display = 'none';
}

/* ── LIVE PREVIEW IN RIGHT COLUMN ─────────────────────────────────────────── */
function scheduleLivePreview(delay = 600) {
  clearTimeout(STATE.livePreviewTimer);
  STATE.livePreviewTimer = setTimeout(renderLivePreview, delay);
}

async function renderLivePreview() {
  if (!STATE.templates[1]) return; // no template = no preview

  const entry = getFormData();
  if (!entry.productName && !entry.fabric && !entry.pieces) return; // nothing to show

  showLiveSpinner(true);

  try {
    const dataURL = await compositeTemplate1(entry);
    const img   = document.getElementById('live-preview-img');
    const wrap  = document.getElementById('live-preview-img-wrap');
    const ph    = document.getElementById('live-preview-placeholder');
    const cap   = document.getElementById('live-preview-caption');

    img.src = dataURL;
    wrap.style.display = 'block';
    ph.style.display   = 'none';
    cap.textContent    = entry.productName ? `Preview: ${entry.productName}` : 'Live Preview';
    markStep(4);
  } catch (_) {
    /* silent fail */
  } finally {
    showLiveSpinner(false);
  }
}

function showLiveSpinner(show) {
  document.getElementById('live-preview-spinner').style.display = show ? 'flex' : 'none';
}

function hideLivePreview() {
  document.getElementById('live-preview-img-wrap').style.display = 'none';
  document.getElementById('live-preview-placeholder').style.display = 'flex';
}

/* ── PDF FILENAME PREVIEW ─────────────────────────────────────────────────── */
function buildPdfFilename(entry) {
  const r = entry.rate        ? `RS=${parseFloat(entry.rate||0).toFixed(2)}` : '';
  const n = entry.productName ? entry.productName.toUpperCase()               : '';
  const f = entry.fabric      ? `(${entry.fabric.toUpperCase()})`             : '';
  let fn  = [r, n].filter(Boolean).join(' - ');
  if (f) fn += ` ${f}`;
  return (fn || 'Catalog').replace(/[\\/:*?"<>|]/g, '_') + '.pdf';
}

function updateFilenamePreview() {
  const entry = getFormData();
  const strip = document.getElementById('filename-preview-strip');
  const val   = document.getElementById('filename-preview-val');
  if (!entry.productName) { strip.style.display = 'none'; return; }
  val.textContent    = buildPdfFilename(entry);
  strip.style.display = 'flex';
}

/* ── STEPS ────────────────────────────────────────────────────────────────── */
function markStep(n) {
  const el = document.getElementById(`step${n}-ind`);
  if (el) el.classList.add('done');
}

/* ── SAVED ENTRIES ────────────────────────────────────────────────────────── */
function renderSavedEntries() {
  const card  = document.getElementById('saved-entries-card');
  const tbody = document.getElementById('entries-tbody');
  const badge = document.getElementById('entries-count-badge');
  card.style.display = STATE.savedEntries.length ? 'block' : 'none';
  if (badge) badge.textContent = STATE.savedEntries.length;
  tbody.innerHTML = '';
  STATE.savedEntries.forEach((entry, idx) => {
    const fn = buildPdfFilename(entry);
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td>${idx+1}</td>
      <td>${entry.category||'—'}</td>
      <td><strong>${entry.productName}</strong></td>
      <td>${entry.pieces||'—'}</td>
      <td>${entry.fabric||'—'}</td>
      <td>${entry.work||'—'}</td>
      <td>${entry.rate?'₹'+entry.rate:'—'}</td>
      <td style="font-size:0.68rem;color:var(--clr-text-muted);max-width:160px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap" title="${fn}">${fn}</td>
      <td>
        <button class="btn-primary btn-sm" data-action="select" data-id="${entry.id}">Select</button>
        <button class="btn-secondary btn-sm" data-action="edit" data-id="${entry.id}">Edit</button>
        <button class="btn-danger-sm" data-action="delete" data-id="${entry.id}">✕</button>
      </td>`;
    tbody.appendChild(tr);
  });
  tbody.querySelectorAll('[data-action="select"]').forEach(btn => btn.addEventListener('click', () => loadSavedEntryToForm(+btn.dataset.id)));
  tbody.querySelectorAll('[data-action="edit"]').forEach(btn   => btn.addEventListener('click', () => openEditModal(+btn.dataset.id)));
  tbody.querySelectorAll('[data-action="delete"]').forEach(btn => btn.addEventListener('click', () => {
    if (!confirm('Delete this entry?')) return;
    STATE.savedEntries = STATE.savedEntries.filter(e => e.id !== +btn.dataset.id);
    renderSavedEntries(); saveSession();
  }));
}

function loadSavedEntryToForm(id) {
  const entry = STATE.savedEntries.find(e => e.id === id); if (!entry) return;
  const sel   = document.getElementById('spec-category');
  sel.value   = entry.category || 'OTHER';
  renderCategoryFields(sel.value);
  setTimeout(() => {
    Object.entries(entry).forEach(([key, val]) => { const el = document.getElementById('dynfield-' + key); if (el) el.value = val; });
    highlightFormAsAutofilled();
    updateFilenamePreview();
    scheduleLivePreview(0);
    toast(`Loaded: ${entry.productName}`, 'success');
  }, 50);
}

/* ── EDIT MODAL ───────────────────────────────────────────────────────────── */
function openEditModal(id) {
  const entry = STATE.savedEntries.find(e => e.id === id); if (!entry) return;
  STATE.editingEntryId = id;
  const body = document.getElementById('edit-modal-body'); body.innerHTML = '';
  const fields = [
    {key:'category',label:'Category',type:'text'},{key:'productName',label:'Product Name',type:'text'},
    {key:'pieces',label:'Pieces',type:'text'},{key:'length',label:'Length/Style',type:'text'},
    {key:'fabric',label:'Fabric',type:'text'},{key:'work',label:'Work/Pattern',type:'text'},
    {key:'salesPackage',label:'Sales Package',type:'text'},{key:'packing',label:'Packing',type:'text'},
    {key:'rate',label:'Rate ₹',type:'text'},{key:'notes',label:'Notes',type:'textarea'},
  ];
  const grid = document.createElement('div'); grid.style.cssText = 'display:grid;grid-template-columns:1fr 1fr;gap:12px';
  fields.forEach(f => {
    const fg = document.createElement('div'); fg.className = 'field-group';
    if (f.key === 'notes' || f.key === 'salesPackage') fg.style.gridColumn = '1/-1';
    const lbl = document.createElement('label'); lbl.className = 'field-label'; lbl.textContent = f.label;
    const inp = f.type === 'textarea' ? Object.assign(document.createElement('textarea'),{className:'field-input',rows:2}) : Object.assign(document.createElement('input'),{className:'field-input',type:'text'});
    inp.dataset.editKey = f.key; inp.value = entry[f.key] || '';
    fg.appendChild(lbl); fg.appendChild(inp); grid.appendChild(fg);
  });
  body.appendChild(grid);
  document.getElementById('edit-modal-overlay').style.display = 'flex';
}
function closeEditModal() { document.getElementById('edit-modal-overlay').style.display = 'none'; STATE.editingEntryId = null; }
function saveEditModal() {
  const idx = STATE.savedEntries.findIndex(e => e.id === STATE.editingEntryId); if (idx < 0) { closeEditModal(); return; }
  document.querySelectorAll('[data-edit-key]').forEach(inp => { STATE.savedEntries[idx][inp.dataset.editKey] = inp.value.trim(); });
  renderSavedEntries(); saveSession(); closeEditModal(); toast('Entry updated!', 'success');
}

/* ── TAB 4: OUTPUT ────────────────────────────────────────────────────────── */
function initOutput() {
  document.getElementById('btn-load-preview').addEventListener('click', loadCatalogPreview);
  document.getElementById('btn-generate-pdf').addEventListener('click', generatePDF);
}
function refreshOutputSummary() {
  document.getElementById('sum-images').textContent    = STATE.appliedImages.length || STATE.productImages.length;
  document.getElementById('sum-templates').textContent = [1,2,3].filter(n => STATE.templates[n]).length;
  document.getElementById('sum-entries').textContent   = STATE.savedEntries.length;
}
function renderOutputEntries() {
  const panel = document.getElementById('output-entries-panel');
  const list  = document.getElementById('output-entries-list');
  const fnbox = document.getElementById('output-filename-box');
  if (!STATE.savedEntries.length) { panel.style.display = 'none'; return; }
  panel.style.display = 'block';
  list.innerHTML = '';
  STATE.savedEntries.forEach((e, i) => {
    const item = document.createElement('div'); item.className = 'output-entry-item';
    item.innerHTML = `<span class="oei-num">${i+1}</span>
      <div class="oei-info">
        <strong>${e.productName}</strong>
        <span>${e.category||''} ${e.fabric ? '· '+e.fabric : ''}</span>
      </div>
      <span class="oei-rate">${e.rate?'₹'+e.rate:''}</span>`;
    list.appendChild(item);
  });
  if (fnbox && STATE.savedEntries[0]) {
    fnbox.innerHTML = `<span class="fn-label">PDF Filename:</span> <span class="fn-val">${buildPdfFilename(STATE.savedEntries[0])}</span>`;
  }
}
function buildPagesList() {
  const pages  = [], images = STATE.appliedImages.length ? STATE.appliedImages : STATE.productImages;
  images.forEach((img, i) => pages.push({ src: img.dataURL, label: `Product ${i+1}`, type: 'product' }));
  if (STATE.templates[1]) pages.push({ src: STATE.templates[1], label: 'Template 1 (Dynamic)', type: 'template1' });
  if (STATE.templates[2]) pages.push({ src: STATE.templates[2], label: 'Template 2',           type: 'template2' });
  if (STATE.templates[3]) pages.push({ src: STATE.templates[3], label: 'Template 3',           type: 'template3' });
  return pages;
}
function loadCatalogPreview() {
  const container = document.getElementById('catalog-pages-preview'); container.innerHTML = '';
  const pages = buildPagesList();
  if (!pages.length) { toast('No content to preview.', 'warning'); container.innerHTML='<div class="preview-placeholder"><div class="placeholder-icon">◈</div><p>No pages yet.</p></div>'; return; }
  pages.forEach((page, idx) => {
    const div = document.createElement('div'); div.className = 'catalog-page-thumb';
    div.innerHTML = `<img src="${page.src}" alt="Page ${idx+1}" loading="lazy"/><div class="page-label">Page ${idx+1} — ${page.label}</div>`;
    div.querySelector('img').addEventListener('click', () => openLightbox(page.src));
    container.appendChild(div);
  });
  renderOutputEntries();
  toast(`${pages.length} page(s) loaded.`, 'success');
}

async function generatePDF() {
  const pages = buildPagesList(); 
  if (!pages.length) return toast('No content to generate PDF from.', 'warning');

  // Master Constant: Match the Template Width (1816px)
  const MASTER_W = 1816; 
  const mm = v => v * 0.264583; // Pixel to Millimeter conversion factor

  const pw = document.getElementById('generation-progress'), 
        pb = document.getElementById('progress-bar'), 
        pl = document.getElementById('progress-label');
  
  pw.style.display = 'block';
  const btn = document.getElementById('btn-generate-pdf'); 
  btn.disabled = true; 
  btn.innerHTML = '⏳ Generating HD Catalog...';

  try {
    const { jsPDF } = window.jspdf;
    let pdf;

    for (let i = 0; i < pages.length; i++) {
      const page = pages[i];
      pb.style.width = Math.round((i / pages.length) * 90) + '%';
      pl.textContent = `Processing Page ${i + 1} of ${pages.length}...`;

      let src = page.src;
      // Handle Template 1 separately for dynamic text overlay
      if (page.type === 'template1' && STATE.savedEntries.length) {
        src = await compositeTemplate1(STATE.savedEntries[0]);
      }

      // Load image to calculate specific page height
      const img = new Image(); 
      img.src = src; 
      await new Promise(r => { img.onload = r; img.onerror = r; });

      // Calculate Height: Force width to 1816 and scale height proportionally
      const aspectRatio = img.naturalHeight / img.naturalWidth;
      const mmW = mm(MASTER_W);
      const mmH = mm(MASTER_W * aspectRatio);

      // Determine orientation based on calculated dimensions
      const ori = MASTER_W > (MASTER_W * aspectRatio) ? 'l' : 'p';

      // Add page with exact calculated dimensions to prevent cropping
      if (i === 0) {
        pdf = new jsPDF({ orientation: ori, unit: 'mm', format: [mmW, mmH] });
      } else {
        pdf.addPage([mmW, mmH], ori);
      }

      // Add image at 100% size with NO compression for high-quality output
      pdf.addImage(src, 'JPEG', 0, 0, mmW, mmH, undefined, 'NONE');
    }

    pb.style.width = '100%'; 
    pl.textContent = 'Finalizing PDF...';
    
    // Slight delay to ensure buffer is ready
    await new Promise(r => setTimeout(r, 200));

    const filename = STATE.savedEntries.length ? buildPdfFilename(STATE.savedEntries[0]) : 'CatalogCraft_Catalog.pdf';
    pdf.save(filename);

    pw.style.display = 'none'; 
    btn.disabled = false; 
    btn.innerHTML = '⬇ Generate & Download PDF';
    toast(`✓ Saved as: ${filename}`, 'success', 5000);
  } catch (err) {
    console.error(err); 
    pw.style.display = 'none'; 
    btn.disabled = false;
    btn.innerHTML = '⬇ Generate & Download PDF'; 
    toast('PDF generation failed.', 'error');
  }
}

/* ── DEFAULT ASSETS ───────────────────────────────────────────────────────── */
async function loadDefaultAssets() {
  const response = await fetch('/api/predefined-assets');
  const data = await response.json();
  if (data.logo) {
    STATE.logoDataURL = data.logo;
    const thumb = document.getElementById('logo-preview-thumb');
    if (thumb) thumb.src = data.logo;
    document.getElementById('logo-controls').style.display = 'block';
    document.getElementById('dz-logo').style.display = 'none';
    document.getElementById('btn-use-default-logo').classList.add('active');
    document.getElementById('btn-show-logo-upload').classList.remove('active');
  }
  [1, 2, 3].forEach(n => {
    if (data.templates[n]) {
      STATE.templates[n] = data.templates[n];
      const img = document.getElementById(`tmpl${n}-img`); if (img) img.src = data.templates[n];
      document.getElementById(`tmpl${n}-preview-wrap`).style.display = 'block';
      document.getElementById(`dz-tmpl${n}`).style.display = 'none';
      const btns = document.getElementById(`tcard-${n}`).querySelectorAll('.btn-toggle');
      btns[0].classList.add('active'); btns[1].classList.remove('active');
    }
  });
}
function setupAssetToggles() {
  document.getElementById('btn-show-logo-upload').onclick = () => {
    document.getElementById('dz-logo').style.display = 'block';
    document.getElementById('btn-show-logo-upload').classList.add('active');
    document.getElementById('btn-use-default-logo').classList.remove('active');
  };
  document.getElementById('btn-use-default-logo').onclick = loadDefaultAssets;
}
async function useDefaultTemplate(n) {
  const data = await (await fetch('/api/predefined-assets')).json();
  STATE.templates[n] = data.templates[n];
  document.getElementById(`tmpl${n}-img`).src = data.templates[n];
  document.getElementById(`tmpl${n}-preview-wrap`).style.display = 'block';
  document.getElementById(`dz-tmpl${n}`).style.display = 'none';
  const btns = document.getElementById(`tcard-${n}`).querySelectorAll('.btn-toggle');
  btns[0].classList.add('active'); btns[1].classList.remove('active');
}
function showTemplateUpload(n) {
  document.getElementById(`dz-tmpl${n}`).style.display = 'block';
  document.getElementById(`tmpl${n}-preview-wrap`).style.display = 'none';
  const btns = document.getElementById(`tcard-${n}`).querySelectorAll('.btn-toggle');
  btns[0].classList.remove('active'); btns[1].classList.add('active');
}
