// ══════════════════════════════════════════════════════
//  COMASA PRO — script.js  v2.0  (Flask + MySQL)
// ══════════════════════════════════════════════════════

const API = 'https://costos-comasa.onrender.com';   // mismo origen; si Flask corre en otro puerto: 'http://localhost:5000'

// ─── CATÁLOGO en memoria (cargado desde la BD) ───
let CAT = {};             // { "NOMBRE CATEGORÍA": [ {codigo, descripcion, peso_unit, unidad}, ... ] }
let CATEGORIAS = [];      // [ {id_categoria, nombre} ]

// ─── ESTADO ───
let items    = [];
let activeId = null;
let editingRate = null;
let toastTm    = null;
let puFab  = 15.20;
let puMont =  5.20;

// ─── DOM ───
const $cat    = document.getElementById('categoria');
const $prod   = document.getElementById('producto');
const $preview= document.getElementById('preview');
const $cant   = document.getElementById('cantidad');
const $unidad = document.getElementById('unidad');
const $est    = document.getElementById('est_box');
const $estV   = document.getElementById('est_value');

// ══════════════════════════════════════════════════════
//  CARGA INICIAL DESDE API
// ══════════════════════════════════════════════════════
async function cargarCatalogo() {
  // 1. Verificar conexión
  try {
    const r = await fetch(`${API}/api/health`);
    const d = await r.json();
    const badge = document.getElementById('db_badge');
    const status = document.getElementById('db_status');
    if (d.status === 'ok') {
      badge.className = 'db-badge ok';
      status.textContent = 'BD conectada';
    } else {
      throw new Error('sin conexión');
    }
  } catch {
    const badge = document.getElementById('db_badge');
    document.getElementById('db_status').textContent = 'sin BD';
    badge.className = 'db-badge err';
    toast('⚠ No se pudo conectar a la base de datos', 1);
    return;
  }

  // 2. Cargar categorías
  try {
    const r = await fetch(`${API}/api/categorias`);
    CATEGORIAS = await r.json();

    $cat.innerHTML = '<option value="">— Seleccionar categoría —</option>';
    CATEGORIAS.forEach(c => {
      const o = document.createElement('option');
      o.value = c.id_categoria;
      o.textContent = c.nombre;
      $cat.appendChild(o);
    });
  } catch (e) {
    toast('Error al cargar categorías', 1);
    return;
  }

  // 3. Pre-cargar todos los productos en memoria
  try {
    const r = await fetch(`${API}/api/productos`);
    const prods = await r.json();
    CAT = {};
    prods.forEach(p => {
      if (!CAT[p.categoria]) CAT[p.categoria] = [];
      CAT[p.categoria].push(p);
    });
  } catch (e) {
    toast('Error al cargar productos', 1);
  }
}

// ══════════════════════════════════════════════════════
//  WORKFLOW
// ══════════════════════════════════════════════════════
let currentStep = 0;
function goStep(n) {
  currentStep = n;
  document.querySelectorAll('.step-panel').forEach((p,i) => p.classList.toggle('active', i===n));
  document.querySelectorAll('.wf-step').forEach((s,i) => {
    s.classList.toggle('active', i===n);
    s.classList.toggle('done', i<n);
  });
  document.querySelectorAll('.wf-connector').forEach((c,i) => c.classList.toggle('done', i<n));
}

// ══════════════════════════════════════════════════════
//  ÍTEMS
// ══════════════════════════════════════════════════════
function openItemModal() {
  const modal = document.getElementById('item_modal');
  const inp   = document.getElementById('modal_input');
  inp.value   = document.getElementById('new_item_name').value;
  modal.classList.add('open');
  setTimeout(() => inp.focus(), 120);
}
function closeItemModal() {
  document.getElementById('item_modal').classList.remove('open');
}
function createItem() {
  const name = document.getElementById('modal_input').value.trim();
  if (!name) { toast('Escribe un nombre para el ítem', 1); return; }
  const id = 'i' + Date.now();
  items.push({ id, nombre: name, subs: [] });
  setActive(id);
  document.getElementById('new_item_name').value = '';
  closeItemModal();
  renderItems(); renderTable(); updateStats();
  toast('Ítem creado');
  goStep(1);
}
function deleteItem(id) {
  const it = items.find(x => x.id === id);
  if (!it) return;
  if (it.subs.length > 0 && !confirm(`¿Eliminar "${it.nombre}" con sus ${it.subs.length} sub-ítem(s)?`)) return;
  items = items.filter(x => x.id !== id);
  if (activeId === id) setActive(items.length ? items[items.length-1].id : null);
  else setActive(activeId);
  renderItems(); renderTable(); updateStats();
  toast('Ítem eliminado');
}
function setActive(id) {
  activeId = id;
  const it = items.find(x => x.id === id);
  const banner = document.getElementById('active_banner');
  if (it) {
    const idx = items.indexOf(it);
    banner.className = 'active-banner';
    banner.innerHTML = `<span class="ab-name">${idx+1}. ${it.nombre}</span><span class="ab-count">${it.subs.length} sub-ítem(s)</span>`;
  } else {
    banner.className = 'active-banner none';
    banner.innerHTML = '<span class="ab-name">Selecciona un ítem en el paso 1</span>';
  }
  renderItems();
}
function renderItems() {
  const list  = document.getElementById('items_list');
  const badge = document.getElementById('item_count_badge');
  badge.textContent = items.length ? `(${items.length})` : '';
  if (!items.length) {
    list.innerHTML = `<div class="empty"><div class="empty-icon">📦</div><p>Aún no tienes ítems.<br><b>Crea el primero</b> arriba.</p></div>`;
    return;
  }
  list.innerHTML = '';
  items.forEach((it, idx) => {
    const total = it.subs.reduce((s,x) => s+x.sub, 0);
    const card  = document.createElement('div');
    card.className = 'item-card' + (it.id===activeId ? ' active' : '');
    card.innerHTML = `
      <div class="item-card-header" onclick="setActive('${it.id}');goStep(1)">
        <div class="item-num">${idx+1}</div>
        <div class="item-info">
          <div class="item-name">${it.nombre}</div>
          <div class="item-meta">${it.subs.length} sub-ítem(s)${it.subs.length?` · ${it.subs.reduce((s,x)=>s+x.pesoT,0).toFixed(2)} kg`:''}</div>
        </div>
        ${total>0?`<div class="item-total">S/ ${total.toFixed(2)}</div>`:''}
        <button class="item-del" onclick="event.stopPropagation();deleteItem('${it.id}')" title="Eliminar ítem">✕</button>
      </div>
      <div class="item-footer">
        <button class="item-action" onclick="setActive('${it.id}');goStep(1)">+ Agregar sub-ítem</button>
      </div>
    `;
    list.appendChild(card);
  });
}

// ══════════════════════════════════════════════════════
//  SELECTORES DE CATÁLOGO
// ══════════════════════════════════════════════════════
function onCategoryChange() {
  const catId = $cat.value;
  $prod.innerHTML = '<option value="">— Seleccionar producto —</option>';
  if (!catId) { updatePreview(null); return; }

  // Buscar nombre de categoría
  const catObj = CATEGORIAS.find(c => c.id_categoria == catId);
  const catNombre = catObj ? catObj.nombre : '';
  const prods = CAT[catNombre] || [];

  prods.forEach(p => {
    const o = document.createElement('option');
    o.value = p.codigo;
    o.setAttribute('data-cat', catNombre);
    o.textContent = p.descripcion;
    $prod.appendChild(o);
  });
  updatePreview(null);
}

function onProductChange() {
  const p = getProduct();
  updatePreview(p);
  if (p) { $cant.focus(); updateEstimated(); }
}

function getProduct() {
  const cod = $prod.value;
  const opt = $prod.options[$prod.selectedIndex];
  if (!cod || !opt) return null;
  const catNombre = opt.getAttribute('data-cat') || '';
  if (!catNombre) return null;
  return (CAT[catNombre] || []).find(p => p.codigo === cod) || null;
}

function updatePreview(p) {
  if (!p) {
    $preview.className = 'product-card';
    $preview.innerHTML = '<span class="product-card-text">Selecciona un producto para ver sus detalles</span>';
    $unidad.value = '';
    $est.style.display = 'none';
    return;
  }
  $preview.className = 'product-card has-product';
  $preview.innerHTML = `<span class="product-card-text">${p.descripcion}</span><span class="product-badge">${p.peso_unit} kg/${p.unidad}</span>`;
  $unidad.value = p.unidad;
}

function updateEstimated() {
  const p = getProduct();
  const c = parseFloat($cant.value) || 0;
  if (!p || c <= 0) { $est.style.display='none'; return; }
  const sub = c * p.peso_unit * (puFab + puMont);
  $estV.textContent = `S/ ${sub.toFixed(2)}`;
  $est.style.display = 'flex';
}
$cant.addEventListener('input', updateEstimated);

// ══════════════════════════════════════════════════════
//  SUB-ÍTEMS
// ══════════════════════════════════════════════════════
function addSubItem() {
  if (!activeId) { toast('Selecciona un ítem primero', 1); return; }
  const p = getProduct();
  if (!p) { toast('Selecciona un producto', 1); return; }
  const cant = parseFloat($cant.value);
  if (!cant || cant <= 0) { toast('Ingresa una cantidad válida', 1); return; }
  const pesoT = cant * p.peso_unit;
  const sub   = pesoT * (puFab + puMont);
  const item  = items.find(x => x.id === activeId);
  item.subs.push({
    id: 's'+Date.now(),
    desc: p.descripcion,
    cant, unidad: p.unidad,
    pesoU: p.peso_unit, pesoT, puFab, puMont, sub
  });
  renderTable(); updateStats(); renderItems();
  setActive(activeId);
  $cant.value = '';
  $est.style.display = 'none';
  toast('Sub-ítem agregado ✓');
  $cant.focus();
}
function deleteSub(itemId, subId) {
  const it = items.find(x => x.id === itemId);
  if (!it) return;
  it.subs = it.subs.filter(s => s.id !== subId);
  renderTable(); updateStats(); renderItems(); setActive(activeId);
  toast('Sub-ítem eliminado');
}
function undoLast() {
  for (let i = items.length-1; i >= 0; i--) {
    if (items[i].subs.length) {
      const last = items[i].subs[items[i].subs.length-1];
      deleteSub(items[i].id, last.id);
      return;
    }
  }
  toast('No hay más sub-ítems', 1);
}
function clearAll() {
  if (!items.length) return;
  if (!confirm('¿Eliminar todos los ítems y sub-ítems?')) return;
  items = []; activeId = null;
  setActive(null); renderItems(); renderTable(); updateStats();
  toast('Todo limpiado');
}

// ══════════════════════════════════════════════════════
//  TABLA
// ══════════════════════════════════════════════════════
function renderTable() {
  const tbody = document.getElementById('tbody');
  const empty = document.getElementById('empty_table');
  tbody.innerHTML = '';
  if (!items.length) { empty.style.display='flex'; return; }
  empty.style.display = 'none';
  items.forEach((it,ii) => {
    const iPeso = it.subs.reduce((s,x) => s+x.pesoT, 0);
    const iSub  = it.subs.reduce((s,x) => s+x.sub, 0);
    const mr = document.createElement('tr');
    mr.className = 'row-main';
    mr.onclick = () => { setActive(it.id); goStep(1); };
    mr.innerHTML = `
      <td class="row-main-num r">${ii+1}</td>
      <td class="row-main-name">${it.nombre}</td>
      <td class="c" colspan="2" style="color:var(--text-dim);font-size:11px">${it.subs.length} sub-ítem(s)</td>
      <td></td>
      <td class="r" style="font-family:'DM Mono',monospace;color:var(--blue-text);font-size:12px">${iPeso>0?iPeso.toFixed(3):''}</td>
      <td></td><td></td>
      <td class="row-main-total r">${iSub>0?'S/ '+iSub.toFixed(2):''}</td>
      <td class="c"><button class="del" onclick="event.stopPropagation();deleteItem('${it.id}')">✕</button></td>
    `;
    tbody.appendChild(mr);
    it.subs.forEach((s,si) => {
      const sr = document.createElement('tr');
      sr.className = 'row-sub';
      sr.innerHTML = `
        <td class="sub-idx">${ii+1}.${si+1}</td>
        <td class="sub-desc">${s.desc}</td>
        <td class="c">${s.cant}</td>
        <td class="c" style="color:var(--text-dim)">${s.unidad}</td>
        <td class="sub-weight">${s.pesoU.toFixed(3)}</td>
        <td class="sub-weight">${s.pesoT.toFixed(3)}</td>
        <td class="sub-price">${s.puFab.toFixed(2)}</td>
        <td class="sub-price">${s.puMont.toFixed(2)}</td>
        <td class="sub-total">${s.sub.toFixed(2)}</td>
        <td class="c"><button class="del" onclick="deleteSub('${it.id}','${s.id}')">✕</button></td>
      `;
      tbody.appendChild(sr);
    });
  });
}

// ══════════════════════════════════════════════════════
//  STATS
// ══════════════════════════════════════════════════════
function updateStats() {
  const all  = items.flatMap(x => x.subs);
  const kg   = all.reduce((s,x) => s+x.pesoT, 0);
  const fab  = all.reduce((s,x) => s+x.pesoT*x.puFab, 0);
  const mont = all.reduce((s,x) => s+x.pesoT*x.puMont, 0);
  const tot  = all.reduce((s,x) => s+x.sub, 0);
  const ts   = 'S/ '+tot.toFixed(2);
  document.getElementById('s_items').textContent  = items.length;
  document.getElementById('s_subs').textContent   = all.length;
  document.getElementById('s_kg').textContent     = kg.toFixed(3);
  document.getElementById('s_fab').textContent    = fab.toFixed(2);
  document.getElementById('s_mont').textContent   = mont.toFixed(2);
  document.getElementById('s_total').textContent  = ts;
  document.getElementById('top_total').textContent= ts;
  document.getElementById('footer_total').textContent = 'TOTAL: '+ts;
  document.getElementById('sum_fab').textContent  = 'S/ '+fab.toFixed(2);
  document.getElementById('sum_mont').textContent = 'S/ '+mont.toFixed(2);
  document.getElementById('sum_total').textContent= ts;
}

// ══════════════════════════════════════════════════════
//  TARIFAS
// ══════════════════════════════════════════════════════
function openRateModal(which) {
  editingRate = which;
  const isFab = which === 'fab';
  document.getElementById('rate_modal_title').textContent = isFab ? 'Tarifa de Fabricación' : 'Tarifa de Montaje';
  document.getElementById('rate_modal_hint').textContent  = isFab
    ? 'Precio por kg de acero fabricado. Se aplicará a todos los sub-ítems.'
    : 'Precio por kg de acero montado. Se aplicará a todos los sub-ítems.';
  const inp = document.getElementById('rate_input');
  inp.value = isFab ? puFab : puMont;
  document.getElementById('rate_modal').classList.add('open');
  setTimeout(() => { inp.focus(); inp.select(); }, 120);
}
function closeRateModal() {
  document.getElementById('rate_modal').classList.remove('open');
  editingRate = null;
}
function saveRate() {
  const v = parseFloat(document.getElementById('rate_input').value) || 0;
  if (editingRate==='fab') { puFab = v; document.getElementById('fab_disp').textContent = v.toFixed(2); }
  else { puMont = v; document.getElementById('mont_disp').textContent = v.toFixed(2); }
  items.forEach(it => it.subs.forEach(s => {
    s.puFab = puFab; s.puMont = puMont;
    s.sub   = s.pesoT * (puFab + puMont);
  }));
  renderTable(); updateStats(); renderItems(); updateEstimated();
  closeRateModal();
  toast('Tarifas actualizadas y recalculadas ✓');
}
document.getElementById('rate_input').addEventListener('keydown', e => { if(e.key==='Enter') saveRate(); });

// ══════════════════════════════════════════════════════
//  MODAL CATÁLOGO — AGREGAR PRODUCTO / CATEGORÍA
// ══════════════════════════════════════════════════════
function openCatModal() {
  // Rellenar select de categorías
  const sel = document.getElementById('new_cat_sel');
  sel.innerHTML = '<option value="">— Selecciona categoría —</option>';
  CATEGORIAS.forEach(c => {
    const o = document.createElement('option');
    o.value = c.id_categoria;
    o.textContent = c.nombre;
    sel.appendChild(o);
  });
  document.getElementById('cat_modal').classList.add('open');
}
function closeCatModal() {
  document.getElementById('cat_modal').classList.remove('open');
  // limpiar
  ['new_codigo','new_desc','new_peso','new_cat_nombre','new_cat_desc'].forEach(id => {
    document.getElementById(id).value = '';
  });
  document.getElementById('new_cat_sel').value = '';
}

function switchTab(tab) {
  document.getElementById('panel_prod').style.display = tab==='prod' ? '' : 'none';
  document.getElementById('panel_cat').style.display  = tab==='cat'  ? '' : 'none';
  document.getElementById('tab_prod').classList.toggle('active', tab==='prod');
  document.getElementById('tab_cat').classList.toggle('active',  tab==='cat');
}

// Mostrar/ocultar campo "otro" en unidad
document.getElementById('new_unidad').addEventListener('change', function() {
  const wrap = document.getElementById('unidad_otro_wrap');
  wrap.style.display = this.value === 'otro' ? '' : 'none';
});

async function guardarProducto() {
  const catId  = document.getElementById('new_cat_sel').value;
  const codigo = document.getElementById('new_codigo').value.trim();
  const desc   = document.getElementById('new_desc').value.trim();
  const peso   = parseFloat(document.getElementById('new_peso').value);
  let unidad   = document.getElementById('new_unidad').value;
  if (unidad === 'otro') unidad = document.getElementById('new_unidad_otro').value.trim();

  if (!catId)   { toast('Selecciona una categoría', 1); return; }
  if (!codigo)  { toast('Ingresa un código', 1); return; }
  if (!desc)    { toast('Ingresa una descripción', 1); return; }
  if (!peso || peso <= 0) { toast('Ingresa un peso válido mayor a 0', 1); return; }
  if (!unidad)  { toast('Ingresa la unidad', 1); return; }

  setCatLoading(true);
  try {
    const r = await fetch(`${API}/api/productos`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ id_categoria: catId, codigo, descripcion: desc, peso_unit: peso, unidad })
    });
    const data = await r.json();
    if (!r.ok) { toast(data.error || 'Error al guardar', 1); return; }

    // Agregar al catálogo en memoria
    const catObj = CATEGORIAS.find(c => c.id_categoria == catId);
    if (catObj) {
      if (!CAT[catObj.nombre]) CAT[catObj.nombre] = [];
      CAT[catObj.nombre].push({ ...data, categoria: catObj.nombre });
    }

    toast(`✓ Producto "${codigo}" guardado en el catálogo`);
    closeCatModal();

    // Refrescar selector si la categoría está activa
    if ($cat.value == catId) onCategoryChange();

  } catch (e) {
    toast('Error de conexión con la BD', 1);
  } finally {
    setCatLoading(false);
  }
}

async function guardarCategoria() {
  const nombre = document.getElementById('new_cat_nombre').value.trim();
  const desc   = document.getElementById('new_cat_desc').value.trim();
  if (!nombre) { toast('Ingresa un nombre de categoría', 1); return; }

  setCatLoading(true);
  try {
    const r = await fetch(`${API}/api/categorias`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ nombre, descripcion: desc })
    });
    const data = await r.json();
    if (!r.ok) { toast(data.error || 'Error al guardar', 1); return; }

    // Agregar a memoria y selectores
    CATEGORIAS.push({ id_categoria: data.id_categoria, nombre });
    CAT[nombre] = [];

    // Actualizar selector principal
    const o1 = document.createElement('option');
    o1.value = data.id_categoria; o1.textContent = nombre;
    $cat.appendChild(o1);

    toast(`✓ Categoría "${nombre}" creada`);
    // Cambiar a tab de producto para agregar productos a esa categoría
    switchTab('prod');
    document.getElementById('new_cat_nombre').value = '';
    document.getElementById('new_cat_desc').value   = '';
    // Pre-seleccionar la nueva categoría
    document.getElementById('new_cat_sel').innerHTML += `<option value="${data.id_categoria}" selected>${nombre}</option>`;

  } catch (e) {
    toast('Error de conexión con la BD', 1);
  } finally {
    setCatLoading(false);
  }
}

function setCatLoading(v) {
  document.getElementById('cat_loading').className = 'loading-overlay' + (v ? ' visible' : '');
}

// ══════════════════════════════════════════════════════
//  EXPORT XLSX
// ══════════════════════════════════════════════════════
function exportXLSX() {
  const all = items.flatMap(x => x.subs);
  if (!all.length) { toast('No hay datos para exportar', 1); return; }

  const fecha = new Date().toLocaleDateString('es-PE');
  const notas = document.getElementById('notas').value;
  const wb = XLSX.utils.book_new();
  const ws = {};
  const merges = [];
  let R = 0;
  const COLS = 8; // columnas 0..8 → 9 columnas en total

  // ── helpers ──
  function setCell(c, r, v, s) {
    ws[XLSX.utils.encode_cell({ c, r })] = { v, t: typeof v === 'number' ? 'n' : 's', s };
  }
  function numCell(c, r, v, fmt, s) {
    ws[XLSX.utils.encode_cell({ c, r })] = { v, t: 'n', z: fmt || '#,##0.00', s };
  }
  function fmlaCell(c, r, f, fmt, s) {
    ws[XLSX.utils.encode_cell({ c, r })] = { f, t: 'n', z: fmt || '#,##0.00', s };
  }
  function merge(c1, r1, c2, r2) {
    merges.push({ s: { c: c1, r: r1 }, e: { c: c2, r: r2 } });
  }

  // ── estilos ──
  const sBrand       = { font:{bold:true,sz:14,color:{rgb:'FFFFFF'}}, fill:{fgColor:{rgb:'1A3A5C'}}, alignment:{horizontal:'center',vertical:'center'} };
  const sMeta        = { font:{sz:10,color:{rgb:'333333'}}, fill:{fgColor:{rgb:'EEF2F7'}}, alignment:{horizontal:'left'} };
  const sMetaLabel   = { font:{bold:true,sz:10,color:{rgb:'1A3A5C'}}, fill:{fgColor:{rgb:'EEF2F7'}} };
  const sHeader      = { font:{bold:true,sz:10,color:{rgb:'FFFFFF'}}, fill:{fgColor:{rgb:'2E6DA4'}}, alignment:{horizontal:'center',wrapText:true} };
  const sItemGroup   = { font:{bold:true,sz:10,color:{rgb:'1A3A5C'}}, fill:{fgColor:{rgb:'D6E4F0'}}, alignment:{horizontal:'left'} };
  const sItemGroupR  = { font:{bold:true,sz:10,color:{rgb:'1A3A5C'}}, fill:{fgColor:{rgb:'D6E4F0'}}, alignment:{horizontal:'right'} };
  const sSub         = { font:{sz:9,color:{rgb:'222222'}}, fill:{fgColor:{rgb:'FFFFFF'}}, alignment:{horizontal:'left',indent:1} };
  const sSubNum      = { font:{sz:9,color:{rgb:'222222'}}, fill:{fgColor:{rgb:'FFFFFF'}}, alignment:{horizontal:'right'} };
  const sSubFml      = { font:{sz:9,color:{rgb:'000099'}}, fill:{fgColor:{rgb:'FFFFFF'}}, alignment:{horizontal:'right'} };
  const sSubRate     = { font:{sz:9,color:{rgb:'0000CC'}}, fill:{fgColor:{rgb:'FFFFFF'}}, alignment:{horizontal:'right'} };
  const sTotal       = { font:{bold:true,sz:11,color:{rgb:'FFFFFF'}}, fill:{fgColor:{rgb:'1A3A5C'}}, alignment:{horizontal:'right'} };
  const sTotalLabel  = { font:{bold:true,sz:11,color:{rgb:'FFFFFF'}}, fill:{fgColor:{rgb:'1A3A5C'}}, alignment:{horizontal:'center'} };

  // ── CABECERA MARCA ──
  merge(0, R, COLS, R);
  setCell(0, R, '🏗  COMASA — Cotización', sBrand);
  for (let c = 1; c <= COLS; c++) setCell(c, R, '', sBrand);
  R++;

  // ── META ──
  const metaRows = [
    ['Fecha', fecha],
    notas ? ['Notas', notas] : null,
    ['Tarifa Fabricación', `S/ ${puFab.toFixed(2)} / kg`],
    ['Tarifa Montaje',     `S/ ${puMont.toFixed(2)} / kg`],
  ].filter(Boolean);

  metaRows.forEach(([lbl, val]) => {
    setCell(0, R, lbl, sMetaLabel);
    merge(1, R, COLS, R);
    setCell(1, R, val, sMeta);
    for (let c = 2; c <= COLS; c++) setCell(c, R, '', sMeta);
    R++;
  });
  R++; // fila vacía separadora

  // ── ENCABEZADOS COLUMNAS ──
  ['#', 'Descripción', 'Cant.', 'Unidad', 'Peso Unit (kg)', 'Peso Total (kg)', 'PU Fab (S/)', 'PU Mont (S/)', 'Subtotal (S/)']
    .forEach((h, c) => setCell(c, R, h, sHeader));
  R++;

  // ── DATOS ──
  const itemTotalExcelRows = []; // filas Excel (base 1) de los totales de cada ítem

  items.forEach((it, i) => {
    if (!it.subs.length) return; // saltar ítems sin sub-ítems

    const firstSubR = R; // fila base-0 donde empieza el primer sub-ítem

    // Insertar sub-ítems primero
    it.subs.forEach((s, j) => {
      const exR = R + 1; // fila Excel base-1
      setCell(0, R, `${i + 1}.${j + 1}`, sSub);
      setCell(1, R, s.desc,   sSub);
      numCell(2, R, s.cant,   '0.##',           sSubNum);
      setCell(3, R, s.unidad, sSub);
      numCell(4, R, s.pesoU,  '0.000',           sSubNum);
      fmlaCell(5, R, `C${exR}*E${exR}`, '0.000', sSubFml);
      numCell(6, R, s.puFab,  '"S/"#,##0.00',    sSubRate);
      numCell(7, R, s.puMont, '"S/"#,##0.00',    sSubRate);
      fmlaCell(8, R, `F${exR}*(G${exR}+H${exR})`, '"S/"#,##0.00', sSubFml);
      R++;
    });

    const lastSubR = R - 1; // fila base-0 del último sub-ítem
    // Convertir a Excel base-1
    const exFirst = firstSubR + 1;
    const exLast  = lastSubR  + 1;
    const exGroup = R + 1;           // fila Excel base-1 para la fila de grupo

    // Fila resumen del ítem (DESPUÉS de los sub-ítems)
    setCell(0, R, `${i + 1}`, sItemGroup);
    merge(1, R, 7, R);
    setCell(1, R, it.nombre, sItemGroup);
    for (let c = 2; c <= 7; c++) setCell(c, R, '', sItemGroup);
    fmlaCell(8, R, `SUM(I${exFirst}:I${exLast})`, '"S/"#,##0.00', sItemGroupR);
    itemTotalExcelRows.push(exGroup);
    R++;
    R++; // fila vacía entre ítems
  });

  // ── TOTAL GENERAL ──
  merge(0, R, 4, R);
  setCell(0, R, 'TOTAL GENERAL', sTotalLabel);
  for (let c = 1; c <= 4; c++) setCell(c, R, '', sTotalLabel);
  setCell(5, R, '', sTotal);
  setCell(6, R, '', sTotal);
  setCell(7, R, '', sTotal);
  // Suma de todas las filas de grupo de ítem
  const totalFormula = itemTotalExcelRows.map(r => `I${r}`).join('+');
  fmlaCell(8, R, totalFormula, '"S/"#,##0.00', sTotal);
  R++;

  // ── CONFIGURAR HOJA ──
  ws['!ref']    = XLSX.utils.encode_range({ s: { c: 0, r: 0 }, e: { c: COLS, r: R } });
  ws['!merges'] = merges;
  ws['!cols']   = [{ wch:8 },{ wch:36 },{ wch:7 },{ wch:8 },{ wch:13 },{ wch:13 },{ wch:12 },{ wch:12 },{ wch:14 }];
  ws['!rows']   = [{ hpt:28 }];
  XLSX.utils.book_append_sheet(wb, ws, 'Cotización');

  // ── HOJA TARIFAS ──
  const wsTar = {};
  wsTar['A1'] = { v:'Parámetro',           t:'s', s:sHeader };
  wsTar['B1'] = { v:'Valor',               t:'s', s:sHeader };
  wsTar['A2'] = { v:'Tarifa Fabricación (S//kg)', t:'s' };
  wsTar['B2'] = { v:puFab,  t:'n', z:'"S/"0.00' };
  wsTar['A3'] = { v:'Tarifa Montaje (S//kg)',      t:'s' };
  wsTar['B3'] = { v:puMont, t:'n', z:'"S/"0.00' };
  wsTar['!ref']  = 'A1:B3';
  wsTar['!cols'] = [{ wch:28 },{ wch:14 }];
  XLSX.utils.book_append_sheet(wb, wsTar, 'Tarifas');

  // ── DESCARGAR ──
  const fechaFile = fecha.replace(/\//g, '-');
  XLSX.writeFile(wb, `COMASA_${fechaFile}.xlsx`, { bookType:'xlsx', cellStyles:true });
  toast('✅ Excel exportado con fórmulas');
}

// ══════════════════════════════════════════════════════
//  TOAST
// ══════════════════════════════════════════════════════
function toast(msg, err=0) {
  clearTimeout(toastTm);
  const el = document.getElementById('toast');
  el.textContent  = err ? '✕ '+msg : '✓ '+msg;
  el.className    = 'toast show' + (err ? ' err' : '');
  toastTm = setTimeout(() => el.classList.remove('show'), 2600);
}

// ══════════════════════════════════════════════════════
//  ARRANQUE
// ══════════════════════════════════════════════════════
document.addEventListener('DOMContentLoaded', () => {
  cargarCatalogo();
});
