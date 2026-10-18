/**
 * PEDIDOS A PROVEEDORES · Buenas y Santas
 * Conecta la app de la tablet (GitHub Pages) con esta hoja de Google.
 *
 * PASOS (solo la primera vez):
 *  1. Pega este código y guarda.
 *  2. Implementar → Nueva implementación → tipo "Aplicación web"
 *     · Ejecutar como: Yo   · Quién tiene acceso: Cualquier usuario
 *  Los proveedores y productos se leen SIEMPRE de las pestañas APP_Proveedores y APP_Productos:
 *  aquí no hay ninguna lista escrita. Si falta alguna pestaña APP_, se crea vacía (solo títulos).
 *  Tus otras pestañas no se tocan.
 *
 * Si más adelante cambias este código: Implementar → Gestionar implementaciones →
 * ✏ editar → Versión: "Nueva versión". Así la URL no cambia.
 */

// Sin clave: la app conecta directamente. Si algún día quieres protegerla, pon aquí una palabra
// y la misma en la app (⚙ Ajustes).
const CLAVE = '';

const HOJA_PROV = 'APP_Proveedores';
const HOJA_PROD = 'APP_Productos';
const HOJA_PED = 'APP_Pedidos';

const COL_PROV = ['Proveedor', 'Contacto', 'Teléfono', 'WhatsApp', 'Entregas', 'Nota', 'Activo'];
const COL_PROD = ['ID', 'Proveedor', 'Sección', 'Producto', 'Unidad', 'Mínimo', 'Ref', 'Precio', 'Activo'];
const COL_PED = ['Fecha', 'Hora', 'Entrega', 'Proveedor', 'Producto', 'Cantidad', 'Unidad', 'ID producto', 'Pedido', 'Dispositivo'];

const CLAVES = {
  'Proveedor': 'proveedor', 'Contacto': 'contacto', 'Teléfono': 'telefono', 'WhatsApp': 'whatsapp',
  'Días de pedido': 'dias', 'Entregas': 'entregas', 'Nota': 'nota',
  'ID': 'id', 'Sección': 'seccion', 'Producto': 'producto', 'Unidad': 'unidad',
  'Habitual': 'habitual', 'Mínimo': 'habitual', 'Minimo': 'habitual', 'Ref': 'ref', 'Precio': 'precio', 'Activo': 'activo'
};

/* ---------- API para la app ---------- */

function doGet(e) {
  return salida({ ok: true, mensaje: 'Conexión de Pedidos activa. Usa la app para ver los datos.' });
}

function doPost(e) {
  let d;
  try { d = JSON.parse(e.postData.contents); } catch (err) { return salida({ ok: false, error: 'Datos no válidos' }); }
  if (CLAVE && d.clave !== CLAVE) return salida({ ok: false, error: 'Clave incorrecta' });
  if (d.accion === 'catalogo') {
    try { return salida(Object.assign({ ok: true }, leerCatalogo())); }
    catch (err) { return salida({ ok: false, error: String(err) }); }
  }
  if (d.accion === 'historial') {
    try { return salida({ ok: true, filas: leerHistorial(d.dias || 183) }); }
    catch (err) { return salida({ ok: false, error: String(err) }); }
  }
  if (d.accion !== 'pedido') return salida({ ok: false, error: 'Acción desconocida' });
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const hoja = hojaOCrear(HOJA_PED, COL_PED);
    // Evita duplicados si la app reintenta el mismo envío
    const ult = hoja.getLastRow();
    if (ult > 1) {
      const desde = Math.max(2, ult - 400);
      const ids = hoja.getRange(desde, 9, ult - desde + 1, 1).getValues().flat();
      if (ids.indexOf(d.pedido) >= 0) return salida({ ok: true, pedido: d.pedido, repetido: true });
    }
    const ahora = new Date(d.ts || Date.now());
    const filas = (d.items || []).map(it => [
      ahora, Utilities.formatDate(ahora, Session.getScriptTimeZone(), 'HH:mm'),
      d.entrega || '', d.proveedor, it.producto, Number(it.cantidad), it.unidad || '', it.id, d.pedido, d.dispositivo || ''
    ]);
    if (filas.length) {
      hoja.getRange(hoja.getLastRow() + 1, 1, filas.length, COL_PED.length).setValues(filas);
    }
    return salida({ ok: true, pedido: d.pedido, filas: filas.length });
  } finally {
    lock.releaseLock();
  }
}

function salida(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function leerTabla(nombre) {
  const hoja = SpreadsheetApp.getActive().getSheetByName(nombre);
  if (!hoja || hoja.getLastRow() < 2) return [];
  const datos = hoja.getDataRange().getDisplayValues();
  const cab = datos.shift().map(h => CLAVES[String(h).trim()] || String(h).trim());
  return datos
    .filter(f => f.some(c => String(c).trim() !== ''))
    .map(f => {
      const o = {};
      cab.forEach((k, i) => { o[k] = String(f[i] == null ? '' : f[i]).trim(); });
      return o;
    });
}

// Si falta alguna pestaña APP_ la crea vacía, solo con los títulos. Nunca borra nada.
function asegurarDatos() {
  const ss = SpreadsheetApp.getActive();
  if (!ss.getSheetByName(HOJA_PROV) || !ss.getSheetByName(HOJA_PROD) || !ss.getSheetByName(HOJA_PED)) configurar();
  ponerAlDia();
}

function reglaSiNo() {
  return SpreadsheetApp.newDataValidation().requireValueInList(['SI', 'NO'], true).build();
}

function cabecera(hoja) {
  return hoja.getRange(1, 1, 1, hoja.getLastColumn()).getDisplayValues()[0].map(h => String(h).trim());
}

// Pone al día las pestañas creadas con versiones anteriores y pone ID a los productos nuevos.
function ponerAlDia() {
  const ss = SpreadsheetApp.getActive();
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    // Proveedores: columna Activo (SI/NO) para quitarlos sin borrarlos
    const hp = ss.getSheetByName(HOJA_PROV);
    if (hp && cabecera(hp).indexOf('Activo') < 0) {
      const col = hp.getLastColumn() + 1;
      hp.getRange(1, col).setValue('Activo').setFontWeight('bold').setBackground('#1f5c45').setFontColor('#ffffff')
        .setNote('NO = el proveedor deja de salir en la app sin borrarlo.');
      if (hp.getLastRow() > 1) hp.getRange(2, col, hp.getLastRow() - 1, 1).setValue('SI');
      hp.getRange(2, col, 1000, 1).setDataValidation(reglaSiNo());
    }
    const hd = ss.getSheetByName(HOJA_PROD);
    if (!hd) return;
    const cab = cabecera(hd);
    // "Habitual" en la hoja es en realidad el mínimo que hay que tener
    const iHab = cab.indexOf('Habitual');
    if (iHab >= 0) hd.getRange(1, iHab + 1).setValue('Mínimo')
      .setNote('Lo mínimo que hay que tener. Ej: 8, 1/2 pieza, 2 bote. Lo habitual lo calcula la app con los pedidos.');
    // Productos sin ID (filas nuevas) o con el ID repetido (fila copiada): se les pone uno nuevo
    const iId = cab.indexOf('ID'), iProd = cab.indexOf('Producto'), n = hd.getLastRow() - 1;
    if (iId < 0 || iProd < 0 || n < 1) return;
    const ids = hd.getRange(2, iId + 1, n, 1).getDisplayValues().map(f => String(f[0]).trim());
    const nombres = hd.getRange(2, iProd + 1, n, 1).getDisplayValues().map(f => String(f[0]).trim());
    let max = ids.reduce((m, v) => { const k = /^p(\d+)$/i.exec(v); return k ? Math.max(m, +k[1]) : m; }, 0);
    const vistos = new Set();
    let cambio = false;
    const nuevos = ids.map((v, i) => {
      if (nombres[i] && (!v || vistos.has(v))) { v = 'p' + String(++max).padStart(3, '0'); cambio = true; }
      if (v) vistos.add(v);
      return [v];
    });
    if (cambio) hd.getRange(2, iId + 1, n, 1).setValues(nuevos);
  } finally {
    lock.releaseLock();
  }
}

function leerCatalogo() {
  asegurarDatos();
  const activo = x => String(x.activo || 'SI').trim().toUpperCase() !== 'NO';
  const proveedores = leerTabla(HOJA_PROV).filter(p => p.proveedor && activo(p));
  // El nombre del proveedor en productos se compara sin mayúsculas ni espacios de más
  const clave = s => String(s).trim().replace(/\s+/g, ' ').toUpperCase();
  const nombre = {};
  proveedores.forEach(p => { nombre[clave(p.proveedor)] = p.proveedor; });
  const productos = leerTabla(HOJA_PROD)
    .filter(p => p.producto && p.proveedor && activo(p) && nombre[clave(p.proveedor)])
    .map(p => Object.assign(p, { proveedor: nombre[clave(p.proveedor)] }));
  return { proveedores, productos, ultimos: ultimosPedidos(), actualizado: new Date().toISOString() };
}

// Último pedido de cada proveedor (y de cada día de entrega), para "Repetir último"
function ultimosPedidos() {
  const hoja = SpreadsheetApp.getActive().getSheetByName(HOJA_PED);
  if (!hoja || hoja.getLastRow() < 2) return {};
  const ult = hoja.getLastRow();
  const desde = Math.max(2, ult - 3000);
  const filas = hoja.getRange(desde, 1, ult - desde + 1, COL_PED.length).getValues();
  const res = {};
  for (let i = filas.length - 1; i >= 0; i--) {
    const f = filas[i];
    const clave = f[3] + '|' + (f[2] || '');
    if (!res[clave]) res[clave] = { pedido: f[8], fecha: f[0] instanceof Date ? f[0].toISOString() : String(f[0]), items: [] };
    if (res[clave].pedido === f[8]) res[clave].items.push({ id: f[7], producto: f[4], cantidad: f[5] });
  }
  return res;
}

/* ---------- Preparación de la hoja ---------- */

// Pedidos de los últimos N días, en formato compacto para la app:
// [fecha ISO, entrega, proveedor, producto, cantidad, unidad, id producto, nº pedido]
function leerHistorial(dias) {
  const hoja = SpreadsheetApp.getActive().getSheetByName(HOJA_PED);
  if (!hoja || hoja.getLastRow() < 2) return [];
  const limite = new Date(Date.now() - dias * 86400000);
  const filas = hoja.getRange(2, 1, hoja.getLastRow() - 1, COL_PED.length).getValues();
  const res = [];
  for (const f of filas) {
    const fecha = f[0] instanceof Date ? f[0] : new Date(f[0]);
    if (isNaN(fecha) || fecha < limite) continue;
    res.push([fecha.toISOString(), f[2], f[3], f[4], Number(f[5]) || 0, f[6], f[7], f[8]]);
  }
  return res;
}

function hojaOCrear(nombre, columnas) {
  const ss = SpreadsheetApp.getActive();
  let hoja = ss.getSheetByName(nombre);
  if (!hoja) {
    hoja = ss.insertSheet(nombre);
    hoja.getRange(1, 1, 1, columnas.length).setValues([columnas])
      .setFontWeight('bold').setBackground('#1f5c45').setFontColor('#ffffff');
    hoja.setFrozenRows(1);
  }
  return hoja;
}

function configurar() {
  const ss = SpreadsheetApp.getActive();
  const hechas = [];

  if (!ss.getSheetByName(HOJA_PROV)) {
    const h = hojaOCrear(HOJA_PROV, COL_PROV);
    h.getRange('D:D').setNumberFormat('@');
    h.getRange(2, 7, 1000, 1).setDataValidation(reglaSiNo());
    h.getRange('G1').setNote('NO = el proveedor deja de salir en la app sin borrarlo.');
    h.setColumnWidth(1, 160); h.setColumnWidth(3, 200); h.setColumnWidth(6, 320);
    h.getRange('E1').setNote('Días de entrega si el pedido se reparte en varios (ej. M,V). Vacío = una sola entrega');
    h.getRange('D1').setNote('Número con prefijo 34 y sin espacios, ej. 34616272201. Vacío = sin WhatsApp');
    hechas.push(HOJA_PROV);
  }

  if (!ss.getSheetByName(HOJA_PROD)) {
    const h = hojaOCrear(HOJA_PROD, COL_PROD);
    h.getRange('F:F').setNumberFormat('@');
    h.setColumnWidth(2, 150); h.setColumnWidth(4, 320);
    h.getRange(2, 9, 1000, 1).setDataValidation(reglaSiNo());
    h.getRange('A1').setNote('Para un producto nuevo deja el ID vacío: se pone solo al abrir la app.');
    h.getRange('F1').setNote('Lo mínimo que hay que tener. Ej: 8, 1/2 pieza, 2 bote. Lo habitual lo calcula la app con los pedidos.');
    h.getRange('I1').setNote('NO = el producto deja de salir en la app sin borrarlo.');
    hechas.push(HOJA_PROD);
  }

  if (!ss.getSheetByName(HOJA_PED)) {
    const h = hojaOCrear(HOJA_PED, COL_PED);
    h.getRange('A:A').setNumberFormat('dd/mm/yyyy');
    h.setColumnWidth(5, 300);
    hechas.push(HOJA_PED);
  }

  const msg = hechas.length ? 'Pestañas creadas: ' + hechas.join(', ') : 'Las pestañas ya existían; no he cambiado nada.';
  Logger.log(msg);
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) { /* llamado desde la app: sin ventana */ }
}
