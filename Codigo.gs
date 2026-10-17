/**
 * PEDIDOS A PROVEEDORES · Buenas y Santas
 * Conecta la app de la tablet (GitHub Pages) con esta hoja de Google.
 *
 * PASOS (solo la primera vez):
 *  1. Pega este código y guarda.
 *  2. Implementar → Nueva implementación → tipo "Aplicación web"
 *     · Ejecutar como: Yo   · Quién tiene acceso: Cualquier usuario
 *  No hace falta ejecutar nada: la primera vez que la app se conecta, se crean solas
 *  las pestañas APP_Proveedores, APP_Productos y APP_Pedidos. Tus pestañas no se tocan.
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

// Si faltan las pestañas APP_ (o están vacías) las crea solas con los datos de la pestaña "lunes".
function asegurarDatos() {
  const ss = SpreadsheetApp.getActive();
  [HOJA_PROV, HOJA_PROD].forEach(n => {
    const h = ss.getSheetByName(n);
    if (h && h.getLastRow() < 2) ss.deleteSheet(h);
  });
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
    const filas = DATOS.proveedores.map(p => [p.proveedor, p.contacto, p.telefono, p.whatsapp, p.entregas, p.nota, 'SI']);
    h.getRange(2, 1, filas.length, COL_PROV.length).setValues(filas);
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
    const filas = DATOS.productos.map(p => [p.id, p.proveedor, p.seccion, p.producto, p.unidad, p.habitual, p.ref, p.precio, p.activo]);
    h.getRange(2, 1, filas.length, COL_PROD.length).setValues(filas);
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

// Borra APP_Proveedores y APP_Productos y los vuelve a crear con los datos de este código.
// El historial de APP_Pedidos no se toca.
function reiniciarDatos() {
  const ss = SpreadsheetApp.getActive();
  [HOJA_PROV, HOJA_PROD].forEach(n => { const h = ss.getSheetByName(n); if (h) ss.deleteSheet(h); });
  configurar();
}

const DATOS = {
"proveedores": [
{
"proveedor": "MERCADONA",
"contacto": "",
"telefono": "",
"whatsapp": "",
"dias": "L",
"entregas": "M",
"nota": "Compra en tienda u online."
},
{
"proveedor": "ATLANTA",
"contacto": "",
"telefono": "911 425 093",
"whatsapp": "",
"dias": "L",
"entregas": "M,V",
"nota": ""
},
{
"proveedor": "IBEROLACTAL",
"contacto": "Víctor",
"telefono": "916 621 017 / 610 518 276",
"whatsapp": "34610518276",
"dias": "L",
"entregas": "M,J",
"nota": ""
},
{
"proveedor": "CBG",
"contacto": "Doménico",
"telefono": "661 559 248 / 972 70 33 70",
"whatsapp": "34661559248",
"dias": "L",
"entregas": "",
"nota": ""
},
{
"proveedor": "CREM DE LUX",
"contacto": "Paulo",
"telefono": "643 142 820",
"whatsapp": "34643142820",
"dias": "L",
"entregas": "M",
"nota": "Helados"
},
{
"proveedor": "BILANDERA",
"contacto": "Clara",
"telefono": "655 31 15 69",
"whatsapp": "34655311569",
"dias": "L",
"entregas": "",
"nota": "Mínimo 3 por envío."
},
{
"proveedor": "DUO HARINERO",
"contacto": "Sandra",
"telefono": "639 133 978",
"whatsapp": "34639133978",
"dias": "L",
"entregas": "",
"nota": "Pedido mínimo 80 €; si no, 7,26 € de portes."
},
{
"proveedor": "DISBESA",
"contacto": "Vicente",
"telefono": "916 499 300 / 617 48 15 65",
"whatsapp": "34617481565",
"dias": "L",
"entregas": "M,V",
"nota": ""
},
{
"proveedor": "BERLYS",
"contacto": "Cliente 177462",
"telefono": "914 842 079",
"whatsapp": "",
"dias": "L",
"entregas": "X,S",
"nota": ""
},
{
"proveedor": "HUGHES",
"contacto": "",
"telefono": "916 314 006 / 607 656 226",
"whatsapp": "34607656226",
"dias": "L",
"entregas": "M,V",
"nota": ""
},
{
"proveedor": "DELTA",
"contacto": "José Antonio",
"telefono": "617 48 87 92",
"whatsapp": "34617488792",
"dias": "L",
"entregas": "",
"nota": "Café Delta Gold: mínimo 5 bolsas."
},
{
"proveedor": "LOS FUENTEÑOS",
"contacto": "Isaac",
"telefono": "662 43 96 19",
"whatsapp": "34662439619",
"dias": "L",
"entregas": "",
"nota": "Vinos"
},
{
"proveedor": "WINEUP",
"contacto": "Carolina",
"telefono": "62 41 22 47 (incompleto)",
"whatsapp": "",
"dias": "L",
"entregas": "",
"nota": "Revisa el teléfono: en la hoja le falta un número."
},
{
"proveedor": "ENVAPRO",
"contacto": "",
"telefono": "91 126 50 15",
"whatsapp": "",
"dias": "L",
"entregas": "M",
"nota": "Limpieza y envases"
},
{
"proveedor": "VARIOS",
"contacto": "",
"telefono": "",
"whatsapp": "",
"dias": "L",
"entregas": "",
"nota": "Compras sueltas (Amazon…)"
}
],
"productos": [
{
"id": "p050",
"proveedor": "MERCADONA",
"seccion": "",
"producto": "leche semides Hacendado 1,5",
"unidad": "",
"habitual": "15",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p051",
"proveedor": "MERCADONA",
"seccion": "",
"producto": "agua c/gas grande Aguadoy1.5 lt",
"unidad": "",
"habitual": "12",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p052",
"proveedor": "MERCADONA",
"seccion": "",
"producto": "FANTA limón 2 lt",
"unidad": "",
"habitual": "4 unidades",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p053",
"proveedor": "MERCADONA",
"seccion": "",
"producto": "gaseosa LA CASERA grande",
"unidad": "",
"habitual": "3 u",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p054",
"proveedor": "MERCADONA",
"seccion": "",
"producto": "vino blanco Casón H",
"unidad": "",
"habitual": "6",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p055",
"proveedor": "MERCADONA",
"seccion": "",
"producto": "vino tinto Casón H",
"unidad": "",
"habitual": "6",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p056",
"proveedor": "MERCADONA",
"seccion": "",
"producto": "pan de pita original precocido",
"unidad": "",
"habitual": "5",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p057",
"proveedor": "MERCADONA",
"seccion": "",
"producto": "couscous mediano Bia 1k",
"unidad": "",
"habitual": "5c/menu",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p058",
"proveedor": "MERCADONA",
"seccion": "",
"producto": "garbanzos cocidos hacendado",
"unidad": "",
"habitual": "12",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p059",
"proveedor": "MERCADONA",
"seccion": "",
"producto": "maíz dulce hacen 3x140",
"unidad": "",
"habitual": "8c/menu",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p060",
"proveedor": "MERCADONA",
"seccion": "",
"producto": "spaguetti 1k hacendado",
"unidad": "",
"habitual": "10 c/menu",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p061",
"proveedor": "MERCADONA",
"seccion": "",
"producto": "Hélices con vegetales Hacendado 1k",
"unidad": "",
"habitual": "",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p062",
"proveedor": "MERCADONA",
"seccion": "",
"producto": "EL PAVO placas lasaña 200 gr",
"unidad": "",
"habitual": "",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p063",
"proveedor": "MERCADONA",
"seccion": "",
"producto": "Avena Molida Hacendado",
"unidad": "",
"habitual": "1",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p064",
"proveedor": "MERCADONA",
"seccion": "",
"producto": "Semillas de Chia",
"unidad": "",
"habitual": "1",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p065",
"proveedor": "MERCADONA",
"seccion": "",
"producto": "Lavavajillas Todo en 1",
"unidad": "",
"habitual": "2",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p066",
"proveedor": "MERCADONA",
"seccion": "",
"producto": "harina trigo hacndado 1 Kg",
"unidad": "",
"habitual": "8",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p067",
"proveedor": "MERCADONA",
"seccion": "",
"producto": "azúcar blanco 1 Kg",
"unidad": "",
"habitual": "8",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p068",
"proveedor": "MERCADONA",
"seccion": "",
"producto": "pan de molde SIN GLUTEN",
"unidad": "",
"habitual": "",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p069",
"proveedor": "MERCADONA",
"seccion": "",
"producto": "filetes de anchoa X3",
"unidad": "",
"habitual": "2",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p070",
"proveedor": "MERCADONA",
"seccion": "",
"producto": "Tónica original Schweppes 33 cl",
"unidad": "",
"habitual": "8",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p071",
"proveedor": "MERCADONA",
"seccion": "",
"producto": "SPRITE",
"unidad": "",
"habitual": "8",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p072",
"proveedor": "MERCADONA",
"seccion": "",
"producto": "Ambientador automatico",
"unidad": "",
"habitual": "2",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p073",
"proveedor": "MERCADONA",
"seccion": "",
"producto": "Lenteja cocida Hacendado 570g",
"unidad": "",
"habitual": "9",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p074",
"proveedor": "MERCADONA",
"seccion": "",
"producto": "bolsa de basura 30 l chica",
"unidad": "",
"habitual": "3",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p075",
"proveedor": "MERCADONA",
"seccion": "",
"producto": "Papel Multiusos Resist Bosque",
"unidad": "",
"habitual": "14",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p076",
"proveedor": "MERCADONA",
"seccion": "",
"producto": "mantequilla hacen s/sal 250 gr",
"unidad": "",
"habitual": "8",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p077",
"proveedor": "MERCADONA",
"seccion": "",
"producto": "mayonesa hacendado 500 ml",
"unidad": "",
"habitual": "4",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p078",
"proveedor": "MERCADONA",
"seccion": "",
"producto": "nata ligera para cocinar",
"unidad": "",
"habitual": "24",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p079",
"proveedor": "MERCADONA",
"seccion": "",
"producto": "bebida de soja",
"unidad": "",
"habitual": "8",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p080",
"proveedor": "MERCADONA",
"seccion": "",
"producto": "Dulce de Leche Chimbote",
"unidad": "",
"habitual": "7",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p081",
"proveedor": "MERCADONA",
"seccion": "",
"producto": "Yogur Griego",
"unidad": "",
"habitual": "",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p082",
"proveedor": "MERCADONA",
"seccion": "",
"producto": "Azucar Glas Hacendado 500grs",
"unidad": "",
"habitual": "2",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p083",
"proveedor": "MERCADONA",
"seccion": "",
"producto": "agua con gas 500 ml",
"unidad": "",
"habitual": "8 PACK",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p084",
"proveedor": "MERCADONA",
"seccion": "",
"producto": "Aceite de Oliva 3 litros",
"unidad": "",
"habitual": "3",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p085",
"proveedor": "MERCADONA",
"seccion": "",
"producto": "Margarina",
"unidad": "",
"habitual": "6",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p086",
"proveedor": "MERCADONA",
"seccion": "",
"producto": "Leche condensada",
"unidad": "",
"habitual": "4",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p087",
"proveedor": "MERCADONA",
"seccion": "",
"producto": "frutos rojos LA SIERNA",
"unidad": "",
"habitual": "8",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p088",
"proveedor": "MERCADONA",
"seccion": "",
"producto": "Mezcla de setas boletus",
"unidad": "",
"habitual": "4 bolsa",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p089",
"proveedor": "MERCADONA",
"seccion": "",
"producto": "Nuez troceda pelada",
"unidad": "",
"habitual": "",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p090",
"proveedor": "MERCADONA",
"seccion": "",
"producto": "Café Soluble",
"unidad": "",
"habitual": "",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p091",
"proveedor": "MERCADONA",
"seccion": "",
"producto": "Aquarius Limon",
"unidad": "",
"habitual": "12",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p092",
"proveedor": "MERCADONA",
"seccion": "",
"producto": "Aquarius Naranja",
"unidad": "",
"habitual": "12",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p093",
"proveedor": "ATLANTA",
"seccion": "",
"producto": "Caesar dressing 4 bote",
"unidad": "",
"habitual": "2 bote",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p094",
"proveedor": "ATLANTA",
"seccion": "",
"producto": "Croutons Italian 6 bolsas",
"unidad": "",
"habitual": "2 bol",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p095",
"proveedor": "ATLANTA",
"seccion": "",
"producto": "Palomitas de pollo 10 kg",
"unidad": "",
"habitual": "12 bol",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p096",
"proveedor": "ATLANTA",
"seccion": "",
"producto": "Patata dipper",
"unidad": "",
"habitual": "1 cajas",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p097",
"proveedor": "ATLANTA",
"seccion": "",
"producto": "Wrap de espinaca",
"unidad": "",
"habitual": "2 paquetes",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p098",
"proveedor": "ATLANTA",
"seccion": "",
"producto": "Falafel",
"unidad": "",
"habitual": "1 bolsa",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p099",
"proveedor": "ATLANTA",
"seccion": "",
"producto": "Maiz congelado 2,5kg",
"unidad": "",
"habitual": "1/2 cajón",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p100",
"proveedor": "IBEROLACTAL",
"seccion": "",
"producto": "Brie",
"unidad": "",
"habitual": "4 pieza",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p101",
"proveedor": "IBEROLACTAL",
"seccion": "",
"producto": "Cabra rulo",
"unidad": "",
"habitual": "3 piezas",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p102",
"proveedor": "IBEROLACTAL",
"seccion": "",
"producto": "Emmental rallado",
"unidad": "",
"habitual": "8 bolsas",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p103",
"proveedor": "IBEROLACTAL",
"seccion": "",
"producto": "Mezcla 70/30 rallado",
"unidad": "",
"habitual": "18 bolsas",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p104",
"proveedor": "IBEROLACTAL",
"seccion": "",
"producto": "Queso azul",
"unidad": "",
"habitual": "1/2 pieza",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p105",
"proveedor": "IBEROLACTAL",
"seccion": "",
"producto": "Escamas par Gran Moravia",
"unidad": "",
"habitual": "4 bol",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p106",
"proveedor": "IBEROLACTAL",
"seccion": "",
"producto": "Feta",
"unidad": "",
"habitual": "1/2 pieza",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p107",
"proveedor": "IBEROLACTAL",
"seccion": "",
"producto": "Supermix",
"unidad": "",
"habitual": "8",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p108",
"proveedor": "IBEROLACTAL",
"seccion": "",
"producto": "Mascarpone 1/2k",
"unidad": "",
"habitual": "2",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p109",
"proveedor": "IBEROLACTAL",
"seccion": "",
"producto": "Quescrem",
"unidad": "",
"habitual": "3",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p110",
"proveedor": "IBEROLACTAL",
"seccion": "",
"producto": "Nata para montar",
"unidad": "",
"habitual": "2",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p111",
"proveedor": "IBEROLACTAL",
"seccion": "",
"producto": "Yogurt 5 kg",
"unidad": "",
"habitual": "3",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p112",
"proveedor": "IBEROLACTAL",
"seccion": "",
"producto": "Chedar rojo",
"unidad": "",
"habitual": "1 paquetes",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p113",
"proveedor": "IBEROLACTAL",
"seccion": "",
"producto": "Aceitunas la sabrosita 4,25l",
"unidad": "",
"habitual": "2",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p114",
"proveedor": "IBEROLACTAL",
"seccion": "",
"producto": "Provolone dulce 5 kgs sin corteza",
"unidad": "",
"habitual": "",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p115",
"proveedor": "CBG",
"seccion": "",
"producto": "Arroz Arboreo",
"unidad": "",
"habitual": "8 paq.",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p116",
"proveedor": "CBG",
"seccion": "",
"producto": "Ravio Calabaza",
"unidad": "",
"habitual": "3 bolsas de 3 kg",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p117",
"proveedor": "CBG",
"seccion": "",
"producto": "Raviol ricotta menú",
"unidad": "",
"habitual": "15kg/menú",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p118",
"proveedor": "CBG",
"seccion": "",
"producto": "Ñoquis patata menú",
"unidad": "",
"habitual": "15kg/menú",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p119",
"proveedor": "CREM DE LUX",
"seccion": "",
"producto": "Dulce de leche Tango",
"unidad": "",
"habitual": "2",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p120",
"proveedor": "CREM DE LUX",
"seccion": "",
"producto": "Vainilla de gold",
"unidad": "",
"habitual": "2",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p121",
"proveedor": "CREM DE LUX",
"seccion": "",
"producto": "frutos rojos",
"unidad": "",
"habitual": "1",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p122",
"proveedor": "CREM DE LUX",
"seccion": "",
"producto": "Heldo menu",
"unidad": "",
"habitual": "",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p123",
"proveedor": "BILANDERA",
"seccion": "",
"producto": "Bailandera LUPULOVE",
"unidad": "",
"habitual": "2",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p124",
"proveedor": "DUO HARINERO",
"seccion": "",
"producto": "Choco cobert negro puro 58",
"unidad": "",
"habitual": "1/2",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p125",
"proveedor": "DUO HARINERO",
"seccion": "",
"producto": "Choco cober blanco puro 32",
"unidad": "",
"habitual": "1/2",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p126",
"proveedor": "DUO HARINERO",
"seccion": "",
"producto": "Papel siliconado",
"unidad": "",
"habitual": "1/4caja",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p127",
"proveedor": "DUO HARINERO",
"seccion": "",
"producto": "Azucar glass",
"unidad": "",
"habitual": "1/2 bolsa",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p128",
"proveedor": "DUO HARINERO",
"seccion": "",
"producto": "Cacao amargo",
"unidad": "",
"habitual": "1/2 bolsa",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p129",
"proveedor": "DISBESA",
"seccion": "Sala",
"producto": "Colacao sobre X 50",
"unidad": "",
"habitual": "8 sobres",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p130",
"proveedor": "DISBESA",
"seccion": "Sala",
"producto": "Miel antigoteo",
"unidad": "",
"habitual": "3",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p131",
"proveedor": "DISBESA",
"seccion": "Sala",
"producto": "Nescafe descaf",
"unidad": "",
"habitual": "8 sobres",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p132",
"proveedor": "DISBESA",
"seccion": "Sala",
"producto": "Azucar moreno en kg",
"unidad": "",
"habitual": "6",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p133",
"proveedor": "DISBESA",
"seccion": "Sala",
"producto": "San Miguel magna barril 30",
"unidad": "",
"habitual": "1",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p134",
"proveedor": "DISBESA",
"seccion": "Sala",
"producto": "San Miguel 00",
"unidad": "",
"habitual": "ver cajas",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p135",
"proveedor": "DISBESA",
"seccion": "Sala",
"producto": "San Miguel 1/3",
"unidad": "",
"habitual": "ver cajas",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p136",
"proveedor": "DISBESA",
"seccion": "Sala",
"producto": "San Miguel Selecta",
"unidad": "",
"habitual": "ver cajas",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p137",
"proveedor": "DISBESA",
"seccion": "Sala",
"producto": "San Migue sin gluten",
"unidad": "",
"habitual": "5",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p138",
"proveedor": "DISBESA",
"seccion": "Sala",
"producto": "San Miguel Sin TOSTADA",
"unidad": "",
"habitual": "",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p139",
"proveedor": "DISBESA",
"seccion": "Sala",
"producto": "Frutos secos 1 kg",
"unidad": "",
"habitual": "3",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p140",
"proveedor": "DISBESA",
"seccion": "Sala",
"producto": "Zumo GRANINI 330 ml",
"unidad": "",
"habitual": "5",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p141",
"proveedor": "DISBESA",
"seccion": "Sala",
"producto": "Toallas Z",
"unidad": "",
"habitual": "3",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p142",
"proveedor": "DISBESA",
"seccion": "Sala",
"producto": "COCA-COLA 24 lata",
"unidad": "",
"habitual": "24",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p143",
"proveedor": "DISBESA",
"seccion": "Sala",
"producto": "COCA-COLA LIGHT 24 lata",
"unidad": "",
"habitual": "12",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p144",
"proveedor": "DISBESA",
"seccion": "Sala",
"producto": "COCA-COLA ZERO 24 lata",
"unidad": "",
"habitual": "36",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p145",
"proveedor": "DISBESA",
"seccion": "Sala",
"producto": "COC-COLA ZEROZERO 24U",
"unidad": "",
"habitual": "24",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p146",
"proveedor": "DISBESA",
"seccion": "Sala",
"producto": "FANTA limón 24 latas",
"unidad": "",
"habitual": "12",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p147",
"proveedor": "DISBESA",
"seccion": "Sala",
"producto": "FANTA naranja 24 latas",
"unidad": "",
"habitual": "12",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p148",
"proveedor": "DISBESA",
"seccion": "Sala",
"producto": "NESTEA 24 latas",
"unidad": "",
"habitual": "18",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p149",
"proveedor": "DISBESA",
"seccion": "Sala",
"producto": "Agua Vidaqua",
"unidad": "",
"habitual": "6 cajas",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p150",
"proveedor": "DISBESA",
"seccion": "Sala",
"producto": "Ketchup Heinz sobre",
"unidad": "",
"habitual": "1 caja",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p151",
"proveedor": "DISBESA",
"seccion": "Sala",
"producto": "Mayonesa Heinz sobre",
"unidad": "",
"habitual": "1 caja",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p152",
"proveedor": "DISBESA",
"seccion": "Sala",
"producto": "Mostaza Heinz sobre",
"unidad": "",
"habitual": "1 caja",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p153",
"proveedor": "DISBESA",
"seccion": "Sala",
"producto": "Te Masala Chai Flor de T 20 u",
"unidad": "",
"habitual": "",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p154",
"proveedor": "DISBESA",
"seccion": "Sala",
"producto": "Te Frutos del bosque de T 20 u",
"unidad": "",
"habitual": "",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p155",
"proveedor": "DISBESA",
"seccion": "Sala",
"producto": "Te Ginger Limon de T 20 u",
"unidad": "",
"habitual": "",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p156",
"proveedor": "DISBESA",
"seccion": "Sala",
"producto": "Te Rojo pu-erh Flor de T 20 u",
"unidad": "",
"habitual": "",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p157",
"proveedor": "DISBESA",
"seccion": "Sala",
"producto": "Te Roibos Naranja de T 20 u",
"unidad": "",
"habitual": "",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p158",
"proveedor": "DISBESA",
"seccion": "Sala",
"producto": "Te Manzanilla Flor de T 100 u",
"unidad": "",
"habitual": "",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p159",
"proveedor": "DISBESA",
"seccion": "Sala",
"producto": "Te Poleo Flor de Tisana 100 u",
"unidad": "",
"habitual": "",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p160",
"proveedor": "DISBESA",
"seccion": "Sala",
"producto": "Te Tila Flor de Tisana 100 u",
"unidad": "",
"habitual": "",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p161",
"proveedor": "DISBESA",
"seccion": "Sala",
"producto": "Te Negro Flor de T 100 u",
"unidad": "",
"habitual": "",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p162",
"proveedor": "DISBESA",
"seccion": "Sala",
"producto": "Te Verde Flor de T 100 u",
"unidad": "",
"habitual": "",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p163",
"proveedor": "DISBESA",
"seccion": "Sala",
"producto": "Cacao Amargo",
"unidad": "",
"habitual": "",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p164",
"proveedor": "DISBESA",
"seccion": "Sala",
"producto": "Tabasco",
"unidad": "",
"habitual": "1",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p165",
"proveedor": "DISBESA",
"seccion": "Cocina",
"producto": "Aceituna negra s/hueso",
"unidad": "",
"habitual": "2",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p166",
"proveedor": "DISBESA",
"seccion": "Cocina",
"producto": "Amaretto",
"unidad": "",
"habitual": "1/2",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p167",
"proveedor": "DISBESA",
"seccion": "Cocina",
"producto": "Arroz grano largo 5kg",
"unidad": "",
"habitual": "1/2",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p168",
"proveedor": "DISBESA",
"seccion": "Cocina",
"producto": "Atun",
"unidad": "",
"habitual": "6",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p169",
"proveedor": "DISBESA",
"seccion": "Cocina",
"producto": "Caldo Avecrem verdu",
"unidad": "",
"habitual": "3",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p170",
"proveedor": "DISBESA",
"seccion": "Cocina",
"producto": "Maicena express",
"unidad": "",
"habitual": "1",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p171",
"proveedor": "DISBESA",
"seccion": "Cocina",
"producto": "Mostaza grande",
"unidad": "",
"habitual": "1",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p172",
"proveedor": "DISBESA",
"seccion": "Cocina",
"producto": "Pan rallado",
"unidad": "",
"habitual": "3 bolsa",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p173",
"proveedor": "DISBESA",
"seccion": "Cocina",
"producto": "Salsa de Soja S/GLUTEN",
"unidad": "",
"habitual": "2",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p174",
"proveedor": "DISBESA",
"seccion": "Cocina",
"producto": "Tomate frit 3k Gallina",
"unidad": "",
"habitual": "2",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p175",
"proveedor": "DISBESA",
"seccion": "Cocina",
"producto": "Uvas pasas",
"unidad": "",
"habitual": "1",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p176",
"proveedor": "DISBESA",
"seccion": "Cocina",
"producto": "Vinagre vino bco 2l",
"unidad": "",
"habitual": "1/2b",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p177",
"proveedor": "DISBESA",
"seccion": "Cocina",
"producto": "Sal Maldon",
"unidad": "",
"habitual": "1/4 bote",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p178",
"proveedor": "DISBESA",
"seccion": "Cocina",
"producto": "Tomate pomod Ginos",
"unidad": "",
"habitual": "1,5",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p179",
"proveedor": "DISBESA",
"seccion": "Cocina",
"producto": "Ajo granulado Dani",
"unidad": "",
"habitual": "1/2 b",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p180",
"proveedor": "DISBESA",
"seccion": "Cocina",
"producto": "Canela molida",
"unidad": "",
"habitual": "1/2 bote",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p181",
"proveedor": "DISBESA",
"seccion": "Cocina",
"producto": "Cebolla molida",
"unidad": "",
"habitual": "1/2 bote",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p182",
"proveedor": "DISBESA",
"seccion": "Cocina",
"producto": "Comino molido",
"unidad": "",
"habitual": "1/2 bote",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p183",
"proveedor": "DISBESA",
"seccion": "Cocina",
"producto": "Laurel en hojas",
"unidad": "",
"habitual": "1/2 bote",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p184",
"proveedor": "DISBESA",
"seccion": "Cocina",
"producto": "Nuez moscada",
"unidad": "",
"habitual": "1/2 bote",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p185",
"proveedor": "DISBESA",
"seccion": "Cocina",
"producto": "Oregano 150grs",
"unidad": "",
"habitual": "1",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p186",
"proveedor": "DISBESA",
"seccion": "Cocina",
"producto": "Perejil hoja 150grs",
"unidad": "",
"habitual": "",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p187",
"proveedor": "DISBESA",
"seccion": "Cocina",
"producto": "Pimentón dulce",
"unidad": "",
"habitual": "1 bote",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p188",
"proveedor": "DISBESA",
"seccion": "Cocina",
"producto": "Pimienta negra en grano",
"unidad": "",
"habitual": "",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p189",
"proveedor": "DISBESA",
"seccion": "Cocina",
"producto": "Pimienta negra molid",
"unidad": "",
"habitual": "1/2",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p190",
"proveedor": "DISBESA",
"seccion": "Cocina",
"producto": "Piñones",
"unidad": "",
"habitual": "1/2",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p191",
"proveedor": "DISBESA",
"seccion": "Cocina",
"producto": "Sesamo bote",
"unidad": "",
"habitual": "1/2 bote",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p192",
"proveedor": "DISBESA",
"seccion": "Cocina",
"producto": "Levadura Promolac",
"unidad": "",
"habitual": "1",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p193",
"proveedor": "DISBESA",
"seccion": "Cocina",
"producto": "Salsa Tartufato",
"unidad": "",
"habitual": "2",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p194",
"proveedor": "DISBESA",
"seccion": "Cocina",
"producto": "crema balsamica",
"unidad": "",
"habitual": "4",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p195",
"proveedor": "DISBESA",
"seccion": "Cocina",
"producto": "Girasol altoleico 25lt",
"unidad": "",
"habitual": "1",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p196",
"proveedor": "DISBESA",
"seccion": "Cocina",
"producto": "Girasol 5 l",
"unidad": "",
"habitual": "4 bot",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p197",
"proveedor": "DISBESA",
"seccion": "Licores",
"producto": "Absolut",
"unidad": "",
"habitual": "1",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p198",
"proveedor": "DISBESA",
"seccion": "Licores",
"producto": "Bacardi",
"unidad": "",
"habitual": "1",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p199",
"proveedor": "DISBESA",
"seccion": "Licores",
"producto": "Ballantaines",
"unidad": "",
"habitual": "1",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p200",
"proveedor": "DISBESA",
"seccion": "Licores",
"producto": "Beefeter",
"unidad": "",
"habitual": "2",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p201",
"proveedor": "DISBESA",
"seccion": "Licores",
"producto": "Bombay",
"unidad": "",
"habitual": "2",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p202",
"proveedor": "DISBESA",
"seccion": "Licores",
"producto": "Brugal",
"unidad": "",
"habitual": "2",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p203",
"proveedor": "DISBESA",
"seccion": "Licores",
"producto": "Havana club 3 años",
"unidad": "",
"habitual": "1",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p204",
"proveedor": "DISBESA",
"seccion": "Licores",
"producto": "J&B",
"unidad": "",
"habitual": "1",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p205",
"proveedor": "DISBESA",
"seccion": "Licores",
"producto": "Johonnie Walker red",
"unidad": "",
"habitual": "1",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p206",
"proveedor": "DISBESA",
"seccion": "Licores",
"producto": "Jose cuervo",
"unidad": "",
"habitual": "1",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p207",
"proveedor": "DISBESA",
"seccion": "Licores",
"producto": "Fernet Branca",
"unidad": "",
"habitual": "1",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p208",
"proveedor": "DISBESA",
"seccion": "Licores",
"producto": "Martini Blanco",
"unidad": "",
"habitual": "1",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p209",
"proveedor": "DISBESA",
"seccion": "Licores",
"producto": "Martini Rojo",
"unidad": "",
"habitual": "1",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p210",
"proveedor": "DISBESA",
"seccion": "Licores",
"producto": "Martini Rosato",
"unidad": "",
"habitual": "1",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p211",
"proveedor": "DISBESA",
"seccion": "Licores",
"producto": "Bailys",
"unidad": "",
"habitual": "2",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p212",
"proveedor": "DISBESA",
"seccion": "Licores",
"producto": "Chincon dulce",
"unidad": "",
"habitual": "2",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p213",
"proveedor": "DISBESA",
"seccion": "Licores",
"producto": "Lemoncello Toschi",
"unidad": "",
"habitual": "2",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p214",
"proveedor": "DISBESA",
"seccion": "Licores",
"producto": "Rua Vieja Crema de Orujo",
"unidad": "",
"habitual": "2",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p215",
"proveedor": "DISBESA",
"seccion": "Licores",
"producto": "Rua Vieja hierbas",
"unidad": "",
"habitual": "2",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p216",
"proveedor": "DISBESA",
"seccion": "Licores",
"producto": "Rua Vieja aguradiente",
"unidad": "",
"habitual": "2",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p217",
"proveedor": "DISBESA",
"seccion": "Licores",
"producto": "Patxaran Etxeko",
"unidad": "",
"habitual": "2",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p218",
"proveedor": "DISBESA",
"seccion": "Licores",
"producto": "Veterano",
"unidad": "",
"habitual": "1",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p219",
"proveedor": "BERLYS",
"seccion": "",
"producto": "Base hojaldre 40x60",
"unidad": "",
"habitual": "10 placas",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p220",
"proveedor": "BERLYS",
"seccion": "",
"producto": "Hamburguesa rustico",
"unidad": "",
"habitual": "4",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p221",
"proveedor": "BERLYS",
"seccion": "",
"producto": "Mini Diamante blanco",
"unidad": "",
"habitual": "2x",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p222",
"proveedor": "BERLYS",
"seccion": "",
"producto": "Diamante maiz",
"unidad": "",
"habitual": "2x",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p223",
"proveedor": "BERLYS",
"seccion": "",
"producto": "Diamante semillas",
"unidad": "",
"habitual": "2x",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p224",
"proveedor": "BERLYS",
"seccion": "",
"producto": "Diamante integral",
"unidad": "",
"habitual": "2x",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p225",
"proveedor": "BERLYS",
"seccion": "",
"producto": "Sin gluten hamburguesas",
"unidad": "",
"habitual": "",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p226",
"proveedor": "HUGHES",
"seccion": "",
"producto": "Tapas emp p/horno GDE",
"unidad": "",
"habitual": "2 caja",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p227",
"proveedor": "HUGHES",
"seccion": "",
"producto": "Aji molido",
"unidad": "",
"habitual": "1/4 bote",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p228",
"proveedor": "HUGHES",
"seccion": "",
"producto": "Yerba krauss",
"unidad": "",
"habitual": "",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p229",
"proveedor": "DELTA",
"seccion": "",
"producto": "Café Delta Gold min 5 bolsas",
"unidad": "",
"habitual": "5",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p230",
"proveedor": "DELTA",
"seccion": "",
"producto": "Café Descaf. 1/4",
"unidad": "",
"habitual": "4",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p231",
"proveedor": "DELTA",
"seccion": "",
"producto": "Azucar Blanca SOBRE",
"unidad": "",
"habitual": "1",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p232",
"proveedor": "DELTA",
"seccion": "",
"producto": "Azucar Morena SOBRE",
"unidad": "",
"habitual": "1",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p233",
"proveedor": "DELTA",
"seccion": "",
"producto": "Edulcorante SOBRE",
"unidad": "",
"habitual": "1",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p234",
"proveedor": "LOS FUENTEÑOS",
"seccion": "",
"producto": "Lagrimas de Maria Rioja",
"unidad": "",
"habitual": "12",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p235",
"proveedor": "LOS FUENTEÑOS",
"seccion": "",
"producto": "Ramon Bilbao crianza",
"unidad": "",
"habitual": "12",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p236",
"proveedor": "LOS FUENTEÑOS",
"seccion": "",
"producto": "Baden Numen Ribera",
"unidad": "",
"habitual": "12",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p237",
"proveedor": "LOS FUENTEÑOS",
"seccion": "",
"producto": "Dos de Mayo Madrid",
"unidad": "",
"habitual": "6",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p238",
"proveedor": "LOS FUENTEÑOS",
"seccion": "",
"producto": "Honoro Vera Organico",
"unidad": "",
"habitual": "6",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p239",
"proveedor": "LOS FUENTEÑOS",
"seccion": "",
"producto": "Rivasanz Rueda",
"unidad": "",
"habitual": "12",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p240",
"proveedor": "LOS FUENTEÑOS",
"seccion": "",
"producto": "Godello",
"unidad": "",
"habitual": "12",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p241",
"proveedor": "WINEUP",
"seccion": "",
"producto": "Quilmes",
"unidad": "",
"habitual": "",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p242",
"proveedor": "WINEUP",
"seccion": "",
"producto": "Trumpeter Malbec",
"unidad": "",
"habitual": "",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p243",
"proveedor": "WINEUP",
"seccion": "",
"producto": "Alfajores 12 u negros",
"unidad": "",
"habitual": "2",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p244",
"proveedor": "WINEUP",
"seccion": "",
"producto": "Alfajores 12 u mixtos",
"unidad": "",
"habitual": "3",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p245",
"proveedor": "WINEUP",
"seccion": "",
"producto": "Alfajores 12 u blancos",
"unidad": "",
"habitual": "",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p246",
"proveedor": "ENVAPRO",
"seccion": "",
"producto": "Desengrasante planc MARRON 10 l",
"unidad": "",
"habitual": "1",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p247",
"proveedor": "ENVAPRO",
"seccion": "",
"producto": "Detergente manual x 5l",
"unidad": "",
"habitual": "3",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p248",
"proveedor": "ENVAPRO",
"seccion": "",
"producto": "Abrillantador maquinas 5 l",
"unidad": "",
"habitual": "1",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p249",
"proveedor": "ENVAPRO",
"seccion": "",
"producto": "Bidon limpiacristales 5 l",
"unidad": "",
"habitual": "1/2",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p250",
"proveedor": "ENVAPRO",
"seccion": "",
"producto": "Gel de mano 5l",
"unidad": "",
"habitual": "1/2",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p251",
"proveedor": "ENVAPRO",
"seccion": "",
"producto": "Lavajillas maquina 5 l UN 1760",
"unidad": "",
"habitual": "2",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p252",
"proveedor": "ENVAPRO",
"seccion": "",
"producto": "Royo bayeta absorbente",
"unidad": "",
"habitual": "",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p253",
"proveedor": "ENVAPRO",
"seccion": "",
"producto": "Fibra verde c/esponja",
"unidad": "",
"habitual": "2",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p254",
"proveedor": "ENVAPRO",
"seccion": "",
"producto": "Estropajo acero Nanax",
"unidad": "",
"habitual": "3",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p255",
"proveedor": "ENVAPRO",
"seccion": "",
"producto": "Guantes nitrilo (L)",
"unidad": "",
"habitual": "1",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p256",
"proveedor": "ENVAPRO",
"seccion": "",
"producto": "Guantes de fregar (M)",
"unidad": "",
"habitual": "2",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p257",
"proveedor": "ENVAPRO",
"seccion": "",
"producto": "Cepillo (escoba)",
"unidad": "",
"habitual": "2",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p258",
"proveedor": "ENVAPRO",
"seccion": "",
"producto": "Fregonas algodón x 5",
"unidad": "",
"habitual": "2",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p259",
"proveedor": "ENVAPRO",
"seccion": "",
"producto": "Lejia diluida x 15",
"unidad": "",
"habitual": "4",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p260",
"proveedor": "ENVAPRO",
"seccion": "",
"producto": "Bol. Basura 85X105",
"unidad": "",
"habitual": "3",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p261",
"proveedor": "ENVAPRO",
"seccion": "",
"producto": "Bobina aluminio (30 cm)",
"unidad": "",
"habitual": "1/2",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p262",
"proveedor": "ENVAPRO",
"seccion": "",
"producto": "Bobina film (30 cm)",
"unidad": "",
"habitual": "1",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p263",
"proveedor": "ENVAPRO",
"seccion": "",
"producto": "Bobina film (45 cm)",
"unidad": "",
"habitual": "1",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p264",
"proveedor": "ENVAPRO",
"seccion": "",
"producto": "Papel higienico",
"unidad": "",
"habitual": "6",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p265",
"proveedor": "ENVAPRO",
"seccion": "",
"producto": "Servilleta 1/6 30x40 253SER52109077",
"unidad": "",
"habitual": "10",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p266",
"proveedor": "ENVAPRO",
"seccion": "",
"producto": "Bolsa pap C/asa 243BKDSKT18",
"unidad": "",
"habitual": "",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p267",
"proveedor": "ENVAPRO",
"seccion": "",
"producto": "Bolsa papel kraft (89143BK)",
"unidad": "",
"habitual": "",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p268",
"proveedor": "ENVAPRO",
"seccion": "",
"producto": "Caja de pizza ch 26(191P2727K)",
"unidad": "",
"habitual": "10",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p269",
"proveedor": "ENVAPRO",
"seccion": "",
"producto": "Caja de pizza GR 33 (243PKK3333)",
"unidad": "",
"habitual": "10",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p270",
"proveedor": "ENVAPRO",
"seccion": "",
"producto": "Cont. Porex GR (171P1654)",
"unidad": "",
"habitual": "1 pack",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p271",
"proveedor": "ENVAPRO",
"seccion": "",
"producto": "Cont porex CH (171P1476)",
"unidad": "",
"habitual": "1/2 pack",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p272",
"proveedor": "ENVAPRO",
"seccion": "",
"producto": "Ensaladera pet (225EV0514P)",
"unidad": "",
"habitual": "1/2 pa",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p273",
"proveedor": "ENVAPRO",
"seccion": "",
"producto": "Comandero 3617708 (doble)",
"unidad": "",
"habitual": "5",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p274",
"proveedor": "ENVAPRO",
"seccion": "",
"producto": "Plato carton (92BE02P)",
"unidad": "",
"habitual": "",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p275",
"proveedor": "ENVAPRO",
"seccion": "",
"producto": "Servilleta CH (153PN232207)",
"unidad": "",
"habitual": "5 pack",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p276",
"proveedor": "ENVAPRO",
"seccion": "",
"producto": "VASO CARTÓN EFECTO KRAFT 12-14 oz",
"unidad": "",
"habitual": "",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p277",
"proveedor": "ENVAPRO",
"seccion": "",
"producto": "VASO TÉRMICO POREX BLA 7 oz 200 ml",
"unidad": "",
"habitual": "",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p278",
"proveedor": "ENVAPRO",
"seccion": "",
"producto": "TAPA PLANA PLÁST VASO TÉRMICO 7oz",
"unidad": "",
"habitual": "",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p279",
"proveedor": "ENVAPRO",
"seccion": "",
"producto": "PERGAMINO ANTIGRAS 31X28 89R2831B",
"unidad": "",
"habitual": "",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p280",
"proveedor": "ENVAPRO",
"seccion": "",
"producto": "pajitas",
"unidad": "",
"habitual": "",
"ref": "",
"precio": "",
"activo": "SI"
},
{
"id": "p281",
"proveedor": "VARIOS",
"seccion": "",
"producto": "Bayeta punto azul plancha (amazon)",
"unidad": "",
"habitual": "",
"ref": "",
"precio": "",
"activo": "SI"
}
]
};
