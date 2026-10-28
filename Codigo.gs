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
const COL_PED = ['Fecha', 'Hora', 'Entrega', 'Proveedor', 'Producto', 'Cantidad', 'Unidad', 'ID producto', 'Pedido', 'Dispositivo',
  'Precio aprox.', 'Importe aprox.'];

const CLAVES = {
  'Proveedor': 'proveedor', 'Contacto': 'contacto', 'Teléfono': 'telefono', 'WhatsApp': 'whatsapp',
  'Días de pedido': 'dias', 'Entregas': 'entregas', 'Nota': 'nota',
  'ID': 'id', 'Sección': 'seccion', 'Producto': 'producto', 'Unidad': 'unidad',
  'Habitual': 'habitual', 'Mínimo': 'habitual', 'Minimo': 'habitual', 'Ref': 'ref', 'Precio': 'precio', 'Activo': 'activo',
  'Pedido mínimo': 'pedidoMinimo', 'Pedido minimo': 'pedidoMinimo'
};

/* ---------- API para la app ---------- */

function doGet(e) {
  return salida({ ok: true, mensaje: 'Conexión de Pedidos activa. Usa la app para ver los datos.' });
}

function doPost(e) {
  let d;
  try { d = JSON.parse(e.postData.contents); } catch (err) { return salida({ ok: false, error: 'Datos no válidos' }); }
  if (CLAVE && d.clave !== CLAVE) return salida({ ok: false, error: 'Clave incorrecta' });
  if (d.claveServidor !== CLAVE_PEDIDOS && !accesoValido(d.acceso)) {
    return salida({ ok: false, sinSesion: true, error: 'Hay que entrar con una cuenta autorizada de Buenas y Santas.' });
  }
  if (d.accion === 'catalogo') {
    try { return salida(Object.assign({ ok: true }, leerCatalogo())); }
    catch (err) { return salida({ ok: false, error: String(err) }); }
  }
  if (d.accion === 'historial') {
    try { return salida({ ok: true, filas: leerHistorial(d.dias || 183) }); }
    catch (err) { return salida({ ok: false, error: String(err) }); }
  }
  if (d.accion === 'precios') {
    try { return salida(guardarPrecios(d)); }
    catch (err) { return salida({ ok: false, error: String(err) }); }
  }
  if (d.accion === 'facturasConPrecios') {
    try { return salida({ ok: true, facturas: facturasConPrecios() }); }
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
    const filas = (d.items || []).map(it => {
      const precio = numero(it.precio);
      return [ahora, Utilities.formatDate(ahora, Session.getScriptTimeZone(), 'HH:mm'),
        d.entrega || '', d.proveedor, it.producto, Number(it.cantidad), it.unidad || '', it.id, d.pedido, d.dispositivo || '',
        precio || '', precio ? Math.round(precio * Number(it.cantidad) * 100) / 100 : ''];
    });
    if (filas.length) {
      hoja.getRange(hoja.getLastRow() + 1, 1, filas.length, COL_PED.length).setValues(filas);
    }
    return salida({ ok: true, pedido: d.pedido, filas: filas.length });
  } finally {
    lock.releaseLock();
  }
}

/* ---------- Acceso ----------
 * Solo entran las cuentas de Google autorizadas. La app inicia sesión con Google en el script
 * del Cierre de caja (que comprueba la cuenta) y recibe un "acceso" ("email|vence|firma")
 * firmado con CLAVE_PEDIDOS (archivo ClavePedidos.js, que no está en GitHub). Aquí solo se
 * comprueba la firma y que no haya vencido. El Cierre de caja, cuando llama a este script,
 * manda la misma clave como claveServidor. */
function accesoValido(acceso) {
  const p = String(acceso || '').split('|');
  if (p.length !== 3 || !(Number(p[1]) > Date.now())) return false;
  const firma = Utilities.computeHmacSha256Signature(p[0] + '|' + p[1], CLAVE_PEDIDOS)
    .map(b => ('0' + (b & 255).toString(16)).slice(-2)).join('');
  return firma === p[2];
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
  cambioUnaVez_ordenMonbake();
}

// UNA SOLA VEZ (19/10/2026): Berlys pasa a llamarse Monbake en todas las pestañas APP_ y los
// proveedores se ordenan como pidió el usuario. Se apunta en las propiedades del script para
// no repetirlo. Después se puede borrar esta función.
function cambioUnaVez_ordenMonbake() {
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty('orden_monbake') === 'hecho') return;
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    if (props.getProperty('orden_monbake') === 'hecho') return;
    const ss = SpreadsheetApp.getActive();
    const renombrar = (nombreHoja, columna) => {
      const h = ss.getSheetByName(nombreHoja);
      if (!h || h.getLastRow() < 2) return;
      const r = h.getRange(2, columna, h.getLastRow() - 1, 1), v = r.getValues();
      if (v.some(f => String(f[0]).trim().toUpperCase() === 'BERLYS'))
        r.setValues(v.map(f => [String(f[0]).trim().toUpperCase() === 'BERLYS' ? 'MONBAKE' : f[0]]));
    };
    const hp = ss.getSheetByName(HOJA_PROV);
    renombrar(HOJA_PROV, 1);
    renombrar(HOJA_PROD, cabecera(ss.getSheetByName(HOJA_PROD)).indexOf('Proveedor') + 1);
    renombrar(HOJA_PED, 4);
    renombrar(HOJA_PRECIOS, 2);
    // Orden: estos primero, CBG, Bilandera y Duo Harinero al final, el resto en medio
    const primeros = ['MERCADONA', 'ATLANTA', 'IBEROLACTAL', 'DISBESA', 'MONBAKE', 'HUGHES', 'ENVAPRO'];
    const ultimos = ['CBG', 'BILANDERA', 'DUO HARINERO'];
    const n = hp.getLastRow() - 1, nc = hp.getLastColumn();
    if (n > 1) {
      const filas = hp.getRange(2, 1, n, nc).getValues();
      const pos = f => { const k = String(f[0]).trim().toUpperCase();
        return primeros.indexOf(k) >= 0 ? primeros.indexOf(k) : ultimos.indexOf(k) >= 0 ? 1000 + ultimos.indexOf(k) : 500; };
      const ordenadas = filas.map((f, i) => ({ f, i })).sort((a, b) => pos(a.f) - pos(b.f) || a.i - b.i).map(x => x.f);
      hp.getRange(2, 1, n, nc).setValues(ordenadas);
    }
    props.setProperty('orden_monbake', 'hecho');
  } finally {
    lock.releaseLock();
  }
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
    // Proveedores: columna "Pedido mínimo" en euros (la app avisa si el pedido no llega)
    if (hp && cabecera(hp).indexOf('Pedido mínimo') < 0) {
      const col = hp.getLastColumn() + 1;
      hp.getRange(1, col).setValue('Pedido mínimo').setFontWeight('bold').setBackground('#1f5c45').setFontColor('#ffffff')
        .setNote('Pedido mínimo en euros, sin IVA (ej. 80). Vacío = no tiene. La app avisa si el pedido no llega.');
      const nombres = hp.getLastRow() > 1 ? hp.getRange(2, 1, hp.getLastRow() - 1, 1).getValues() : [];
      nombres.forEach((f, i) => { if (String(f[0]).trim().toUpperCase() === 'DUO HARINERO') hp.getRange(i + 2, col).setValue(80); });
    }
    // Pedidos: precio e importe aproximados de cada línea
    const hped = ss.getSheetByName(HOJA_PED);
    if (hped && cabecera(hped).indexOf('Importe aprox.') < 0) {
      hped.getRange(1, 11, 1, 2).setValues([['Precio aprox.', 'Importe aprox.']])
        .setFontWeight('bold').setBackground('#1f5c45').setFontColor('#ffffff');
      hped.getRange(1, 12).setNote('Cantidad × último precio conocido (de las facturas), sin IVA. Aproximado.');
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
  const proveedores = leerTabla(HOJA_PROV).filter(p => p.proveedor && activo(p))
    .map(p => Object.assign(p, { pedidoMinimo: numero(p.pedidoMinimo) || '' }));
  // El nombre del proveedor en productos se compara sin mayúsculas ni espacios de más
  const clave = s => String(s).trim().replace(/\s+/g, ' ').toUpperCase();
  const nombre = {};
  proveedores.forEach(p => { nombre[clave(p.proveedor)] = p.proveedor; });
  const productos = leerTabla(HOJA_PROD)
    .filter(p => p.producto && p.proveedor && activo(p) && nombre[clave(p.proveedor)])
    .map(p => Object.assign(p, { proveedor: nombre[clave(p.proveedor)], precio: numero(p.precio) || '' }));
  return { proveedores, productos, ultimos: ultimosPedidos(), cambios: cambiosDePrecio(60), actualizado: new Date().toISOString() };
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

/* ---------- Precios (leídos de las facturas por la app de Cierre de caja) ---------- */
// La app de Cierre de caja, al escanear una factura, manda aquí sus líneas ya emparejadas con
// los productos: { accion:'precios', proveedor, factura, fecha:'AAAA-MM-DD', foto, lineas:[
//   { texto, id, cantidad, unidad, precio } ] }. Se apuntan en APP_Precios y el último precio
// de cada producto pasa a la columna Precio de APP_Productos. Precios sin IVA.

const HOJA_PRECIOS = 'APP_Precios';
const COL_PRECIOS = ['Fecha', 'Proveedor', 'Factura', 'En la factura', 'ID producto', 'Producto', 'Cantidad',
  'Unidad', 'Precio', 'Precio anterior', 'Cambio %', 'Foto', 'Registrado'];

// "3,20 €", "3.2" o 3.2 → 3.2
function numero(v) {
  if (typeof v === 'number') return v;
  let s = String(v || '').replace(/[^\d,.\-]/g, '');
  if (s.indexOf(',') >= 0) s = s.replace(/\./g, '').replace(',', '.');
  const n = parseFloat(s);
  return isNaN(n) ? 0 : n;
}

function fechaISO(v) {
  if (v instanceof Date) return Utilities.formatDate(v, 'Europe/Madrid', 'yyyy-MM-dd');
  return String(v || '').slice(0, 10);
}

function guardarPrecios(d) {
  const lineas = (d.lineas || []).filter(l => numero(l.precio) > 0);
  if (!d.proveedor || !lineas.length) return { ok: true, filas: 0 };
  const fecha = /^\d{4}-\d{2}-\d{2}$/.test(String(d.fecha || '')) ? d.fecha : fechaISO(new Date());
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const ss = SpreadsheetApp.getActive();
    const hoja = hojaOCrear(HOJA_PRECIOS, COL_PRECIOS);
    let previas = hoja.getLastRow() > 1 ? hoja.getRange(2, 1, hoja.getLastRow() - 1, COL_PRECIOS.length).getValues() : [];
    // La misma factura escaneada otra vez no se apunta dos veces
    const fotoId = idDrive(d.foto);
    const esta = r => (fotoId && idDrive(r[11]) === fotoId) ||
      (r[1] === d.proveedor && (d.factura ? String(r[2]) === String(d.factura) : !r[2] && fechaISO(r[0]) === fecha));
    // "Volver a leer" (botón de la app): se quitan las líneas que había de esa factura
    if (d.reemplazar) {
      const quedan = previas.filter(r => !esta(r));
      if (quedan.length < previas.length) {
        // Se reescriben las que quedan (borrar todas las filas de golpe no se puede)
        hoja.getRange(2, 1, previas.length, COL_PRECIOS.length).clearContent();
        if (quedan.length) hoja.getRange(2, 1, quedan.length, COL_PRECIOS.length).setValues(quedan);
      }
      previas = quedan;
    }
    if (previas.some(esta)) return { ok: true, filas: 0, repetida: true };

    // Productos de la hoja y su fila, y el último precio conocido de cada uno
    const hp = ss.getSheetByName(HOJA_PROD), cab = cabecera(hp);
    const iId = cab.indexOf('ID'), iProd = cab.indexOf('Producto'), iPrecio = cab.indexOf('Precio');
    const datos = hp.getRange(2, 1, Math.max(1, hp.getLastRow() - 1), cab.length).getValues();
    const prod = {};
    datos.forEach((f, i) => { const id = String(f[iId]).trim(); if (id) prod[id] = { fila: i + 2, nombre: f[iProd], precio: numero(f[iPrecio]) }; });
    const ultimo = {};   // id → { fecha, precio }
    previas.forEach(r => { const id = String(r[4]); const p = numero(r[8]); const f = fechaISO(r[0]);
      if (id && p > 0 && (!ultimo[id] || f >= ultimo[id].fecha)) ultimo[id] = { fecha: f, precio: p }; });

    const ahora = new Date(), filas = [];
    for (const l of lineas) {
      const id = prod[String(l.id || '').trim()] ? String(l.id).trim() : '';
      const precio = Math.round(numero(l.precio) * 10000) / 10000;
      const antes = id ? (ultimo[id] ? ultimo[id].precio : prod[id].precio) : 0;
      const cambio = antes ? Math.round((precio - antes) / antes * 1000) / 10 : '';
      filas.push([new Date(fecha + 'T12:00:00'), d.proveedor, d.factura || '', String(l.texto || ''), id, id ? prod[id].nombre : '',
        numero(l.cantidad) || '', String(l.unidad || ''), precio, antes || '', cambio, d.foto || '', ahora]);
      // El precio de la hoja de productos es siempre el de la factura más reciente
      if (id && iPrecio >= 0 && (!ultimo[id] || fecha >= ultimo[id].fecha)) {
        hp.getRange(prod[id].fila, iPrecio + 1).setValue(precio);
        ultimo[id] = { fecha, precio };
      }
    }
    hoja.getRange(hoja.getLastRow() + 1, 1, filas.length, COL_PRECIOS.length).setValues(filas);
    hoja.getRange(2, 1, hoja.getLastRow() - 1, 1).setNumberFormat('dd/mm/yyyy');
    return { ok: true, filas: filas.length, emparejadas: filas.filter(f => f[4]).length };
  } finally {
    lock.releaseLock();
  }
}

// ID del archivo de Drive de un enlace de factura (".../file/d/<ID>/view")
function idDrive(url) {
  const m = /\/d\/([\w-]{20,})/.exec(String(url || '')) || /[?&]id=([\w-]{20,})/.exec(String(url || ''));
  return m ? m[1] : '';
}

// Facturas que ya tienen sus precios apuntados (para el botón "Leer precios de una factura",
// que lista las facturas de Drive desde la app de Cierre de caja): [{ proveedor, factura, fotoId, filas }]
function facturasConPrecios() {
  const hoja = SpreadsheetApp.getActive().getSheetByName(HOJA_PRECIOS);
  if (!hoja || hoja.getLastRow() < 2) return [];
  const res = {};
  for (const r of hoja.getRange(2, 1, hoja.getLastRow() - 1, COL_PRECIOS.length).getValues()) {
    const fotoId = idDrive(r[11]), k = fotoId || r[1] + '|' + r[2];
    (res[k] = res[k] || { proveedor: r[1], factura: String(r[2]), fotoId, filas: 0 }).filas++;
  }
  return Object.values(res);
}

// Último cambio de precio de cada producto en los últimos N días (para el aviso de la app)
function cambiosDePrecio(dias) {
  const hoja = SpreadsheetApp.getActive().getSheetByName(HOJA_PRECIOS);
  if (!hoja || hoja.getLastRow() < 2) return [];
  const limite = fechaISO(new Date(Date.now() - dias * 86400000));
  const filas = hoja.getRange(2, 1, hoja.getLastRow() - 1, COL_PRECIOS.length).getValues();
  const res = {};
  for (const r of filas) {
    const id = String(r[4]), f = fechaISO(r[0]);
    if (!id || f < limite) continue;
    if (!res[id] || f >= res[id].fecha) res[id] = { id, producto: r[5], proveedor: r[1], fecha: f,
      antes: numero(r[9]), ahora: numero(r[8]), cambio: r[10] === '' ? 0 : Number(r[10]), factura: String(r[2]), foto: String(r[11]) };
  }
  // Solo cuenta si la factura más reciente de ese producto cambió el precio
  return Object.values(res).filter(c => c.cambio).sort((a, b) => b.fecha.localeCompare(a.fecha));
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
