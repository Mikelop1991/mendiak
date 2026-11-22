/**
 * MENDIAK · back de la app (Google Apps Script)
 * Pega este código en Extensiones > Apps Script de la hoja "Mendiak_datos_app".
 * La hoja debe tener las pestañas "Salidas" y "Listas" tal como vienen en el archivo.
 *
 * Primera vez: ejecuta la función `configurar` (botón ▶ arriba) y copia la clave que sale en el registro.
 *
 * EXCEL BONITO EN DRIVE: sube "MENDIAK_plantilla.xlsx" a tu Drive (sin convertir) y ejecuta una vez
 * `activarExcelAutomatico`. Crea "Mendiak Carlos.xlsx" junto a la plantilla y lo rehace cada hora si hubo cambios.
 */

const HOJA_SALIDAS = 'Salidas';
const HOJA_LISTAS = 'Listas';
const CAMPOS = ['id', 'fecha', 'fechaFin', 'monte', 'sierra', 'provincia',
  'cota1', 'cota2', 'cota3', 'cota4', 'cota5', 'km', 'salida', 'participantes',
  'modalidad', 'idRuta', 'notas', 'creado', 'modificado'];

/** Genera la clave secreta de la app (solo hace falta una vez). */
function configurar() {
  const props = PropertiesService.getScriptProperties();
  let clave = props.getProperty('CLAVE');
  if (!clave) {
    clave = Utilities.getUuid().replace(/-/g, '').slice(0, 20);
    props.setProperty('CLAVE', clave);
  }
  Logger.log('Clave de la app: ' + clave);
  return clave;
}

function doGet(e) {
  return responder_(() => {
    comprobarClave_(e.parameter.k);
    if (e.parameter.accion === 'excel') return Object.assign({ ok: true }, generarExcel());
    const v = version_();
    if (e.parameter.v && e.parameter.v === v) return { ok: true, sinCambios: true, version: v, hora: new Date().toISOString() };
    return Object.assign({ ok: true, version: v, hora: new Date().toISOString() }, datos_());
  });
}

/** Lo que la app necesita para pintar todo. */
function datos_() { return { excel: excelInfo_(), salidas: leerSalidas_(), listas: leerListas_(), coords: leerCoords_() }; }

/* ---- versión de los datos: cambia con cada escritura, desde la app o a mano en la hoja.
   La app la manda al pedir datos y, si no ha cambiado nada, no hace falta leer la hoja entera. ---- */
function version_() { return PropertiesService.getScriptProperties().getProperty('VERSION') || '0'; }
function subirVersion_() { PropertiesService.getScriptProperties().setProperty('VERSION', String(Date.now())); }
/** Se ejecuta solo al editar la hoja a mano: así la app se entera del cambio. */
function onEdit() { try { subirVersion_(); } catch (err) { /* sin permiso en este contexto: la app refresca igual cada 10 min */ } }

function doPost(e) {
  return responder_(() => {
    const p = JSON.parse(e.postData.contents);
    comprobarClave_(p.k);
    const lock = LockService.getScriptLock();
    if (!lock.tryLock(20000)) throw new Error('OCUPADO');
    try {
      // cada cambio por separado: si uno falla, los demás se guardan igual y la app sabe cuál no entró
      const hechos = (p.ops || [p]).map(op => {
        try { return Object.assign({ ok: true }, aplicar_(op)); }
        catch (err) { return { ok: false, accion: op && op.accion, error: String(err.message || err) }; }
      });
      SpreadsheetApp.flush();
      if (hechos.some(h => h.ok)) { PropertiesService.getScriptProperties().setProperty('CAMBIOS', '1'); subirVersion_(); }
      // con datos: true devuelve ya los datos al día, y la app se ahorra pedirlos otra vez
      return Object.assign({ ok: true, hechos: hechos }, p.datos ? Object.assign({ version: version_(), hora: new Date().toISOString() }, datos_()) : {});
    } finally {
      lock.releaseLock();
    }
  });
}

function aplicar_(op) {
  const hoja = SpreadsheetApp.getActive().getSheetByName(HOJA_SALIDAS);
  const ahora = new Date();
  if (op.accion === 'crear') {
    const s = limpiar_(op.salida);
    // si ya existe (reintento sin conexión), no se duplica
    if (s.id && buscarFila_(hoja, s.id) > 0) return { accion: 'crear', id: s.id, repetida: true };
    s.id = s.id || nuevoId_(hoja);
    s.creado = ahora; s.modificado = ahora;
    hoja.appendRow(aFila_(s));
    return { accion: 'crear', id: s.id };
  }
  if (op.accion === 'editar') {
    const fila = buscarFila_(hoja, op.salida.id);
    if (fila < 0) throw new Error('No encuentro la salida ' + op.salida.id);
    const actual = hoja.getRange(fila, 1, 1, CAMPOS.length).getValues()[0];
    const s = limpiar_(op.salida);
    s.creado = actual[CAMPOS.indexOf('creado')] || '';
    s.modificado = ahora;
    hoja.getRange(fila, 1, 1, CAMPOS.length).setValues([aFila_(s)]);
    return { accion: 'editar', id: s.id };
  }
  if (op.accion === 'borrar') {
    const fila = buscarFila_(hoja, op.id);
    if (fila > 0) {
      // copia de seguridad antes de borrar
      const papelera = SpreadsheetApp.getActive().getSheetByName('Papelera') ||
        SpreadsheetApp.getActive().insertSheet('Papelera');
      papelera.appendRow(hoja.getRange(fila, 1, 1, CAMPOS.length).getValues()[0].concat([ahora]));
      hoja.deleteRow(fila);
    }
    return { accion: 'borrar', id: op.id };
  }
  throw new Error('Acción desconocida');
}

/** Revisa lo que llega de la app antes de escribirlo en la hoja. */
function limpiar_(s) {
  if (!s) throw new Error('Salida vacía');
  const fecha = /^\d{4}-\d{2}-\d{2}$/;
  if (!fecha.test(String(s.fecha || ''))) throw new Error('Fecha no válida');
  const monte = String(s.monte || '').trim();
  if (!monte) throw new Error('Falta el monte');
  const o = {};
  CAMPOS.forEach(c => {
    const x = s[c];
    o[c] = (x === undefined || x === null) ? '' : (typeof x === 'string' ? x.trim().slice(0, 1000) : x);
  });
  o.monte = monte.slice(0, 200);
  if (o.fechaFin && !fecha.test(String(o.fechaFin))) o.fechaFin = '';
  return o;
}

function leerSalidas_() {
  const hoja = SpreadsheetApp.getActive().getSheetByName(HOJA_SALIDAS);
  const n = hoja.getLastRow() - 1;
  if (n < 1) return [];
  const tz = Session.getScriptTimeZone();
  const vals = hoja.getRange(2, 1, n, 17).getValues();
  return vals.filter(v => v[0] !== '' || v[3] !== '').map(v => v.map((x, i) => {
    if (x instanceof Date) return Utilities.formatDate(x, tz, 'yyyy-MM-dd');
    return x;
  }));
}

function leerListas_() {
  const hoja = SpreadsheetApp.getActive().getSheetByName(HOJA_LISTAS);
  const v = hoja.getDataRange().getValues().slice(1);
  const col = i => v.map(r => r[i]).filter(x => x !== '' && x !== null);
  const pares = (a, b) => v.filter(r => r[a] !== '' && r[b] !== '').map(r => [String(r[a]), String(r[b])]);
  return {
    propietario: col(0)[0] || 'Carlos',
    companeros: col(1),
    aliasComp: pares(3, 4),
    aliasMonte: pares(6, 7),
    aliasProv: pares(9, 10),
    provincias: col(12)
  };
}

function aFila_(s) {
  return CAMPOS.map(c => {
    let x = s[c];
    if (x === undefined || x === null) return '';
    if ((c === 'fecha' || c === 'fechaFin') && x) {
      const p = String(x).split('-');
      return new Date(+p[0], +p[1] - 1, +p[2]);
    }
    return x;
  });
}

function buscarFila_(hoja, id) {
  const n = hoja.getLastRow() - 1;
  if (n < 1) return -1;
  const ids = hoja.getRange(2, 1, n, 1).getValues();
  for (let i = 0; i < ids.length; i++) if (String(ids[i][0]) === String(id)) return i + 2;
  return -1;
}

function nuevoId_(hoja) {
  return 'S' + Utilities.formatDate(new Date(), 'UTC', 'yyMMddHHmmss') + Math.floor(Math.random() * 90 + 10);
}

function comprobarClave_(k) {
  const clave = PropertiesService.getScriptProperties().getProperty('CLAVE');
  if (!clave || k !== clave) throw new Error('CLAVE');
}

function responder_(fn) {
  let out;
  try { out = fn(); } catch (err) { out = { ok: false, error: String(err.message || err) }; }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

/* ====================== Excel bonito en Drive ====================== */
const PLANTILLA = 'MENDIAK_plantilla.xlsx';
const NOMBRE_EXCEL = 'Mendiak Carlos.xlsx';
const MIME_XLSX = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

/** Ejecútala una vez: genera el Excel, lo rehace cada hora si hubo cambios y guarda una copia de la hoja cada semana. */
function activarExcelAutomatico() {
  ScriptApp.getProjectTriggers().filter(t => ['excelSiHayCambios', 'copiaSemanal'].includes(t.getHandlerFunction())).forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('excelSiHayCambios').timeBased().everyHours(1).create();
  ScriptApp.newTrigger('copiaSemanal').timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(4).create();
  copiaSemanal();
  const r = generarExcel();
  Logger.log('Excel listo: ' + r.url);
}
/** Copia de seguridad de la hoja de datos en la carpeta "Mendiak copias" (guarda las 8 últimas). */
function copiaSemanal() {
  const ss = SpreadsheetApp.getActive();
  const madre = DriveApp.getFileById(ss.getId()).getParents();
  const base = madre.hasNext() ? madre.next() : DriveApp.getRootFolder();
  const it = base.getFoldersByName('Mendiak copias');
  const carpeta = it.hasNext() ? it.next() : base.createFolder('Mendiak copias');
  DriveApp.getFileById(ss.getId()).makeCopy('Mendiak datos · copia ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'), carpeta);
  const copias = []; const f = carpeta.getFiles(); while (f.hasNext()) copias.push(f.next());
  copias.sort((x, y) => y.getDateCreated() - x.getDateCreated()).slice(8).forEach(x => x.setTrashed(true));
}
function excelSiHayCambios() {
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty('CAMBIOS') === '1') {
    generarExcel(); props.setProperty('CAMBIOS', '0');
    try { buscarCoordenadas(); } catch (err) { Logger.log('Mapa: ' + err.message); }   // montes nuevos para el mapa
  }
}
function excelInfo_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('EXCEL_ID');
  return id ? { url: 'https://drive.google.com/file/d/' + id + '/view', fecha: props.getProperty('EXCEL_FECHA') || '' } : null;
}

function generarExcel() {
  const props = PropertiesService.getScriptProperties();
  let plantilla;
  const pid = props.getProperty('PLANTILLA_ID');
  try { if (pid) plantilla = DriveApp.getFileById(pid); } catch (e) { plantilla = null; }
  if (!plantilla) {
    const it = DriveApp.getFilesByName(PLANTILLA);
    if (!it.hasNext()) throw new Error('No encuentro ' + PLANTILLA + ' en tu Drive. Súbela sin convertir.');
    plantilla = it.next(); props.setProperty('PLANTILLA_ID', plantilla.getId());
  }
  const blobs = Utilities.unzip(plantilla.getBlob().setContentType('application/zip'));
  const files = {};
  blobs.forEach(b => { if (/\.xml$|\.rels$/.test(b.getName())) files[b.getName()] = b.getDataAsString('UTF-8'); });
  const cambiados = rellenarPlantilla_(files, leerSalidas_(), leerListas_());
  const nuevos = blobs.map(b => cambiados[b.getName()] !== undefined
    ? Utilities.newBlob(cambiados[b.getName()], 'application/xml', b.getName())
    : b);
  const zip = Utilities.zip(nuevos, NOMBRE_EXCEL).setContentType(MIME_XLSX);

  let id = props.getProperty('EXCEL_ID');
  let ok = false;
  if (id) {
    try {
      DriveApp.getFileById(id).setTrashed(false);
      const r = UrlFetchApp.fetch('https://www.googleapis.com/upload/drive/v3/files/' + id + '?uploadType=media', {
        method: 'patch', contentType: MIME_XLSX, payload: zip.getBytes(),
        headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true
      });
      ok = r.getResponseCode() < 300;
    } catch (e) { ok = false; }
  }
  if (!ok) {
    const parents = plantilla.getParents();
    const carpeta = parents.hasNext() ? parents.next() : DriveApp.getRootFolder();
    const f = carpeta.createFile(zip);
    id = f.getId(); props.setProperty('EXCEL_ID', id);
  }
  const fecha = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
  props.setProperty('EXCEL_FECHA', fecha);
  props.setProperty('CAMBIOS', '0');
  return { url: 'https://drive.google.com/file/d/' + id + '/view', fecha: fecha };
}

/* --- relleno de la plantilla (texto XML puro, sin dependencias) --- */
function rellenarPlantilla_(files, filas, listas) {
  const wb = files['xl/workbook.xml'], rels = files['xl/_rels/workbook.xml.rels'];
  const hoja = nombre => {
    const m = wb.match(new RegExp('<sheet[^>]*name="' + nombre + '"[^>]*r:id="([^"]+)"'));
    const t = rels.match(new RegExp('<Relationship[^>]*Id="' + m[1] + '"[^>]*Target="([^"]+)"')) ||
              rels.match(new RegExp('<Relationship[^>]*Target="([^"]+)"[^>]*Id="' + m[1] + '"'));
    const target = t[1].replace(/^\/?xl\//, '');
    return 'xl/' + target;
  };
  const out = {};
  const pDatos = hoja('Datos'), pListas = hoja('Listas');

  // salidas en orden alfabético por monte, como en el Excel original
  const salidas = filas.filter(r => r[3] !== '' && r[3] != null)
    .sort((a, b) => String(a[3]).localeCompare(String(b[3]), 'es') || String(a[1]).localeCompare(String(b[1])));
  const COLS = 'ABCDEFGHIJKLMNOP'.split('');
  // fila de la app: [id, fecha, fechaFin, monte, sierra, provincia, c1..c5, km, salida, participantes, modalidad, idRuta, notas]
  const valores = r => [fechaSerie_(r[1]), fechaSerie_(r[2]), r[3], r[4], r[5], r[6], r[7], r[8], r[9], r[10], r[11], r[12], r[13], r[14], r[15], r[16]];
  out[pDatos] = editarFilas_(files[pDatos], (n, fila) => {
    if (n < 2) return fila;
    const v = n - 2 < salidas.length ? valores(salidas[n - 2]) : null;
    COLS.forEach((c, i) => { fila = ponerCelda_(fila, c, n, v ? v[i] : '', c === 'A' || c === 'B'); });
    return fila;
  });
  // compañeros y erratas de compañeros
  const comp = listas.companeros || [], ali = listas.aliasComp || [];
  out[pListas] = editarFilas_(files[pListas], (n, fila) => {
    if (n < 5 || n > 120) return fila;
    fila = ponerCelda_(fila, 'V', n, comp[n - 5] || '');
    if (n <= 34) { const a = ali[n - 5] || ['', '']; fila = ponerCelda_(fila, 'AA', n, a[0]); fila = ponerCelda_(fila, 'AB', n, a[1]); }
    return fila;
  });
  // que Excel lo recalcule todo al abrir
  out['xl/workbook.xml'] = /fullCalcOnLoad/.test(wb) ? wb : wb.replace(/<calcPr/, '<calcPr fullCalcOnLoad="1"');
  return out;
}
function editarFilas_(xml, fn) {
  return xml.replace(/<row r="(\d+)"[^>]*?(?:\/>|>[\s\S]*?<\/row>)/g, (fila, n) => {
    if (/\/>$/.test(fila) && !/<\/row>$/.test(fila)) fila = fila.replace(/\/>$/, '></row>');
    return fn(+n, fila);
  });
}
function colNum_(c) { let n = 0; for (const ch of c) n = n * 26 + ch.charCodeAt(0) - 64; return n; }
function xmlEsc_(s) { return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, ''); }
function ponerCelda_(fila, col, n, valor, esFecha) {
  const ref = col + n;
  const re = new RegExp('<c r="' + ref + '"[^>]*?(?:\\/>|>[\\s\\S]*?<\\/c>)');
  const prev = fila.match(re);
  const sm = prev && prev[0].match(/ s="(\d+)"/);
  const s = sm ? ' s="' + sm[1] + '"' : '';
  let celda;
  if (valor === '' || valor === null || valor === undefined) celda = '<c r="' + ref + '"' + s + '/>';
  else if (typeof valor === 'number' || (esFecha && !isNaN(valor))) celda = '<c r="' + ref + '"' + s + '><v>' + valor + '</v></c>';
  else if (/^-?\d+(\.\d+)?$/.test(String(valor).trim()) && String(valor).trim().length < 15) celda = '<c r="' + ref + '"' + s + '><v>' + String(valor).trim() + '</v></c>';
  else celda = '<c r="' + ref + '"' + s + ' t="inlineStr"><is><t xml:space="preserve">' + xmlEsc_(valor) + '</t></is></c>';
  if (prev) return fila.replace(re, celda);
  if (!s && celda.endsWith('/>')) return fila; // nada que poner
  // insertar en su sitio por orden de columna
  const k = colNum_(col);
  const cells = /<c r="([A-Z]+)\d+"/g; let m;
  while ((m = cells.exec(fila))) { if (colNum_(m[1]) > k) return fila.slice(0, m.index) + celda + fila.slice(m.index); }
  return fila.replace(/<\/row>$/, celda + '</row>');
}
function fechaSerie_(v) {
  if (v === '' || v == null) return '';
  let y, m, d;
  if (v instanceof Date) { y = v.getFullYear(); m = v.getMonth() + 1; d = v.getDate(); }
  else { const p = String(v).slice(0, 10).split('-'); if (p.length < 3) return ''; y = +p[0]; m = +p[1]; d = +p[2]; }
  return Math.round((Date.UTC(y, m - 1, d) - Date.UTC(1899, 11, 30)) / 86400000);
}

/* ====================== Mapa: coordenadas de cada monte ======================
   La app pinta en un mapa los montes subidos. Las coordenadas se buscan en OpenStreetMap
   cruzando el nombre del monte con la cota anotada, y se guardan en la hoja «Coordenadas»,
   donde se pueden revisar y corregir a mano:
     · ESTADO «ok»: nombre y altitud cuadran.            · «revisar»: conviene mirarlo (enlace VER).
     · «no encontrado»: no sale en el mapa.              · «manual»: puesto a mano; el script no lo toca.
   Para corregir uno: pon LATITUD y LONGITUD buenas y escribe «manual» en ESTADO.
   La primera vez, ejecuta buscarCoordenadas() a mano (tarda unos minutos).
   Después, los montes nuevos se buscan solos cada hora junto con el Excel. */
const HOJA_COORDS = 'Coordenadas';
const OVERPASS = 'https://overpass-api.de/api/interpreter';
// zonas donde buscar (sur, oeste, norte, este): Navarra y Euskal Herria, Pirineo, norte, Asturias y Picos, centro
const ZONAS_MAPA = [[42.3, -3.6, 43.5, -0.7], [42.0, -0.7, 43.2, 3.4], [41.6, -4.5, 43.4, -1.6], [42.6, -7.2, 43.6, -4.3], [40.1, -5.6, 41.3, -3.2]];

/** Busca las coordenadas de los montes que aún no están en la hoja «Coordenadas». */
function buscarCoordenadas() {
  const hoja = hojaCoords_();
  const hechos = new Set(hoja.getLastRow() > 1 ? hoja.getRange(2, 1, hoja.getLastRow() - 1, 1).getValues().map(r => normal_(r[0])) : []);
  const faltan = cumbresConCota_().filter(c => !hechos.has(normal_(c.nombre)));
  if (!faltan.length) { Logger.log('Todos los montes tienen ya su fila en «Coordenadas».'); return 0; }
  Logger.log('Montes por buscar: ' + faltan.length);
  const picos = [];
  ZONAS_MAPA.forEach(z => { try { picos.push.apply(picos, picosEn_(z)); } catch (err) { Logger.log('Zona ' + z + ': ' + err.message); } });
  Logger.log('Cumbres de OpenStreetMap descargadas: ' + picos.length);
  const indice = indicePicos_(picos);
  const t0 = Date.now(); let fuera = 0;
  const filas = faltan.map(c => {
    let r = elegir_(c, indice.get(clave_(c.nombre)) || []);
    // si no está en las zonas de siempre (Alpes, Marruecos…), una búsqueda más amplia, con límite de tiempo
    if (!r && fuera < 15 && Date.now() - t0 < 180000) { fuera++; try { r = elegir_(c, buscarFuera_(c.nombre)); } catch (err) { /* sin conexión con OSM: queda «no encontrado» */ } }
    return r ? [c.nombre, r.lat, r.lon, r.ele || '', r.nombre, r.estado, c.cota || '']
             : [c.nombre, '', '', '', '', 'no encontrado', c.cota || ''];
  });
  const ini = hoja.getLastRow() + 1;
  hoja.getRange(ini, 1, filas.length, 7).setValues(filas);
  hoja.getRange(ini, 8, filas.length, 1).setFormulas(filas.map((f, i) => [`=IF(B${ini + i}="","",HYPERLINK("https://www.openstreetmap.org/?mlat="&B${ini + i}&"&mlon="&C${ini + i}&"#map=15/"&B${ini + i}&"/"&C${ini + i},"VER"))`]));
  subirVersion_();
  const n = s => filas.filter(f => f[5] === s).length;
  Logger.log(`Listo: ${n('ok')} bien, ${n('revisar')} para revisar, ${n('no encontrado')} sin encontrar.`);
  return filas.length;
}
/** Igual que buscarCoordenadas(), pero vuelve a probar los «no encontrado» (por si se ha corregido el nombre). */
function reintentarNoEncontrados() {
  const hoja = hojaCoords_(); const n = hoja.getLastRow() - 1; if (n < 1) return;
  const v = hoja.getRange(2, 6, n, 1).getValues();
  for (let i = n - 1; i >= 0; i--) if (v[i][0] === 'no encontrado') hoja.deleteRow(i + 2);
  buscarCoordenadas();
}
function hojaCoords_() {
  const ss = SpreadsheetApp.getActive();
  let h = ss.getSheetByName(HOJA_COORDS);
  if (!h) {
    h = ss.insertSheet(HOJA_COORDS);
    h.getRange(1, 1, 1, 8).setValues([['MONTE', 'LATITUD', 'LONGITUD', 'ALTITUD EN EL MAPA', 'NOMBRE EN EL MAPA', 'ESTADO', 'COTA ANOTADA', 'VER']]).setFontWeight('bold');
    h.setFrozenRows(1);
  }
  return h;
}
/** Lo que la app necesita: [monte, lat, lon] de los que tienen coordenadas. */
function leerCoords_() {
  const h = SpreadsheetApp.getActive().getSheetByName(HOJA_COORDS);
  if (!h || h.getLastRow() < 2) return [];
  return h.getRange(2, 1, h.getLastRow() - 1, 6).getValues()
    .filter(r => r[0] !== '' && r[1] !== '' && r[2] !== '' && r[5] !== 'no encontrado')
    .map(r => [String(r[0]), Number(r[1]), Number(r[2])]);
}
/** Cada monte principal (como lo agrupa la app) con su cota más repetida. */
function cumbresConCota_() {
  const alias = new Map(leerListas_().aliasMonte.map(p => [normal_(p[0]), String(p[1]).toUpperCase()]));
  const seps = [',', ' Y ', ' E ', ' O ', ' CON ', ' EN BICI', '(', '-'];
  const m = new Map();
  leerSalidas_().forEach(f => {
    const c = String(f[3] || '').toUpperCase().trim(); if (!c) return;
    let cut = c.length; seps.forEach(s => { const i = c.indexOf(s); if (i > 0 && i < cut) cut = i; });
    const raw = c.slice(0, cut).trim(); const nombre = alias.get(normal_(raw)) || raw;
    const e = m.get(nombre) || { nombre: nombre, cotas: {} };
    const cota = primeraCota_(f[6]); if (cota) e.cotas[cota] = (e.cotas[cota] || 0) + 1;
    m.set(nombre, e);
  });
  return [...m.values()].map(e => {
    const l = Object.keys(e.cotas).sort((a, b) => e.cotas[b] - e.cotas[a]);
    return { nombre: e.nombre, cota: l.length ? +l[0] : null };
  });
}
function primeraCota_(v) {
  if (v === '' || v == null) return null;
  if (typeof v === 'number') { const n = Math.trunc(v); return n >= 100 && n < 9000 ? n : null; }
  const n = parseInt(String(v).split(/[^\d.]+/)[0].replace(/\./g, ''), 10);
  return n >= 100 && n < 9000 ? n : null;
}
function normal_(s) { return String(s || '').normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase().trim(); }
/** Clave para comparar nombres escritos de formas distintas: Aitzkorri/Aizkorri, Ezkaba/Ezcaba, Pico de Anie/Anie… */
function clave_(s) {
  let t = normal_(s).replace(/[’'`]/g, ' ').replace(/[^a-z ]/g, ' ').replace(/\s+/g, ' ').trim();
  const gen = /^(pico|pic|picu|punta|pena|monte|mont|mendi|alto|cerro|tuca|tuc|tossal|tozal|cima|cumbre|puy|soum|tuca|la|el|los|las|de|del|d|l|le|du|des)\s+/;
  for (let i = 0; i < 4 && gen.test(t); i++) t = t.replace(gen, '');
  return t.replace(/\s+/g, '')
    .replace(/tx/g, 'ch').replace(/tz/g, 'z').replace(/ts/g, 's').replace(/qu/g, 'c').replace(/k/g, 'c')
    .replace(/gu(?=[ei])/g, 'g').replace(/mb/g, 'nb').replace(/v/g, 'b').replace(/y/g, 'i').replace(/h/g, '').replace(/z/g, 's')
    .replace(/(.)\1+/g, '$1');
}
/** Descarga las cumbres con nombre de una zona: [{lat, lon, ele, nombres}] */
function picosEn_(z) {
  const q = `[out:csv(::lat,::lon,name,"name:eu","name:es","name:fr",alt_name,ele;false;"\\t")][timeout:120];` +
            `node["natural"~"^(peak|volcano|hill)$"]["name"](${z.join(',')});out;`;
  return leerCsv_(pedirOverpass_(q));
}
function buscarFuera_(nombre) {
  const txt = normal_(nombre).replace(/[^a-z ]/g, '').trim(); if (txt.length < 4) return [];
  const q = `[out:csv(::lat,::lon,name,"name:eu","name:es","name:fr",alt_name,ele;false;"\\t")][timeout:60];` +
            `node["natural"~"^(peak|volcano)$"]["name"~"${txt}",i](27,-18,60,20);out;`;
  return leerCsv_(pedirOverpass_(q));
}
function pedirOverpass_(q) {
  for (let intento = 0; intento < 3; intento++) {
    const r = UrlFetchApp.fetch(OVERPASS, { method: 'post', payload: { data: q }, muteHttpExceptions: true });
    if (r.getResponseCode() === 200) return r.getContentText('UTF-8');
    Utilities.sleep(15000 * (intento + 1));   // OSM ocupado: esperar y repetir
  }
  throw new Error('OpenStreetMap no responde');
}
function leerCsv_(txt) {
  return txt.split('\n').filter(Boolean).map(l => {
    const c = l.split('\t');
    const nombres = [];
    c.slice(2, 7).forEach(n => String(n || '').split(/[\/;,]|\(|\)/).forEach(x => { if (x.trim()) nombres.push(x.trim()); }));
    return { lat: +c[0], lon: +c[1], ele: parseInt(String(c[7] || '').replace(',', '.'), 10) || null, nombre: c[2] || nombres[0] || '', nombres: nombres };
  }).filter(p => isFinite(p.lat) && isFinite(p.lon) && p.nombres.length);
}
function indicePicos_(picos) {
  const m = new Map();
  picos.forEach(p => new Set(p.nombres.map(clave_)).forEach(k => { if (!k) return; const l = m.get(k) || []; l.push(p); m.set(k, l); }));
  return m;
}
/** De los picos con ese nombre, el de altitud más parecida a la anotada. */
function elegir_(c, cands) {
  const k = clave_(c.nombre);
  const l = cands.filter(p => p.nombres.some(n => clave_(n) === k));
  if (!l.length) return null;
  const dif = p => (c.cota && p.ele) ? Math.abs(p.ele - c.cota) : null;
  l.sort((a, b) => (dif(a) ?? 9999) - (dif(b) ?? 9999));
  const p = l[0], d = dif(p);
  // bien si la altitud cuadra (±60 m), o si no hay cota pero solo hay un pico con ese nombre
  const ok = d !== null ? d <= 60 : l.length === 1;
  return { lat: Math.round(p.lat * 1e5) / 1e5, lon: Math.round(p.lon * 1e5) / 1e5, ele: p.ele, nombre: p.nombre, estado: ok ? 'ok' : 'revisar' };
}
