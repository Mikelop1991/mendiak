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
    return { ok: true, excel: excelInfo_(), salidas: leerSalidas_(), listas: leerListas_(), hora: new Date().toISOString() };
  });
}

function doPost(e) {
  return responder_(() => {
    const p = JSON.parse(e.postData.contents);
    comprobarClave_(p.k);
    const lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try {
      const hechos = [];
      (p.ops || [p]).forEach(op => hechos.push(aplicar_(op)));
      PropertiesService.getScriptProperties().setProperty('CAMBIOS', '1');
      return { ok: true, hechos: hechos };
    } finally {
      lock.releaseLock();
    }
  });
}

function aplicar_(op) {
  const hoja = SpreadsheetApp.getActive().getSheetByName(HOJA_SALIDAS);
  const ahora = new Date();
  if (op.accion === 'crear') {
    const s = op.salida;
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
    const s = op.salida;
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

/** Ejecútala una vez: genera el Excel y programa que se rehaga cada hora si hubo cambios. */
function activarExcelAutomatico() {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'excelSiHayCambios').forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('excelSiHayCambios').timeBased().everyHours(1).create();
  const r = generarExcel();
  Logger.log('Excel listo: ' + r.url);
}
function excelSiHayCambios() {
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty('CAMBIOS') === '1') { generarExcel(); props.setProperty('CAMBIOS', '0'); }
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
