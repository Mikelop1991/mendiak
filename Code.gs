/**
 * MENDIAK · back de la app (Google Apps Script)
 * Pega este código en Extensiones > Apps Script de la hoja "Mendiak_datos_app".
 * La hoja debe tener las pestañas "Salidas" y "Listas" tal como vienen en el archivo.
 *
 * Primera vez: ejecuta la función `configurar` (botón ▶ arriba) y copia la clave que sale en el registro.
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
    return { ok: true, salidas: leerSalidas_(), listas: leerListas_(), hora: new Date().toISOString() };
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
