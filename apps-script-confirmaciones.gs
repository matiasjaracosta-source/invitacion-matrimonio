/**
 * Confirmaciones de asistencia · Matrimonio Mati & Javi
 * Recibe las respuestas del formulario de la invitación, las anota en la planilla
 * y mantiene al día la hoja "Lista definitiva" (formulario + confirmación manual).
 */

const HOJAS_INVITADOS = ['Mati', 'Javi'];
const HOJA_RESPUESTAS = 'Respuestas';
const HOJA_DEFINITIVA = 'Lista definitiva';
const COLUMNAS = ['Confirmación formulario', 'Detalle confirmación', 'Confirmación manual', 'Observaciones', 'Fecha de confirmación'];
const COLUMNAS_QUITAR = ['Restricciones alimentarias', 'Mensaje para los novios', 'Mensaje'];
const VERDE = '#22482A';

function doPost(e) {
  // ejecutado desde el editor (sin datos): prepara la planilla
  if (!e || !e.postData) { prepararColumnas(); return json({ ok: true, preparado: true }); }
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const d = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const personas = (d.personas || []).filter(p => p && p.nombre);
    const ahora = new Date();

    // 1) registro de cada respuesta
    let log = ss.getSheetByName(HOJA_RESPUESTAS);
    if (!log) {
      log = ss.insertSheet(HOJA_RESPUESTAS);
      log.appendRow(['Fecha', 'Hoja', 'Saludo', 'Nombre', 'Asiste']);
      log.getRange(1, 1, 1, 5).setFontWeight('bold').setBackground(VERDE).setFontColor('#FFFFFF');
      log.setFrozenRows(1);
    }
    personas.forEach(p => log.appendRow([ahora, d.hoja || '', d.saludo || '', p.nombre, p.asiste ? 'Sí' : 'No']));

    // 2) marcar la fila de la invitación
    const hit = buscarFila(ss, d.hoja, d.saludo);
    if (hit) {
      const sh = hit.sheet, r = hit.row, col = asegurarColumnas(sh);
      const si = personas.filter(p => p.asiste).length, total = personas.length;
      sh.getRange(r, col['Confirmación formulario']).setValue(si ? '✓ ' + si + ' de ' + total : '✗ No asistirán')
        .setBackground(si ? '#E3F1DA' : '#F6E0DC');
      sh.getRange(r, col['Detalle confirmación']).setValue(personas.map(p => (p.asiste ? '✓ ' : '✗ ') + p.nombre).join('\n'));
      sh.getRange(r, col['Fecha de confirmación']).setValue(ahora);
    }
    actualizarListaDefinitiva();
    return json({ ok: true, encontrado: !!hit });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function doGet(e) {
  const p = (e && e.parameter) || {};
  if (!p.saludo) return json({ ok: true });
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hit = buscarFila(ss, p.t, p.saludo);
  if (!hit) return json({ ok: true, encontrado: false, confirmado: false });
  const col = asegurarColumnas(hit.sheet);
  const det = String(hit.sheet.getRange(hit.row, col['Detalle confirmación']).getValue() || '').trim();
  const personas = det ? det.split(/\n|;/).map(leerLinea).filter(Boolean) : [];
  return json({ ok: true, encontrado: true, confirmado: personas.length > 0, personas: personas });
}

/** Correcciones a mano: casilla manual, detalle (✓/✗, o borrarlo), nombres o saludo → se actualiza todo. */
function onEdit(e) {
  try {
    const sh = e.range.getSheet();
    if (HOJAS_INVITADOS.indexOf(sh.getName()) < 0) return;
    const c1 = e.range.getColumn(), nc = e.range.getNumColumns(), r1 = e.range.getRow(), nr = e.range.getNumRows();
    const enc = sh.getRange(1, c1, 1, nc).getValues()[0].map(String);
    let refrescar = false;
    enc.forEach(h => {
      if (h === 'Detalle confirmación' || h === 'Confirmación formulario') { for (let r = r1; r < r1 + nr; r++) recalcularResumen(sh, r); refrescar = true; }
      if (h === 'Confirmación manual' || h.indexOf('Nombres') === 0 || h.indexOf('Saludo') === 0) refrescar = true;
    });
    if (refrescar) actualizarListaDefinitiva();
  } catch (err) {}
}

/** Lee una línea del detalle: ✓/sí → asiste, ✗/x/no → no asiste. */
function leerLinea(l) {
  l = String(l || '').trim();
  const m = l.match(/^(?:✓|✔|\+)\s*(.+)$/) || l.match(/^(?:si|sí|ok)(?:\s*[:\-]\s*|\s+)(.+)$/i);
  if (m) return { asiste: true, nombre: m[1].trim() };
  const n = l.match(/^(?:✗|✘|-)\s*(.+)$/) || l.match(/^(?:x|no)(?:\s*[:\-]\s*|\s+)(.+)$/i);
  if (n) return { asiste: false, nombre: n[1].trim() };
  return null;
}

/** Recalcula "✓ X de Y" cuando se corrige o se borra el detalle a mano. */
function recalcularResumen(sh, r) {
  if (r < 2) return;
  const col = asegurarColumnas(sh);
  const txt = String(sh.getRange(r, col['Detalle confirmación']).getValue() || '').trim();
  const celda = sh.getRange(r, col['Confirmación formulario']);
  if (!txt) { celda.clearContent().setBackground(null); sh.getRange(r, col['Fecha de confirmación']).clearContent(); return; }
  const ps = txt.split(/\n|;/).map(leerLinea).filter(Boolean);
  // deja el detalle con el formato estándar
  sh.getRange(r, col['Detalle confirmación']).setValue(ps.map(p => (p.asiste ? '✓ ' : '✗ ') + p.nombre).join('\n'));
  const si = ps.filter(p => p.asiste).length;
  celda.setValue(si ? '✓ ' + si + ' de ' + ps.length : '✗ No asistirán').setBackground(si ? '#E3F1DA' : '#F6E0DC');
}

/** Reconstruye la hoja "Lista definitiva" con todas las personas confirmadas. */
function actualizarListaDefinitiva() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const filas = [];
  HOJAS_INVITADOS.forEach(nombreHoja => {
    const sh = ss.getSheetByName(nombreHoja);
    if (!sh || sh.getLastRow() < 2) return;
    const col = asegurarColumnas(sh);
    const datos = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
    datos.forEach(f => {
      const saludo = String(f[1] || '').trim();
      if (!saludo) return;
      const detalle = String(f[col['Detalle confirmación'] - 1] || '').trim();
      const manual = f[col['Confirmación manual'] - 1] === true;
      const fecha = f[col['Fecha de confirmación'] - 1] || '';
      const obs = f[col['Observaciones'] - 1] || '';
      if (detalle) {
        detalle.split(/\n|;/).map(leerLinea).filter(p => p && p.asiste)
          .forEach(p => filas.push([p.nombre, saludo, nombreHoja, 'Formulario', fecha, obs]));
      } else if (manual) {
        String(f[2] || saludo).split(',').map(s => s.trim()).filter(String).forEach(n => filas.push([n, saludo, nombreHoja, 'Manual', '', obs]));
      }
    });
  });

  let out = ss.getSheetByName(HOJA_DEFINITIVA);
  if (!out) out = ss.insertSheet(HOJA_DEFINITIVA, 0);
  out.clear();
  const enc = ['N°', 'Nombre', 'Invitación', 'Lado', 'Confirmó por', 'Fecha de confirmación', 'Observaciones'];
  out.getRange(1, 1, 1, enc.length).setValues([enc]).setFontWeight('bold').setBackground(VERDE).setFontColor('#FFFFFF');
  if (filas.length) out.getRange(2, 1, filas.length, enc.length).setValues(filas.map((f, i) => [i + 1].concat(f)));
  out.getRange(1, 9).setValue('Total confirmados').setFontWeight('bold');
  out.getRange(1, 10).setValue(filas.length).setFontWeight('bold').setFontColor(VERDE);
  out.getRange(2, 9).setValue('Lado Mati');
  out.getRange(2, 10).setValue(filas.filter(f => f[2] === 'Mati').length);
  out.getRange(3, 9).setValue('Lado Javi');
  out.getRange(3, 10).setValue(filas.filter(f => f[2] === 'Javi').length);
  out.setFrozenRows(1);
  out.setColumnWidth(2, 220); out.setColumnWidth(3, 200); out.setColumnWidth(7, 220);
  if (filas.length) out.getRange(2, 6, filas.length, 1).setNumberFormat('dd/mm/yyyy hh:mm');
}

function buscarFila(ss, hoja, saludo) {
  const norm = s => String(s || '').trim().toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '');
  const objetivo = norm(saludo);
  if (!objetivo) return null;
  const orden = hoja && HOJAS_INVITADOS.indexOf(hoja) >= 0 ? [hoja].concat(HOJAS_INVITADOS.filter(h => h !== hoja)) : HOJAS_INVITADOS;
  for (const nombre of orden) {
    const sh = ss.getSheetByName(nombre);
    if (!sh || sh.getLastRow() < 2) continue;
    const saludos = sh.getRange(2, 2, sh.getLastRow() - 1, 1).getValues();
    for (let i = 0; i < saludos.length; i++) {
      if (norm(saludos[i][0]) === objetivo) return { sheet: sh, row: i + 2 };
    }
  }
  return null;
}

function asegurarColumnas(sh) {
  const ultima = Math.max(sh.getLastColumn(), 1);
  const encabezados = sh.getRange(1, 1, 1, ultima).getValues()[0].map(String);
  const mapa = {};
  let siguiente = ultima + 1;
  COLUMNAS.forEach(nombre => {
    let i = encabezados.indexOf(nombre);
    if (i < 0) {
      sh.getRange(1, siguiente).setValue(nombre).setFontWeight('bold').setBackground(VERDE).setFontColor('#FFFFFF').setWrap(true);
      i = siguiente - 1; siguiente++;
    }
    mapa[nombre] = i + 1;
  });
  return mapa;
}

function quitarColumnas(sh) {
  if (!sh) return;
  const enc = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1)).getValues()[0].map(String);
  for (let c = enc.length; c >= 1; c--) if (COLUMNAS_QUITAR.indexOf(enc[c - 1]) >= 0) sh.deleteColumn(c);
}

/** Prepara la planilla: columnas, casillas y la lista definitiva. */
function prepararColumnas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  HOJAS_INVITADOS.concat([HOJA_RESPUESTAS]).forEach(n => quitarColumnas(ss.getSheetByName(n)));
  HOJAS_INVITADOS.forEach(nombre => {
    const sh = ss.getSheetByName(nombre);
    if (!sh) return;
    const col = asegurarColumnas(sh);
    const filas = Math.max(sh.getLastRow() - 1, 1);
    const manual = sh.getRange(2, col['Confirmación manual'], filas, 1);
    if (!manual.getDataValidation()) manual.insertCheckboxes();
    sh.getRange(2, col['Detalle confirmación'], filas, 1).setWrap(true);
    sh.setColumnWidth(col['Detalle confirmación'], 260);
    sh.setColumnWidth(col['Observaciones'], 220);
  });
  actualizarListaDefinitiva();
}

function json(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
