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
  if (p.cod) return json(buscarCodigo(p.cod));   // ("c" es un parámetro reservado de Apps Script)
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
      if (h.indexOf('Nombres') === 0 || h.indexOf('Saludo') === 0 || h.indexOf('Teléfono') === 0) for (let r = Math.max(r1, 2); r < r1 + nr; r++) armarMensaje(sh, r);
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
        partirNombres(f[2] || saludo).forEach(n => filas.push([n, saludo, nombreHoja, 'Manual', '', obs]));
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

/* ================= links cortos ================= */
const URL_INVITACION = 'https://matiasjaracosta-source.github.io/invitacion-matrimonio/';
const COL_LINK = 5, COL_MENSAJE = 6, COL_ENVIAR = 7, COL_ENVIAR_PDF = 8;
const ENC_PDF = 'Enviar PDF por WhatsApp (lo completamos nosotros)';
// invitación en PDF (una hoja), compartida con "cualquier persona con el enlace"
const PDF_UNO = 'https://drive.google.com/file/d/1Lk3j-nxFZ_hT2N9iqoVT9PStZ30-pCEO/view';
const PDF_VARIOS = 'https://drive.google.com/file/d/1_29KHdELgm7e4vADRc5otVnobpmBQDNa/view';
const PLAZO = '5 de febrero de 2027';
const SUFIJO_LINK = '&calidad=baja';   // todos los invitados ven la versión liviana (la que anda fluida en cualquier celular)

/** Código de una invitación a partir de su link (…/?c=xxxxx&…). */
function codigoDe(link) {
  const m = String(link || '').match(/[?&]c=([a-z0-9]+)/i);
  return m ? m[1] : '';
}

/** "Ana Pérez, Luis Soto y Marta" → ['Ana Pérez', 'Luis Soto', 'Marta'] */
function partirNombres(txt) {
  return String(txt || '').split(/\s*,\s*|\s+y\s+|\s+&\s+/i).map(s => s.trim().replace(/\.$/, '')).filter(String);
}

/** La invitación pregunta por su código (?c=xxxxx) y recibe saludo y nombres desde la planilla. */
function buscarCodigo(c) {
  c = String(c).trim();
  if (!/^[a-z0-9]{3,10}$/i.test(c)) return { ok: false };
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  for (const nombre of HOJAS_INVITADOS) {
    const sh = ss.getSheetByName(nombre);
    if (!sh || sh.getLastRow() < 2) continue;
    const datos = sh.getRange(2, 1, sh.getLastRow() - 1, COL_LINK).getValues();
    for (const f of datos) {
      if (codigoDe(f[COL_LINK - 1]) === c && String(f[1]).trim()) {
        return { ok: true, saludo: String(f[1]).trim(), nombres: partirNombres(f[2]), hoja: nombre };
      }
    }
  }
  return { ok: false };
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Invitación')
    .addItem('Generar links de esta pestaña', 'generarLinks')
    .addToUi();
}

/** Da un link corto a cada invitación de la pestaña activa (sin tocar los que ya tienen) + mensaje y botón de WhatsApp. */
function generarLinks() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getActiveSheet();
  if (HOJAS_INVITADOS.indexOf(sh.getName()) < 0) { SpreadsheetApp.getUi().alert('Abre la pestaña Mati o Javi.'); return; }
  asegurarColumnaPdf(sh);
  const usados = {};
  HOJAS_INVITADOS.forEach(n => {
    const h = ss.getSheetByName(n);
    if (h && h.getLastRow() > 1) h.getRange(2, COL_LINK, h.getLastRow() - 1, 1).getValues()
      .forEach(f => { const c = codigoDe(f[0]); if (c) usados[c] = true; });
  });
  const letras = 'abcdefghjkmnpqrstuvwxyz23456789';
  const nuevo = () => { let c; do { c = ''; for (let i = 0; i < 5; i++) c += letras[Math.floor(Math.random() * letras.length)]; } while (usados[c]); usados[c] = true; return c; };

  const n = sh.getLastRow() - 1;
  if (n < 1) return;
  const filas = sh.getRange(2, 1, n, COL_LINK).getValues();
  const links = filas.map(f => {
    if (!String(f[1]).trim()) return [f[COL_LINK - 1]];
    // conserva el código que ya tenía (el link enviado sigue sirviendo) y le pone la calidad
    const actual = String(f[COL_LINK - 1]);
    const c = actual.indexOf(URL_INVITACION) === 0 && codigoDe(actual) ? codigoDe(actual) : nuevo();
    return [URL_INVITACION + '?c=' + c + SUFIJO_LINK];
  });
  sh.getRange(2, COL_LINK, n, 1).setValues(links);

  for (let r = 2; r <= n + 1; r++) armarMensaje(sh, r);
  sh.getRange(2, COL_MENSAJE, n, 1).setWrap(false);
  sh.setRowHeightsForced(2, n, 24);
  SpreadsheetApp.getActive().toast('Links listos. Nada se envía solo: apreta "Enviar por WhatsApp" en cada fila cuando quieras mandarla.', 'Invitación', 8);
}

/** Columna H: botón para mandar la invitación en PDF (se inserta junto al botón del link si no existe). */
function asegurarColumnaPdf(sh) {
  if (String(sh.getRange(1, COL_ENVIAR_PDF).getValue()) === ENC_PDF) return;
  sh.insertColumnAfter(COL_ENVIAR);
  sh.getRange(1, COL_ENVIAR).copyTo(sh.getRange(1, COL_ENVIAR_PDF), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
  sh.getRange(1, COL_ENVIAR_PDF).setValue(ENC_PDF);
  sh.setColumnWidth(COL_ENVIAR_PDF, sh.getColumnWidth(COL_ENVIAR));
}

/** "Ana, Luis y Marta" */
function listaNombres(ns) {
  return ns.length > 1 ? ns.slice(0, -1).join(', ') + ' y ' + ns[ns.length - 1] : ns.join('');
}

/** Botón que abre WhatsApp con el mensaje listo. No envía nada. */
function botonWhatsApp(celda, tel, msj, texto) {
  let d = String(tel).replace(/\D/g, '');
  if (d.length === 8) d = '569' + d; else if (d.length === 9) d = '56' + d;
  if (d.length < 11) { celda.setValue('Sin WhatsApp: copia el mensaje'); return; }
  celda.setRichTextValue(SpreadsheetApp.newRichTextValue().setText(texto)
    .setLinkUrl('https://wa.me/' + d + '?text=' + encodeURIComponent(msj)).build());
}

/** Mensaje con link (F), botón del link (G) y botón con la invitación en PDF (H). */
function armarMensaje(sh, r) {
  const f = sh.getRange(r, 1, 1, COL_LINK).getValues()[0];
  const saludo = String(f[1]).trim(), link = String(f[COL_LINK - 1]).trim();
  const nombres = partirNombres(f[2]), pl = nombres.length > 1;
  const celdaMsj = sh.getRange(r, COL_MENSAJE), celdaBtn = sh.getRange(r, COL_ENVIAR);
  const tienePdf = String(sh.getRange(1, COL_ENVIAR_PDF).getValue()) === ENC_PDF;
  const celdaPdf = tienePdf ? sh.getRange(r, COL_ENVIAR_PDF) : null;
  const inicio = '¡Hola ' + saludo + '! 💚\n\nCon mucha alegría queremos ' + (pl ? 'invitarlos' : 'invitarte') +
    ' a nuestro matrimonio, el sábado 6 de marzo de 2027 en Calera de Tango.\n\n';

  if (celdaPdf) {
    if (!saludo) celdaPdf.clearContent();
    else {
      const msjPdf = inicio +
        (pl ? 'Les enviamos su invitación en PDF: ' + PDF_VARIOS : 'Te enviamos tu invitación en PDF: ' + PDF_UNO) + '\n\n' +
        (pl ? 'La invitación es para: ' + listaNombres(nombres) + '.\n' : '') +
        (pl ? 'Por favor confirmen su asistencia respondiendo este mensaje hasta el ' : 'Por favor confirma tu asistencia respondiendo este mensaje hasta el ') +
        PLAZO + '.\n\nCon cariño, Mati & Javi';
      botonWhatsApp(celdaPdf, f[3], msjPdf, 'Enviar PDF por WhatsApp');
    }
  }

  if (!saludo || link.indexOf('?c=') < 0) { celdaMsj.clearContent(); celdaBtn.clearContent(); return; }
  const msj = inicio + (pl ? 'Aquí está su invitación: ' : 'Aquí está tu invitación: ') + link + '\n\nCon cariño, Mati & Javi';
  celdaMsj.setValue(msj).setWrap(false);
  botonWhatsApp(celdaBtn, f[3], msj, 'Enviar por WhatsApp');
}

function json(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
