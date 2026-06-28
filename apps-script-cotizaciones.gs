/**
 * REGISTRO CONSECUTIVO DE COTIZACIONES → Google Sheets
 * ------------------------------------------------------
 * Este script recibe los datos de cada cotización que envías desde el
 * cotizador y los agrega como una fila nueva (con folio consecutivo)
 * en una pestaña llamada "Cotizaciones" dentro de tu hoja de cálculo.
 *
 * ====== CÓMO ACTIVARLO (una sola vez) ======
 * 1. Abre tu Google Sheet de precios.
 * 2. Menú: Extensiones → Apps Script.
 * 3. Borra lo que haya y pega TODO este archivo. Guarda (icono de disco).
 * 4. Arriba a la derecha: Implementar → Nueva implementación.
 *      - Tipo: "Aplicación web".
 *      - Ejecutar como: "Yo" (tu cuenta).
 *      - Quién tiene acceso: "Cualquier persona".
 *      - Clic en "Implementar" y autoriza los permisos.
 * 5. Copia la "URL de la aplicación web" (termina en /exec).
 * 6. Pega esa URL en index.html, en la línea:
 *        var APPS_SCRIPT_URL = '';
 *    Debe quedar así:
 *        var APPS_SCRIPT_URL = 'https://script.google.com/macros/s/XXXX/exec';
 * 7. Listo. Cada vez que presiones "💬 WhatsApp" se guardará la cotización.
 *
 * Nota: si más adelante cambias el código del script, debes volver a
 * "Implementar → Administrar implementaciones" y publicar una nueva versión.
 */

function doGet(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000); // evita folios duplicados si se envían dos a la vez
  try {
    var ENCABEZADOS = [
      'Folio', 'Fecha', 'Cliente', 'Producto', 'Medidas (cm)',
      'Material', 'Calibre Platina', 'N° Cierres', 'Cantidad',
      'Precio Unit.', 'Total c/IVA', 'Registrado'
    ];

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var hoja = ss.getSheetByName('Cotizaciones');
    if (!hoja) {
      hoja = ss.insertSheet('Cotizaciones');
    }
    // Asegura que la fila de encabezados siempre esté actualizada
    // (escribe los títulos aunque la pestaña ya existiera de antes).
    hoja.getRange(1, 1, 1, ENCABEZADOS.length).setValues([ENCABEZADOS]);
    hoja.setFrozenRows(1);

    var p = e.parameter;
    var folio = p.folio || ('000' + hoja.getLastRow()).slice(-3);

    hoja.appendRow([
      folio,
      p.fecha          || '',
      p.cliente        || '',
      p.producto       || '',
      p.medidas        || '',
      p.material       || '',
      p.calibrePlatina || '',
      p.cierres        || '',
      p.cantidad       || '',
      p.precio         || '',
      p.total          || '',
      new Date()
    ]);

    return responder(p.callback, { ok: true, folio: folio });
  } catch (err) {
    return responder(e.parameter.callback, { ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function responder(callback, obj) {
  var json = JSON.stringify(obj);
  if (callback) {
    return ContentService
      .createTextOutput(callback + '(' + json + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService
    .createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}
