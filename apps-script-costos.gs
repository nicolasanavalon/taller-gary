/*
 * Backend de sincronización para "Costos y Cotizaciones" (Taller Gary).
 * Mismo patrón que usa el módulo de Producción: un JSON con número de revisión,
 * guardado en una celda de esta misma Google Sheet, con bloqueo para que dos
 * guardados al mismo tiempo no se pisen.
 *
 * CÓMO INSTALARLO (una sola vez):
 * 1) Ve a https://sheets.google.com y crea una hoja de cálculo nueva en blanco.
 *    Ponle de nombre, por ejemplo: "Taller Gary - Costos (nube)".
 * 2) Arriba, en el menú, ve a: Extensiones > Apps Script.
 * 3) Se abre un editor con un archivo "Código.gs" con un código de ejemplo.
 *    BORRA todo ese contenido y PEGA completo este archivo en su lugar.
 * 4) Arriba a la derecha, haz clic en "Implementar" > "Nueva implementación".
 * 5) Junto a "Selecciona el tipo", haz clic en el ícono de engranaje ⚙️ y elige "Aplicación web".
 * 6) Configura:
 *      - Descripción: "Costos Gary v1" (o lo que quieras)
 *      - Ejecutar como: Yo (tu cuenta)
 *      - Quién tiene acceso: Cualquier usuario
 * 7) Haz clic en "Implementar". Te va a pedir autorizar permisos (es tu propio script,
 *    dale "Avanzado" > "Ir a [nombre del proyecto] (no seguro)" si Google muestra esa advertencia
 *    — es normal para scripts propios que no están publicados en la tienda de Google).
 * 8) Al terminar, te da un link que termina en "/exec". Copia ese link completo y pásamelo.
 */

const CELDA_DATA = 'A1';
const CELDA_REV = 'B1';

function hoja_() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
}

function doGet(e) {
  const sh = hoja_();
  const raw = sh.getRange(CELDA_DATA).getValue();
  const rev = Number(sh.getRange(CELDA_REV).getValue() || 0);
  let payload = {};
  if (raw) {
    try { payload = JSON.parse(raw); } catch (err) { payload = {}; }
  }
  payload._rev = rev;
  return ContentService.createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = hoja_();
    const body = JSON.parse(e.postData.contents);
    const currentRev = Number(sh.getRange(CELDA_REV).getValue() || 0);
    const ifRev = Number(body._ifRev || 0);

    if (currentRev !== ifRev) {
      // Alguien más guardó algo más nuevo mientras tanto: no lo pisamos.
      const raw = sh.getRange(CELDA_DATA).getValue();
      let current = {};
      if (raw) { try { current = JSON.parse(raw); } catch (err) { current = {}; } }
      current._rev = currentRev;
      return ContentService.createTextOutput(JSON.stringify({ conflict: true, current }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    delete body._ifRev;
    const newRev = currentRev + 1;
    sh.getRange(CELDA_DATA).setValue(JSON.stringify(body));
    sh.getRange(CELDA_REV).setValue(newRev);
    return ContentService.createTextOutput(JSON.stringify({ rev: newRev }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}
