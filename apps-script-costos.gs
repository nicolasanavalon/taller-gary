/*
 * Backend de Taller Gary: hace DOS cosas en el mismo script.
 * 1) Sincronización en la nube de "Costos y Cotizaciones" (como ya la tenías).
 * 2) Puente ("proxy") hacia la API de Compra Ágil de Mercado Público, porque esa
 *    API no permite ser llamada directo desde una página web (no tiene CORS
 *    habilitado) — así que el navegador le pide el favor a este script, y este
 *    script (que no tiene esa restricción) hace la consulta real y devuelve el
 *    resultado tal cual.
 *
 * CÓMO ACTUALIZAR TU SCRIPT EXISTENTE (mismo link, no se pierde nada):
 * 1) Abre tu Google Sheet de "Taller Gary - Costos (nube)".
 * 2) Extensiones > Apps Script.
 * 3) Selecciona TODO el código actual y bórralo. Pega completo este archivo en su lugar.
 * 4) Arriba a la derecha: Implementar > Administrar implementaciones.
 * 5) Haz clic en el ícono de lápiz (✏️) de la implementación activa.
 * 6) En "Versión", elige "Nueva versión" (NO crees una implementación nueva —
 *    así el link que ya usa la web sigue siendo el mismo, no hay que cambiar nada más).
 * 7) Haz clic en "Implementar". Listo — mismo link de siempre, con el código nuevo.
 */

const CELDA_DATA = 'A1';
const CELDA_REV = 'B1';

function hoja_() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
}

function doGet(e) {
  const accion = (e.parameter && e.parameter.action) || '';
  if (accion === 'compra_agil_lista') return proxyCompraAgilLista_(e);
  if (accion === 'compra_agil_detalle') return proxyCompraAgilDetalle_(e);

  // Comportamiento original: leer los datos de Costos y Cotizaciones guardados.
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

/* ---------- Puente hacia la API de Compra Ágil (api2.mercadopublico.cl) ---------- */
const COMPRA_AGIL_BASE = 'https://api2.mercadopublico.cl/v2/compra-agil';
const PARAMS_LISTA_PERMITIDOS = ['q','region','estado','publicado_desde','publicado_hasta',
  'cambio_desde','cambio_hasta','ttl_cambio_ms','tamano_pagina','numero_pagina','ordenar_por'];

function proxyCompraAgilLista_(e) {
  const ticket = e.parameter.ticket || '';
  const partes = [];
  PARAMS_LISTA_PERMITIDOS.forEach(function(k){
    if (e.parameter[k] !== undefined && e.parameter[k] !== '') {
      partes.push(k + '=' + encodeURIComponent(e.parameter[k]));
    }
  });
  const url = COMPRA_AGIL_BASE + (partes.length ? '?' + partes.join('&') : '');
  return llamarCompraAgil_(url, ticket);
}

function proxyCompraAgilDetalle_(e) {
  const ticket = e.parameter.ticket || '';
  const codigo = e.parameter.codigo || '';
  const url = COMPRA_AGIL_BASE + '/' + encodeURIComponent(codigo);
  return llamarCompraAgil_(url, ticket);
}

function llamarCompraAgil_(url, ticket) {
  try {
    const resp = UrlFetchApp.fetch(url, {
      headers: { ticket: ticket },
      muteHttpExceptions: true,
    });
    return ContentService.createTextOutput(resp.getContentText())
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({
      success: 'NOK', payload: null,
      errors: [{ codigo: '500', mensaje: 'Error del proxy: ' + err.message, detalle: null }],
    })).setMimeType(ContentService.MimeType.JSON);
  }
}
