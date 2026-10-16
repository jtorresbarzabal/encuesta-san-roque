/**
 * Encuesta de satisfacción — Restaurante Escuela San Roque
 *
 * Este archivo vive dentro de un Google Sheet (Extensiones > Apps Script)
 * y hace de "almacén de datos": guarda y lee las respuestas en la propia
 * hoja de cálculo. La página que ven los clientes ya NO vive aquí dentro
 * (vive en una web aparte, en Netlify), así que este archivo solo
 * responde datos en formato JSON cuando esa página se lo pide.
 *
 * No hace falta tocar nada de este archivo salvo que quieras cambiar
 * el nombre de la hoja donde se guardan las respuestas.
 */

var NOMBRE_HOJA_RESPUESTAS = "Respuestas";

// -----------------------------------------------------------------
// Punto de entrada para peticiones de lectura (GET)
// Ejemplos:
//   ?action=getDates
//   ?action=getResponses&date=2026-09-08
//   ?action=getServiceDate
// -----------------------------------------------------------------
function doGet(e) {
  var action = e.parameter.action;
  var resultado;

  if (action === "getDates") {
    resultado = getResponseDates();
  } else if (action === "getResponses") {
    resultado = getResponsesForDate(e.parameter.date);
  } else if (action === "getServiceDate") {
    resultado = getServiceDate();
  } else {
    resultado = { error: "Acción no reconocida" };
  }

  return jsonOutput_(resultado);
}

// -----------------------------------------------------------------
// Punto de entrada para peticiones de escritura (POST)
// El cuerpo debe ser un JSON con un campo "action":
//   { "action": "saveResponse", ...datos de la encuesta }
//   { "action": "setServiceDate", "date": "2026-09-08" }
// -----------------------------------------------------------------
function doPost(e) {
  var body;
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    return jsonOutput_({ error: "JSON inválido" });
  }

  var resultado;
  if (body.action === "saveResponse") {
    resultado = saveResponse(body);
  } else if (body.action === "setServiceDate") {
    resultado = setServiceDate(body.date);
  } else {
    resultado = { error: "Acción no reconocida" };
  }

  return jsonOutput_(resultado);
}

function jsonOutput_(objeto) {
  return ContentService
    .createTextOutput(JSON.stringify(objeto))
    .setMimeType(ContentService.MimeType.JSON);
}

// -----------------------------------------------------------------
// Convierte lo que haya en la celda de fecha a texto "AAAA-MM-DD",
// tanto si Google Sheets lo guardó como texto como si lo convirtió
// automáticamente en una fecha de verdad.
// -----------------------------------------------------------------
function normalizarFecha_(valor) {
  if (valor instanceof Date) {
    return Utilities.formatDate(valor, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }
  return String(valor || "").trim();
}

// -----------------------------------------------------------------
// Convierte cualquier valor de celda a texto seguro, aunque Google
// Sheets lo haya guardado como número (por ejemplo, un comentario
// que era solo cifras).
// -----------------------------------------------------------------
function normalizarTexto_(valor) {
  if (valor === undefined || valor === null) return "";
  return String(valor);
}

// -----------------------------------------------------------------
// Hoja de respuestas: la crea si no existe todavía
// -----------------------------------------------------------------
function getHojaRespuestas_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(NOMBRE_HOJA_RESPUESTAS);
  if (!sheet) {
    sheet = ss.insertSheet(NOMBRE_HOJA_RESPUESTAS);
    sheet.appendRow([
      "timestamp", "fecha",
      "visual", "gusto", "temperatura", "originalidad", "postres", "panes", "platos_global",
      "amabilidad", "imagen", "tecnica", "espera", "respuesta", "confort", "sala_global",
      "precio", "acompanado", "observaciones"
    ]);
    sheet.setFrozenRows(1);
  }
  // La columna "fecha" (B) y la de "observaciones" (S) se guardan
  // siempre como texto plano, para que Google Sheets no las convierta
  // solo en una fecha o un número con otro formato.
  sheet.getRange("B:B").setNumberFormat("@");
  sheet.getRange("S:S").setNumberFormat("@");
  return sheet;
}

// -----------------------------------------------------------------
// Guarda una respuesta nueva de la encuesta
// -----------------------------------------------------------------
function saveResponse(record) {
  var sheet = getHojaRespuestas_();
  var p = (record && record.platos) || {};
  var s = (record && record.sala) || {};

  sheet.appendRow([
    Date.now(),
    "'" + record.dateISO,
    p.visual || "", p.gusto || "", p.temperatura || "", p.originalidad || "", p.postres || "", p.panes || "", p.global || "",
    s.amabilidad || "", s.imagen || "", s.tecnica || "", s.espera || "", s.respuesta || "", s.confort || "", s.global || "",
    (record.precio === undefined || record.precio === null) ? "" : record.precio,
    record.acompanado || "",
    "'" + (record.observaciones || "")
  ]);

  return { ok: true };
}

// -----------------------------------------------------------------
// Lista de fechas que tienen encuestas
// -----------------------------------------------------------------
function getResponseDates() {
  var sheet = getHojaRespuestas_();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  var fechas = sheet.getRange(2, 2, lastRow - 1, 1).getValues();
  var vistas = {};
  fechas.forEach(function (fila) {
    var f = normalizarFecha_(fila[0]);
    if (f) vistas[f] = true;
  });

  var lista = Object.keys(vistas);
  lista.sort();
  lista.reverse();
  return lista;
}

// -----------------------------------------------------------------
// Todas las respuestas de un día concreto
// -----------------------------------------------------------------
function getResponsesForDate(dateISO) {
  var sheet = getHojaRespuestas_();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  var datos = sheet.getRange(2, 1, lastRow - 1, 19).getValues();
  var resultado = [];

  datos.forEach(function (fila) {
    if (normalizarFecha_(fila[1]) !== dateISO) return;
    resultado.push({
      timestamp: fila[0],
      dateISO: normalizarFecha_(fila[1]),
      platos: {
        visual: fila[2], gusto: fila[3], temperatura: fila[4],
        originalidad: fila[5], postres: fila[6], panes: fila[7], global: fila[8]
      },
      sala: {
        amabilidad: fila[9], imagen: fila[10], tecnica: fila[11],
        espera: fila[12], respuesta: fila[13], confort: fila[14], global: fila[15]
      },
      precio: fila[16],
      acompanado: normalizarTexto_(fila[17]),
      observaciones: normalizarTexto_(fila[18])
    });
  });

  return resultado;
}

// -----------------------------------------------------------------
// Fecha de servicio en curso (se guarda aparte, no en la hoja)
// -----------------------------------------------------------------
function getServiceDate() {
  return PropertiesService.getScriptProperties().getProperty("fecha_servicio") || "";
}

function setServiceDate(dateISO) {
  PropertiesService.getScriptProperties().setProperty("fecha_servicio", dateISO);
  return { ok: true };
}
