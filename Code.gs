/**
 * SALAZAR 5K — Google Apps Script
 * Conecta esta aplicación web a la hoja de cálculo de inscripciones.
 * Revisa la política de datos y limita el acceso a la hoja al equipo autorizado.
 */
const SHEET_NAME = "Inscripciones";

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents || "{}");
    const required = ["fullName", "docNumber", "age", "gender", "shirtSize", "email", "healthNote", "emergencyPhone", "priorRaces"];
    for (const key of required) {
      if (data[key] === undefined || String(data[key]).trim() === "") {
        return json_({ ok: false, message: "Falta un campo obligatorio." });
      }
    }
    const age = Number(data.age);
    if (!Number.isInteger(age) || age < 18 || age > 110) {
      return json_({ ok: false, message: "El participante debe tener 18 años o más." });
    }
    if (data.adultDeclaration !== true || data.rulesAccepted !== true || data.dataConsent !== true) {
      return json_({ ok: false, message: "Faltan aceptaciones obligatorias." });
    }
    const safe = value => {
      const s = String(value == null ? "" : value).trim();
      return /^[=+\\-@]/.test(s) ? "'" + s : s;
    };
    const sheet = getSheet_();
    sheet.appendRow([
      new Date(), safe(data.fullName), safe(data.docNumber), age,
      safe(data.gender), safe(data.shirtSize), safe(data.email),
      safe(data.healthNote), safe(data.emergencyPhone), safe(data.priorRaces),
      data.rulesAccepted === true, data.dataConsent === true,
      data.imageConsent === true, data.adultDeclaration === true
    ]);
    return json_({ ok: true, message: "Inscripción registrada." });
  } catch (err) {
    console.error(err);
    return json_({ ok: false, message: "No se pudo registrar la inscripción." });
  }
}
function doGet() {
  return json_({ ok: true, message: "Servicio de inscripción Salazar 5K activo." });
}
function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow([
      "Fecha de registro", "Nombre completo", "Número de documento", "Edad",
      "Sexo", "Talla de camiseta", "Correo electrónico",
      "Condición/restricción comunicada", "Número de contacto de emergencia",
      "Participó antes en carreras", "Acepta reglamento", "Autoriza tratamiento de datos",
      "Autoriza uso de imagen", "Declara mayoría de edad"
    ]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, 14).setFontWeight("bold");
  }
  return sheet;
}
function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
