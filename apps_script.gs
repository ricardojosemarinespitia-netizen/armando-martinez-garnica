/**
 * CONTACTOS WEB — Armando Martínez Garnica
 * Recibe las solicitudes del formulario de la página y las guarda en la hoja,
 * con opción de avisarle a Armando por correo en cada solicitud nueva.
 *
 * INSTALACIÓN (5 minutos):
 * 1. Crea una hoja de cálculo en sheets.google.com, llámala "Contactos Armando".
 * 2. Extensiones → Apps Script. Borra lo que haya y pega este código completo.
 * 3. (Opcional) En NOTIFICAR_A pon el correo de Armando para que le llegue aviso.
 * 4. Implementar → Nueva implementación → tipo "Aplicación web":
 *      - Ejecutar como: Tú
 *      - Acceso: Cualquier persona
 *    → Implementar → copia la URL (termina en /exec).
 * 5. Pega esa URL en index.html donde dice PEGA_AQUI_LA_URL_DE_TU_APPS_SCRIPT.
 */

var NOTIFICAR_A = "";  // ej: "armando@ejemplo.com" — vacío = sin correo de aviso

function doPost(e) {
  try {
    var d = JSON.parse(e.postData.contents);
    var nombre = String(d.nombre || "").slice(0, 120).trim();
    var numero = String(d.numero || "").slice(0, 30).trim();
    var motivo = String(d.motivo || "").slice(0, 600).trim();
    if (!nombre || !numero || !motivo) {
      return salida({ ok: false, error: "campos incompletos" });
    }

    var hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Contactos")
            || SpreadsheetApp.getActiveSpreadsheet().insertSheet("Contactos");
    if (hoja.getLastRow() === 0) {
      hoja.appendRow(["Fecha", "Nombre", "Número", "Motivo", "Atendido"]);
      hoja.getRange("A1:E1").setFontWeight("bold").setBackground("#38404D").setFontColor("#F6F2E8");
      hoja.setColumnWidths(1, 5, 160);
      hoja.setColumnWidth(4, 380);
    }
    hoja.appendRow([new Date(), nombre, numero, motivo, "NO"]);

    if (NOTIFICAR_A) {
      MailApp.sendEmail({
        to: NOTIFICAR_A,
        subject: "Nueva solicitud de contacto: " + nombre,
        body: "Nombre: " + nombre + "\nNúmero: " + numero + "\nMotivo:\n" + motivo +
              "\n\nRegistrado en la hoja 'Contactos Armando'."
      });
    }
    return salida({ ok: true });
  } catch (err) {
    return salida({ ok: false, error: String(err) });
  }
}

function salida(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
