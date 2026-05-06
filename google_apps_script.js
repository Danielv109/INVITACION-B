// ══════════════════════════════════════════════════════════
// GOOGLE APPS SCRIPT - Confirmaciones de Boda
// Copia este código completo en tu Google Apps Script
// (Extensiones → Apps Script dentro de tu Google Sheet)
// ══════════════════════════════════════════════════════════

function doPost(e) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var data = JSON.parse(e.postData.contents);

    // Buscar si ya existe una confirmación con ese nombre
    var nombres = sheet.getRange('A:A').getValues();
    var filaExistente = -1;
    for (var i = 1; i < nombres.length; i++) {
      if (nombres[i][0] === data.nombre) {
        filaExistente = i + 1;
        break;
      }
    }

    if (filaExistente > 0) {
      // Actualizar fila existente
      sheet.getRange(filaExistente, 1, 1, 5).setValues([[
        data.nombre,
        data.pases,
        data.asistencia,
        data.comentario || '',
        data.fecha
      ]]);
    } else {
      // Agregar nueva fila
      sheet.appendRow([
        data.nombre,
        data.pases,
        data.asistencia,
        data.comentario || '',
        data.fecha
      ]);
    }

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok', message: 'Confirmación guardada' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Esto permite verificar que el script funciona
function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok', message: 'API de confirmaciones activa' }))
    .setMimeType(ContentService.MimeType.JSON);
}
