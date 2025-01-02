/**
 * Configuración Inicial
 * Define los IDs de los archivos y carpetas que se usarán en el script.
 * Asegúrate de reemplazar los valores con los IDs reales de tus archivos.
 */
const CONFIG = {
  rvToolsFileId: '1z4KNXBbVGZZzaZYMoFLLY8dJfbT5mcaQ4YVAYHivmq4', // Reemplázalo con el ID del archivo RVTools.
  inventoryFileId: '1lnpM3mz5QzZMFrxj9QaZSf1KyHCQfG7sBkg4y9z-lF8',   // Reemplázalo con el ID del archivo de inventario.
};

/**
 * Función Principal
 * Coordina el proceso de sincronización entre RVTools y el inventario.
 */
function main() {
  try {
    // Abrir los archivos desde Google Drive
    const rvToolsFile = SpreadsheetApp.openById(CONFIG.rvToolsFileId);
    const inventoryFile = SpreadsheetApp.openById(CONFIG.inventoryFileId);

    // Obtener las hojas relevantes
    const vNetworkSheet = rvToolsFile.getSheetByName('vNetwork');
    const vInfoSheet = rvToolsFile.getSheetByName('vInfo');
    const unixSheet = inventoryFile.getSheetByName('Servidores UNIX');
    const intelSheet = inventoryFile.getSheetByName('MV INTEL');

    // Validar que las hojas existen
    if (!vNetworkSheet || !vInfoSheet || !unixSheet || !intelSheet) {
      throw new Error('Una o más hojas requeridas no se encuentran en los archivos especificados.');
    }

    // Sincronizar datos
    synchronizeRvToolsWithInventory(vNetworkSheet, vInfoSheet, unixSheet, intelSheet);
    Logger.log('Sincronización completada exitosamente.');
  } catch (error) {
    Logger.log(`Error en la sincronización: ${error.message}`);
  }
}

/**
 * Sincronización entre RVTools e Inventario
 * @param {Sheet} vNetworkSheet - Hoja vNetwork de RVTools
 * @param {Sheet} vInfoSheet - Hoja vInfo de RVTools
 * @param {Sheet} unixSheet - Hoja Servidores UNIX del inventario
 * @param {Sheet} intelSheet - Hoja MV INTEL del inventario
 */
function synchronizeRvToolsWithInventory(vNetworkSheet, vInfoSheet, unixSheet, intelSheet) {
  // Obtener encabezados y datos de RVTools
  const vNetworkHeaders = vNetworkSheet.getRange(1, 1, 1, vNetworkSheet.getLastColumn()).getValues()[0];
  const vNetworkData = vNetworkSheet.getRange(2, 1, vNetworkSheet.getLastRow() - 1, vNetworkSheet.getLastColumn()).getValues();

  const vInfoHeaders = vInfoSheet.getRange(1, 1, 1, vInfoSheet.getLastColumn()).getValues()[0];
  const vInfoData = vInfoSheet.getRange(2, 1, vInfoSheet.getLastRow() - 1, vInfoSheet.getLastColumn()).getValues();

  // Obtener encabezados del inventario
  const unixHeaders = unixSheet.getRange(3, 1, 1, unixSheet.getLastColumn()).getValues()[0];
  const intelHeaders = intelSheet.getRange(3, 1, 1, intelSheet.getLastColumn()).getValues()[0];

  // Procesar cada registro de vNetwork
  vNetworkData.forEach(row => {
    const hostname = row[vNetworkHeaders.indexOf('VM')];
    const ipAddress = row[vNetworkHeaders.indexOf('IPv4 Address')];
    // Verifica y actualiza el inventario UNIX
    if (hostname && ipAddress) {
      updateOrCreateRow(unixSheet, unixHeaders, { HOSTNAME: hostname, 'IP PROD': ipAddress });
    }
  });

  Logger.log('Datos sincronizados en la hoja Servidores UNIX.');
}

/**
 * Actualiza o Crea una Fila en el Inventario
 * @param {Sheet} sheet - Hoja de cálculo donde actualizar o crear filas
 * @param {Array} headers - Encabezados de la hoja
 * @param {Object} data - Datos a insertar o actualizar
 */
function updateOrCreateRow(sheet, headers, data) {
  const rows = sheet.getDataRange().getValues();
  let found = false;

  // Busca si el registro ya existe
  rows.forEach((row, index) => {
    if (row[headers.indexOf('HOSTNAME')] === data.HOSTNAME) {
      found = true;
      headers.forEach((header, colIndex) => {
        if (data[header]) {
          sheet.getRange(index + 1, colIndex + 1).setValue(data[header]);
        }
      });
    }
  });

  // Si no existe, crea una nueva fila
  if (!found) {
    const newRow = headers.map(header => data[header] || '');
    sheet.appendRow(newRow);
  }
}
