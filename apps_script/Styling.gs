// Definición de la paleta de colores para el tema "Industrial Moderno"
const THEME_COLORS = {
  BACKGROUND: '#1c1c1c', // Negro/Gris muy oscuro
  TEXT: '#ffffff',       // Blanco
  HEADER_BACKGROUND: '#000000', // Negro puro para encabezados
  NEON_GREEN: '#39ff14', // Verde neón brillante
  GRID_LINES: '#444444', // Líneas de cuadrícula grises oscuras
};

/**
 * Aplica un tema estético profesional "Industrial Moderno" a todas las hojas gestionadas por el sistema.
 */
function applyModernIndustrialTheme() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const allSheets = ss.getSheets();

  // Lista de las hojas que queremos estilizar (basado en SHEETS de SheetManager.gs)
  const managedSheetNames = Object.values(SHEETS);

  allSheets.forEach(sheet => {
    // Solo aplicar estilo a las hojas que son parte de nuestro sistema
    if (managedSheetNames.includes(sheet.getName())) {
      const fullRange = sheet.getDataRange();
      const headerRange = sheet.getRange(1, 1, 1, sheet.getLastColumn());

      // 1. Estilo general de la hoja
      fullRange
        .setBackground(THEME_COLORS.BACKGROUND)
        .setFontColor(THEME_COLORS.TEXT)
        .setFontFamily('Consolas'); // Fuente monoespaciada para un look técnico

      // 2. Estilo de los encabezados
      headerRange
        .setBackground(THEME_COLORS.HEADER_BACKGROUND)
        .setFontColor(THEME_COLORS.NEON_GREEN)
        .setFontWeight('bold');

      // 3. Quitar las líneas de cuadrícula por defecto para un look más limpio
      sheet.setGridLinesVisibility(false);

      // 4. (Opcional) Añadir bordes personalizados que actúen como nuevas "líneas de cuadrícula"
      fullRange.setBorder(
        true, true, true, true, true, true,
        THEME_COLORS.GRID_LINES,
        SpreadsheetApp.BorderStyle.SOLID
      );

      // 5. Ajustar el ancho de las columnas al contenido
      sheet.autoResizeColumns(1, sheet.getLastColumn());
    }
  });

  SpreadsheetApp.getUi().alert('Tema Industrial Moderno aplicado correctamente.');
}
