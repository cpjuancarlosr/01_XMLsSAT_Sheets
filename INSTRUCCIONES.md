# Sistema Contable-Fiscal en Google Sheets (Avanzado)

Este documento proporciona las instrucciones para configurar y utilizar la herramienta de contabilidad avanzada.

## 1. Configuración Inicial de Apps Script

1.  **Abrir el Editor de Scripts:** En tu hoja de cálculo de Google, ve a `Extensiones` > `Apps Script`.
2.  **Copiar los Archivos de Código:**
    *   Crea/actualiza los siguientes archivos en el editor con el código proporcionado:
        *   `UI.gs`, `Server.gs`, `Contabilidad.gs`, `SheetManager.gs`, `XMLParser.gs`, `Picker.html`, `Tests.gs`, y el nuevo `Styling.gs`.
3.  **Establecer Propiedades del Script:**
    *   Ve a `Configuración del proyecto` ⚙️ y en `Propiedades del secuencia de comandos`, añade:
        *   **Propiedad:** `RFC_PROPIO`
        *   **Valor:** `TU_RFC_AQUI` (reemplaza con el RFC de tu empresa)

## 2. Primer Uso y Configuración del Espacio de Trabajo

1.  **Recargar la Hoja de Cálculo:** Guarda todos los archivos en el editor de Apps Script y **recarga (F5)** tu hoja de Google Sheets.
2.  **Autorizar el Script:** Al recargar, aparecerá un nuevo menú llamado `Contabilidad Automatizada`. Haz clic en él y en `Cargar XML desde PC`. Se te pedirá que autorices los permisos del script. Sigue los pasos para permitirlo.
3.  **Creación Automática de Hojas:** El script creará automáticamente todas las hojas necesarias (`CatalogoCuentas`, `CuentasPorCobrarPagar`, etc.) la primera vez que se abra la hoja de cálculo. **Ya no necesitas crearlas manualmente.**
4.  **Poblar Catálogos:** Ve a las hojas recién creadas y llena la información mínima requerida:
    *   **`CatalogoCuentas`**: Asegúrate de que las cuentas listadas en la sección de Fórmulas existan.
    *   **`CatalogoProveedores`** y **`CatalogoProdServ`**: Llena con los RFCs de tus proveedores y las claves de producto que usas con más frecuencia.

## 3. Aplicar el Estilo Visual (Opcional)

*   Para aplicar el tema visual profesional "Industrial Moderno" (negro, blanco, verde neón), ve al nuevo menú `Estilo` > `Aplicar Tema Industrial Moderno`.

## 4. Hojas de Reportes (Fórmulas)

Crea las siguientes hojas y pega las fórmulas en la celda indicada.
*(Las fórmulas no han cambiado, puedes usar las de la versión anterior)*.

### `Estado de Resultados`, `Balance General`, `Flujo de Efectivo (Método Indirecto)`
*(Pega aquí las mismas fórmulas de la versión anterior del documento).*

## 5. Instrucciones de Uso

1.  Ve al menú `Contabilidad Automatizada` > `Cargar XML desde PC`.
2.  Selecciona **todos** tus archivos XML (Ingresos, Egresos y Pagos). El sistema los procesará en el orden correcto.
3.  Revisa las hojas `LibroDiario` y `CuentasPorCobrarPagar` para ver los resultados.

---
**Nota sobre Errores Comunes:**
Si ejecutas la función `onOpen` manualmente desde el editor de Apps Script, verás un error `TypeError: SpreadsheetApp.getUi is not a function`. **Esto es normal y no es un error real del código.** Las funciones de UI solo pueden ser llamadas cuando el usuario interactúa con la hoja de cálculo, no desde el editor. Simplemente ignora este error.
