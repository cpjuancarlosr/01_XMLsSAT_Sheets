# Sistema Contable-Fiscal en Google Sheets (Avanzado)

Este documento proporciona las instrucciones para configurar y utilizar la herramienta de contabilidad avanzada.

## 1. Configuración Inicial de Apps Script

1.  **Abrir el Editor de Scripts:** En tu hoja de cálculo de Google, ve a `Extensiones` > `Apps Script`.
2.  **Copiar los Archivos de Código:**
    *   Crea/actualiza los siguientes archivos en el editor con el código proporcionado:
        *   `UI.gs`, `Server.gs`, `Contabilidad.gs`, `SheetManager.gs`, `XMLParser.gs`, `Picker.html`, `Tests.gs`, y `Styling.gs`.
3.  **Establecer Propiedades del Script:**
    *   Ve a `Configuración del proyecto` ⚙️ y en `Propiedades del secuencia de comandos`, añade:
        *   **Propiedad:** `RFC_PROPIO`
        *   **Valor:** `TU_RFC_AQUI` (reemplaza con el RFC de tu empresa)

## 2. Primer Uso y Configuración del Espacio de Trabajo

1.  **Recargar la Hoja de Cálculo:** Guarda todos los archivos en el editor de Apps Script y **recarga (F5)** tu hoja de Google Sheets. Al hacerlo, aparecerán los nuevos menús `CONTABILIDAD` y `SISTEMA`.
2.  **Preparar Espacio de Trabajo:**
    *   Ve al menú `SISTEMA` > `1. Preparar Espacio de Trabajo`.
    *   La primera vez, se te pedirá que autorices los permisos del script. Sigue los pasos para permitirlo.
    *   El script creará automáticamente todas las hojas necesarias (`CatalogoCuentas`, `CuentasPorCobrarPagar`, etc.) con sus encabezados.
3.  **Poblar Catálogos:** Ve a las hojas recién creadas y llena la información mínima requerida:
    *   **`CatalogoCuentas`**: Asegúrate de que las cuentas contables clave existan.
    *   **`CatalogoProveedores`** y **`CatalogoProdServ`**: Llena con los RFCs de tus proveedores y las claves de producto que usas con más frecuencia.

## 3. Aplicar el Estilo Visual (Opcional)

*   Para aplicar el tema visual "Industrial Moderno", ve al menú `SISTEMA` > `2. Aplicar Tema Visual`.

## 4. Hojas de Reportes (Fórmulas)

Crea las siguientes hojas y pega las fórmulas en la celda indicada.

### `Estado de Resultados`, `Balance General`, `Flujo de Efectivo (Método Indirecto)`
*(Las fórmulas de la versión anterior siguen siendo válidas y pueden ser pegadas aquí).*

## 5. Instrucciones de Uso

1.  Ve al menú `CONTABILIDAD` > `Cargar XML desde PC`.
2.  Selecciona tus archivos XML.
3.  Revisa las hojas `LibroDiario` y `CuentasPorCobrarPagar` para ver los resultados.

---
**Nota sobre Errores Comunes:**
Si ejecutas la función `onOpen` manualmente desde el editor de Apps Script, verás un error `TypeError: SpreadsheetApp.getUi is not a function`. **Esto es normal y no es un error real del código.** Las funciones de UI solo pueden ser llamadas cuando el usuario interactúa con la hoja de cálculo, no desde el editor. Simplemente ignora este error.
