# Sistema Contable-Fiscal en Google Sheets

Este documento proporciona las instrucciones para configurar y utilizar la herramienta de contabilidad automatizada en Google Sheets.

## 1. Configuración Inicial de Apps Script

1.  **Abrir el Editor de Scripts:** En tu hoja de cálculo de Google, ve a `Extensiones` > `Apps Script`.
2.  **Copiar los Archivos de Código:**
    *   Crea cada uno de los siguientes archivos `.gs` y `.html` en el editor.
    *   Copia y pega el contenido completo de cada archivo proporcionado en su correspondiente archivo en el editor de Apps Script.
        *   `UI.gs`
        *   `Server.gs`
        *   `Contabilidad.gs` (Contiene la lógica de negocio y las cuentas predeterminadas. Puedes ajustar las cuentas en el objeto `CUENTAS_PREDETERMINADAS` al inicio del archivo si es necesario).
        *   `SheetManager.gs`
        *   `XMLParser.gs`
        *   `Picker.html`
        *   (Opcional) `SAT_Downloader.gs` (si deseas usar la funcionalidad de descarga masiva)
3.  **Establecer Propiedades del Script:**
    *   En el editor de Apps Script, ve a `Configuración del proyecto` (el icono del engranaje ⚙️).
    *   En la sección `Propiedades del secuencia de comandos`, haz clic en `Añadir propiedad de la secuencia de comandos`.
    *   Añade la siguiente propiedad **obligatoria**:
        *   **Propiedad:** `RFC_PROPIO`
        *   **Valor:** `TU_RFC_AQUI` (reemplaza con el RFC de tu empresa)
    *   Guarda los cambios.

## 2. Configuración de Google Cloud y API Picker

Para que el selector de archivos de Google Drive funcione, debes vincular tu proyecto de Apps Script a un proyecto de Google Cloud y habilitar las APIs necesarias.

1.  **Crear un Proyecto en Google Cloud:** Ve a la [Consola de Google Cloud](https://console.cloud.google.com/) y crea un nuevo proyecto.
2.  **Vincular Proyecto:**
    *   En el editor de Apps Script, ve a `Configuración del proyecto` ⚙️.
    *   En la sección `Proyecto de Google Cloud`, haz clic en `Cambiar proyecto`.
    *   Pega el **Número de proyecto** de Google Cloud y haz clic en `Establecer proyecto`.
3.  **Habilitar APIs:**
    *   En tu proyecto de Google Cloud, ve a `APIs y servicios` > `Biblioteca`.
    *   Busca y habilita las siguientes APIs:
        *   **Google Drive API**
        *   **Google Picker API**
4.  **Crear Credenciales:**
    *   Ve a `APIs y servicios` > `Credenciales`.
    *   Haz clic en `+ CREAR CREDENCIALES` y selecciona `Clave de API`. Copia esta clave.
    *   Vuelve a hacer clic en `+ CREAR CREDENCIALES` y selecciona `ID de cliente de OAuth`.
        *   Tipo de aplicación: `Aplicación web`.
        *   En `URI de redireccionamiento autorizados`, añade la URL que te proporciona Apps Script. Para encontrarla, ve a `Implementar` > `Nueva implementación` en el editor de script y busca el `ID de implementación` bajo `Aplicación web`. La URL será `https://script.google.com/macros/d/{ID_IMPLEMENTACION}/usercallback`.
        *   Copia el `ID de cliente`.
5.  **Actualizar el Código HTML:**
    *   Abre el archivo `Picker.html` en el editor de Apps Script.
    *   Reemplaza `'YOUR_API_KEY'` con la **Clave de API** que creaste.
    *   Reemplaza `'YOUR_APP_ID'` con el **ID de cliente** que creaste.

## 3. Estructura de Hojas de Cálculo

Crea las siguientes hojas en tu Google Sheet con los nombres y columnas exactos.

---

### **`CatalogoCuentas`**
*   **Función:** Define tu plan contable.
*   **Columnas:**
    *   `A`: **Número de Cuenta** (e.g., 102-Bancos, 401-Ventas)
    *   `B`: **Nombre de Cuenta** (e.g., Bancos Nacionales, Ventas Nacionales Tasa 16%)

### **`CatalogoProveedores`**
*   **Función:** Asigna una cuenta contable de gasto/costo a cada proveedor.
*   **Columnas:**
    *   `A`: **RFC del Proveedor**
    *   `B`: **Cuenta Contable Asignada** (e.g., 601-Gastos Generales)

### **`CFDI Ingresos`** (Creada automáticamente)
*   **Función:** Registro de todos los XML de ingreso procesados.

### **`CFDI Egresos`** (Creada automáticamente)
*   **Función:** Registro de todos los XML de egreso procesados.

### **`LibroDiario`** (Creada automáticamente)
*   **Función:** Registro de todas las pólizas contables generadas.

---

## 4. Hojas de Reportes (Fórmulas)

Crea las siguientes hojas y pega las fórmulas en la celda indicada.

### **`Libro Mayor`**
*   **Función:** Consulta todos los movimientos de una cuenta específica.
*   **Celda `A1`:** `Cuenta:` (etiqueta)
*   **Celda `B1`:** (Deja esta celda para escribir el número de cuenta a consultar, e.g., `102-Bancos`)
*   **Celda `A3`:** `=QUERY(LibroDiario!A:G, "SELECT * WHERE C = '"&$B$1&"'", 1)`

### **`Estado de Resultados`**
*   **Función:** Muestra los ingresos, costos y gastos para determinar la utilidad o pérdida.
*   **Ejemplo de estructura:**
    *   Celda `A2`: `Ventas`
    *   Celda `B2`: `=SUMIF(LibroDiario!C, "4*", LibroDiario!F) - SUMIF(LibroDiario!C, "4*", LibroDiario!E)`
    *   Celda `A3`: `Costos`
    *   Celda `B3`: `=SUMIF(LibroDiario!C, "5*", LibroDiario!E) - SUMIF(LibroDiario!C, "5*", LibroDiario!F)`
    *   Celda `A4`: `Gastos`
    *   Celda `B4`: `=SUMIF(LibroDiario!C, "6*", LibroDiario!E) - SUMIF(LibroDiario!C, "6*", LibroDiario!F)`
    *   Celda `A6`: `Utilidad / Pérdida`
    *   Celda `B6`: `=B2 - B3 - B4`

### **`Balance General`**
*   **Función:** Presenta un resumen de los activos, pasivos y capital de la empresa.
*   **Ejemplo de estructura:**
    *   Celda `A2`: `Activo`
    *   Celda `A3`: `Activo Circulante`
    *   Celda `B4`: `Bancos`
    *   Celda `C4`: `=SUMIF(LibroDiario!C, "102*", LibroDiario!E) - SUMIF(LibroDiario!C, "102*", LibroDiario!F)`
    *   Celda `B5`: `Clientes`
    *   Celda `C5`: `=SUMIF(LibroDiario!C, "105*", LibroDiario!E) - SUMIF(LibroDiario!C, "105*", LibroDiario!F)`
    *   ... (y así sucesivamente para todas las cuentas)
    *   Repetir la estructura para `Pasivo` (cuentas 2xx) y `Capital` (cuentas 3xx).
    *   Añadir una fila para la `Utilidad del Ejercicio` que tome el valor del Estado de Resultados.
    *   Verificar la ecuación contable: **Total Activo = Total Pasivo + Capital**.

## 5. Instrucciones de Uso

1.  **Poblar Catálogos:** Llena las hojas `CatalogoCuentas` y `CatalogoProveedores` con tu información.
2.  **Recargar la Hoja:** Guarda los cambios y recarga la hoja de cálculo para que aparezca el menú personalizado.
3.  **Cargar XML:**
    *   Ve al menú `Contabilidad Automatizada` > `Cargar XML desde Drive`.
    *   La primera vez, se te pedirá que autorices los permisos para el script.
    *   Se abrirá un selector de archivos. Elige los archivos XML que deseas procesar.
    *   El sistema procesará los archivos, validará duplicados y generará las pólizas automáticamente.
    *   Revisa las hojas de registro y el `LibroDiario` para ver los resultados.
