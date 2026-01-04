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
3.  **Establecer Propiedades del Script:**
    *   En el editor de Apps Script, ve a `Configuración del proyecto` (el icono del engranaje ⚙️).
    *   En la sección `Propiedades del secuencia de comandos`, haz clic en `Añadir propiedad de la secuencia de comandos`.
    *   Añade la siguiente propiedad **obligatoria**:
        *   **Propiedad:** `RFC_PROPIO`
        *   **Valor:** `TU_RFC_AQUI` (reemplaza con el RFC de tu empresa)
    *   Guarda los cambios.

## 2. Estructura de Hojas de Cálculo

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

## 3. Hojas de Reportes (Fórmulas)

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

## 4. Instrucciones de Uso

1.  **Poblar Catálogos:** Llena las hojas `CatalogoCuentas` y `CatalogoProveedores` con tu información.
2.  **Recargar la Hoja:** Guarda los cambios y recarga la hoja de cálculo para que aparezca el menú personalizado.
3.  **Cargar XML:**
    *   Ve al menú `Contabilidad Automatizada` > `Cargar XML desde PC`.
    *   La primera vez, se te pedirá que autorices los permisos para que el script se ejecute.
    *   Se abrirá un cuadro de diálogo. Haz clic en el botón para seleccionar archivos y elige los XML desde tu computadora.
    *   Haz clic en "Cargar y Procesar".
    *   El sistema leerá los archivos, los procesará, validará duplicados y generará las pólizas automáticamente.
    *   Revisa las hojas de registro y el `LibroDiario` para ver los resultados.
