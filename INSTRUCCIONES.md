# Sistema Contable-Fiscal en Google Sheets (Avanzado)

Este documento proporciona las instrucciones para configurar y utilizar la herramienta de contabilidad avanzada, capaz de manejar PPD y Complementos de Pago.

## 1. Configuración Inicial de Apps Script

1.  **Abrir el Editor de Scripts:** Ve a `Extensiones` > `Apps Script`.
2.  **Copiar los Archivos de Código:**
    *   Crea/actualiza los siguientes archivos en el editor con el código proporcionado.
        *   `UI.gs`
        *   `Server.gs`
        *   `Contabilidad.gs`
        *   `SheetManager.gs`
        *   `XMLParser.gs`
        *   `Picker.html`
        *   `Tests.gs`
3.  **Establecer Propiedades del Script:**
    *   Ve a `Configuración del proyecto` ⚙️.
    *   En `Propiedades del secuencia de comandos`, asegúrate de tener la siguiente propiedad:
        *   **Propiedad:** `RFC_PROPIO`
        *   **Valor:** `TU_RFC_AQUI` (reemplaza con el RFC de tu empresa)
    *   Guarda los cambios.

## 2. Estructura de Hojas de Cálculo y Catálogos

Crea/actualiza las siguientes hojas con los nombres y columnas exactos.

---

### **`CatalogoCuentas`**
*   **Función:** Define tu plan contable jerárquico.
*   **Columnas:**
    *   `A`: **Número de Cuenta**
    *   `B`: **Nombre de Cuenta**
*   **Cuentas Mínimas Requeridas:**
    *   `102-Bancos`
    *   `105-Clientes`
    *   `118-IVA Acreditable`
    *   `119-IVA por Acreditar` **(Nueva - para PPD)**
    *   `201-Proveedores`
    *   `208-IVA Trasladado`
    *   `209-IVA por Trasladar` **(Nueva - para PPD)**
    *   `401-Ventas`
    *   `601-Gastos Generales`

### **`CatalogoProveedores`**
*   **Función:** Asigna una cuenta de gasto por defecto a cada proveedor.
*   **Columnas:**
    *   `A`: **RFC del Proveedor**
    *   `B`: **Cuenta Contable de Gasto Asignada**

### **`CatalogoProdServ` (Nueva)**
*   **Función:** Asigna cuentas de ingreso/gasto específicas a nivel de producto/servicio del CFDI. **Tiene prioridad sobre el catálogo de proveedores.**
*   **Columnas:**
    *   `A`: **ClaveProdServ SAT** (Ej. `84111506` para servicios de contabilidad)
    *   `B`: **Cuenta Contable Asignada** (Ej. `401-Ventas Servicios`)

---

### **Hojas Generadas Automáticamente**
*   `CFDI Ingresos`
*   `CFDI Egresos`
*   `CFDI Pagos` **(Nueva)**
*   `LibroDiario`
*   `CuentasPorCobrarPagar` **(Nueva)**

---

## 3. Hojas de Reportes (Fórmulas)

### **`Estado de Resultados`**
*   **Celda `A2`**: `Ventas`
*   **Celda `B2`**: `=SUMIF(LibroDiario!C, "4*", LibroDiario!F) - SUMIF(LibroDiario!C, "4*", LibroDiario!E)`
*   **Celda `A3`**: `Gastos`
*   **Celda `B3`**: `=SUMIF(LibroDiario!C, "6*", LibroDiario!E) - SUMIF(LibroDiario!C, "6*", LibroDiario!F)`
*   **Celda `A5`**: **`Utilidad Neta`**
*   **Celda `B5`**: `=B2 - B3`

### **`Balance General`**
*   **Activo**
    *   `Bancos`: `=SUMIF(LibroDiario!C, "102*", LibroDiario!E) - SUMIF(LibroDiario!C, "102*", LibroDiario!F)`
    *   `Clientes`: `=SUMIF(LibroDiario!C, "105*", LibroDiario!E) - SUMIF(LibroDiario!C, "105*", LibroDiario!F)`
    *   `IVA Acreditable`: `=SUMIF(LibroDiario!C, "118*", LibroDiario!E) - SUMIF(LibroDiario!C, "118*", LibroDiario!F)`
    *   `IVA por Acreditar`: `=SUMIF(LibroDiario!C, "119*", LibroDiario!E) - SUMIF(LibroDiario!C, "119*", LibroDiario!F)`
*   **Pasivo**
    *   `Proveedores`: `=SUMIF(LibroDiario!C, "201*", LibroDiario!F) - SUMIF(LibroDiario!C, "201*", LibroDiario!E)`
    *   `IVA Trasladado`: `=SUMIF(LibroDiario!C, "208*", LibroDiario!F) - SUMIF(LibroDiario!C, "208*", LibroDiario!E)`
    *   `IVA por Trasladar`: `=SUMIF(LibroDiario!C, "209*", LibroDiario!F) - SUMIF(LibroDiario!C, "209*", LibroDiario!E)`
*   **Capital**
    *   `Resultados del Ejercicio`: (Vincular a la Utilidad Neta del Estado de Resultados)
*   **Verificación:** `=SUM(Activos) = SUM(Pasivos) + Capital`

### **`Flujo de Efectivo (Método Indirecto)`**
*   **Función:** Muestra cómo se movió el efectivo de la empresa.
*   **Requiere que el usuario defina un periodo (Fecha Inicio en `B1`, Fecha Fin en `C1`).**
*   **Utilidad Neta del Periodo**:
    *   `=SUMIFS(LibroDiario!F, LibroDiario!C, "4*", LibroDiario!B, ">="&$B$1, LibroDiario!B, "<="&$C$1) - SUMIFS(LibroDiario!E, LibroDiario!C, "4*", LibroDiario!B, ">="&$B$1, LibroDiario!B, "<="&$C$1) - (SUMIFS(LibroDiario!E, LibroDiario!C, "6*", LibroDiario!B, ">="&$B$1, LibroDiario!B, "<="&$C$1) - SUMIFS(LibroDiario!F, LibroDiario!C, "6*", LibroDiario!B, ">="&$B$1, LibroDiario!B, "<="&$C$1))`
*   **Ajustes por Actividades de Operación:**
    *   `Variación en Cuentas por Cobrar (Clientes)`: `=SUMIFS(LibroDiario!E, LibroDiario!C, "105*", LibroDiario!B, ">="&$B$1, LibroDiario!B, "<="&$C$1) - SUMIFS(LibroDiario!F, LibroDiario!C, "105*", LibroDiario!B, ">="&$B$1, LibroDiario!B, "<="&$C$1)`
    *   `Variación en Cuentas por Pagar (Proveedores)`: `=SUMIFS(LibroDiario!F, LibroDiario!C, "201*", LibroDiario!B, ">="&$B$1, LibroDiario!B, "<="&$C$1) - SUMIFS(LibroDiario!E, LibroDiario!C, "201*", LibroDiario!B, ">="&$B$1, LibroDiario!B, "<="&$C$1)`
*   **Flujo de Efectivo Neto de Actividades de Operación:** `=SUM(Utilidad Neta, Ajustes)`

## 4. Instrucciones de Uso

1.  **Poblar Catálogos:** Llena `CatalogoCuentas`, `CatalogoProveedores` y `CatalogoProdServ`.
2.  **Recargar Hoja:** Guarda los cambios y recarga la hoja para que aparezca el menú `Contabilidad Automatizada`.
3.  **Cargar XML:**
    *   Haz clic en `Cargar XML desde PC`.
    *   Selecciona **todos** tus archivos XML (Ingresos, Egresos y Pagos). El sistema los procesará en el orden correcto.
    *   Revisa las hojas `LibroDiario` y `CuentasPorCobrarPagar` para ver los resultados.
