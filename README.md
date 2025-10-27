# 🧾 Factura Generator (Google Apps Script)

Este proyecto automatiza la **generación de facturas en Google Drive** a partir de respuestas de un **Google Form** conectado con una **Google Sheet**.  
Cuando se envía una nueva respuesta, el script crea automáticamente:

- Una **copia del Excel de plantilla**, con los datos del formulario reemplazados.
- Un **PDF** generado a partir del Excel.
- Un **número de factura secuencial** (normal o rectificativa).
- La factura en la carpeta correspondiente al trimestre y año actual.

---

## ⚙️ Funcionalidad principal

El script lee los datos más recientes de la hoja vinculada al Form, calcula importes e IVA, y genera automáticamente los archivos correspondientes.  

### Flujo del proceso

1. **El usuario envía el formulario** con los datos de la factura.  
2. **El trigger** ejecuta `CreateNewInvoice()`.  
3. El script:
   - Determina el año y trimestre actual.
   - Busca o crea la carpeta correcta (`Año > Trimestre > Ingresos > Excels`).
   - Asigna el siguiente número de factura (`24N001`, `24R001`, etc.).
   - Rellena una copia de la plantilla Excel (`EXCELTEMPLATEID`).
   - Sustituye los placeholders `{{Campo}}` en la plantilla.
   - Exporta el Excel a PDF y lo guarda en la carpeta `Ingresos`.

---

## 📁 Estructura de carpetas en Drive

```
📂 Contabilidad (ACCOUNTINGFOLDERID)
 ┣ 📂 2025
 ┃ ┣ 📂 T1 (Enero-Marzo)
 ┃ ┃ ┣ 📂 Ingresos
 ┃ ┃ ┃ ┣ 📂 Excels
 ┃ ┃ ┃ ┗ 📄 Facturas y PDFs generados
 ┃ ┣ 📂 T2 (Abril-Junio)
 ┃ ┣ 📂 T3 (Julio-Septiembre)
 ┃ ┗ 📂 T4 (Octubre-Diciembre)
```

---

## 🔢 Nomenclatura de facturas

- Factura normal: `YYN###` → Ejemplo: **24N007**
- Factura rectificativa: `YYR###` → Ejemplo: **24R002**

El contador se reinicia cada año/trimestre según la estructura de carpetas.

---

## 🧠 Lógica principal (`CreateNewInvoice`)

```javascript
function CreateNewInvoice() {
  // 1. Previene ejecuciones simultáneas con CacheService
  // 2. Detecta trimestre y año actual
  // 3. Busca carpeta destino
  // 4. Crea Excel + PDF desde la última fila no procesada del Sheet
}
```

---

## 🧩 Configuración inicial

Antes de ejecutar el script:

1. Abre la hoja de respuestas del formulario `Creación de facturas (respuestas).xlsx`
2. Ve a `Extensiones > Apps Script` y pega el contenido del archivo `Factura generator.js`.
3. Actualiza las siguientes constantes según tu entorno:

```javascript
const ACCOUNTINGFOLDERID = "ID_de_tu_carpeta_principal_en_Drive";
const EXCELTEMPLATEID = "ID_de_tu_plantilla_de_factura_en_Drive";
```
4. Asegúrate de que el Excel de plantilla contenga los **placeholders** con el formato:
   ```
   {{Nombre}}
   {{Fecha}}
   {{NIF}}
   {{Importe Neto}}
   {{IVA Precio}}
   {{Importe Final}}
   ...
   ```

5. En la hoja, la primera fila debe contener los **mismos nombres de campo** que el script espera:
   ```
   Fecha | Nombre | NIF | IVA (%) | ¿IVA incluido en el precio? | Precio 1 | Cantidad 1 | ... | IsProcessed
   ```

---


## 🧾 Archivos generados

| Tipo | Ubicación | Formato | Descripción |
|------|------------|----------|--------------|
| Excel | `/Año/Trimestre/Ingresos/Excels/` | `.xlsx` | Copia editable de la plantilla con datos reemplazados |
| PDF | `/Año/Trimestre/Ingresos/` | `.pdf` | Exportación en formato imprimible de la factura |

---

## 🧰 Utilidades incluidas

- `FormatDate(date)` → Devuelve `DD/MM/YYYY`.  
- `ParseFormData()` → Convierte los valores de la hoja en datos listos para plantilla.  
- `GetNextInvoiceNumber()` → Genera número correlativo con prefijo anual.  
- `PopulateTemplate()` → Sustituye placeholders en el Excel.  
- `CreatePDF()` → Exporta el Excel recién creado como PDF.

---

## 🧱 Seguridad y concurrencia

- Usa `CacheService` para evitar que dos ejecuciones simultáneas generen facturas duplicadas.
- Si otra instancia está corriendo, espera hasta 3 segundos y reintenta.

---

## 🧮 Variables principales

| Variable | Descripción |
|-----------|--------------|
| `ACCOUNTINGFOLDERID` | Carpeta raíz de contabilidad en Drive |
| `EXCELTEMPLATEID` | ID de la plantilla base de factura |
| `INCOMEFOLDERNAME` | Subcarpeta donde se guardan las facturas (`Ingresos`) |
| `EXCELFOLDERNAME` | Carpeta de los Excels generados (`Excels`) |
| `NUMBEROFPRODUCTS` | Número de líneas de producto soportadas por la plantilla |
| `IS_RUNNING` | Flag de ejecución concurrente |
| `EURFormat` | Formateador de moneda en euros |


---

**Autor:** Gonzalo Estrada Rojo  
**Versión:** 1.0 (Octubre 2025)  
**Entorno:** Google Apps Script + Google Sheets + Drive API
