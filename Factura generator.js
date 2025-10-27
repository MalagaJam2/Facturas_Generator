const ACCOUNTINGFOLDERID = "1IDBHXFsM_0wLpcNeED4i7a82fmN-B5Cj";
const EXCELTEMPLATEID = "1frVhP9gH9UU4fj2-RrS1xPoWJ6qfg4S7RD4BsngChME";

const INCOMEFOLDERNAME = "Ingresos";
const EXCELFOLDERNAME = "Excels";

const NUMBEROFPRODUCTS = 5;
const IS_RUNNING = "isRuning";
const EURFormat = new Intl.NumberFormat('es-ES', { style: 'currency', currency: 'EUR' });

let year;
let quarter;

const Quarters = {
  1: "T1 (Enero-Marzo)",
  2: "T2 (Abril-Junio)",
  3: "T3 (Julio-Septiembre)",
  4: "T4 (Octubre-Diciembre)",
};

//----------------------------------------------
// MAIN EXECUTE FUNCTION
//----------------------------------------------
function CreateNewInvoice() {
  const cache = CacheService.getScriptCache();
  while (cache.get(IS_RUNNING) === "true") {
    Utilities.sleep(3000);
  }

  cache.put(IS_RUNNING, "true"); // 2 minutos de bloqueo
  try {
    Log(`Comienzo del proceso`);
    SetCurrentQuarter();
    SetCurrentYear();

    const currentFolder = GetCurrentInvoiceFolder();
    if (!currentFolder)
    {
      throw new Error("No se pudo encontrar o crear la carpeta actual de facturas.");
    } 
    Log(`Carpeta actual: ${currentFolder.getName()}`);

    CreateFilesFromForm(currentFolder);
  } catch (error) {
    console.error(`[FacturaGenerator] Error: ${error}`);
  } finally {
    cache.remove(IS_RUNNING);
    Log(`Fin del proceso`);
  }
}

//----------------------------------------------
// UTILS Y FORMATEOS
//----------------------------------------------
function Log(msg) {
  console.log(`[FacturaGenerator] ${new Date().toISOString()} - ${msg}`);
}

function SetCurrentYear() {
  year = new Date().getFullYear();
  Log(`Año actual: ${year}`);
}

function SetCurrentQuarter() {
  const month = new Date().getMonth();
  quarter = Math.floor(month / 3) + 1;
}

function FormatDate(date) {
  const currentDate = new Date(date);
  const year = currentDate.getFullYear();
  const month = ("0" + (currentDate.getMonth() + 1)).slice(-2);
  const day = ("0" + currentDate.getDate()).slice(-2);
  return `${day}/${month}/${year}`;
}

//----------------------------------------------
// NUMERACIÓN FACTURAS
//----------------------------------------------
function GetNextInvoiceNumber(folder, isRectification = false) {
  let number = 1;
  while (!folder.getFiles().hasNext()) {
    if (quarter === 1) return GetNewInvoiceNumber(number, isRectification);
    quarter -= 1;
    folder = GetCurrentInvoiceFolder();
  }
  SetCurrentQuarter();
  Log(`Quarter actual: ${quarter}`);
  number = GetLastInvoiceNumber(folder, isRectification) + 1; //Devuelve 007
  return GetNewInvoiceNumber(number, isRectification); //Devuelve 24N007
}

function GetNewInvoiceNumber(number, isRectification = false) {
  const prefix = year.toString().substring(2);
  const letter = isRectification ? "R" : "N";
  return `${prefix}${letter}${number.toString().padStart(3, "0")}`;
}

function GetLastInvoiceNumber(folder, isRectification = false) {
  const folderFiles = folder.getFiles();
  let lastInvoiceNumber = -Infinity;
  while (folderFiles.hasNext()) {
    const file = folderFiles.next();
    const fileName = file.getName();
    const invoiceInfo = fileName.split(" ")[0];
    const invoiceNumber = invoiceInfo.substring(3);
    const letter = invoiceInfo.charAt(2);
    if (isRectification && letter !== "R") continue;
    if (!isRectification && letter !== "N") continue;
    const num = parseInt(invoiceNumber, 10);
    if (num > lastInvoiceNumber) lastInvoiceNumber = num;
  }
  return lastInvoiceNumber === -Infinity ? 0 : lastInvoiceNumber;
}

//----------------------------------------------
// GESTIÓN DE CARPETAS
//----------------------------------------------
function GetCurrentInvoiceFolder() {
  const incomeFolder = GetIncomeFolder();
  CreateFolderIfNotExists(incomeFolder, EXCELFOLDERNAME);
  const folders = incomeFolder.getFoldersByName(EXCELFOLDERNAME);
  return folders.hasNext() ? folders.next() : null;
}

function GetIncomeFolder() {
  const quarterFolder = GetQuarterFolder();
  const folders = quarterFolder.getFoldersByName(INCOMEFOLDERNAME);
  return folders.hasNext() ? folders.next() : null;
}

function GetQuarterFolder() {
  const contabilidadFolder = DriveApp.getFolderById(ACCOUNTINGFOLDERID);
  const yearFolderIter = contabilidadFolder.getFoldersByName(year);
  if (!yearFolderIter.hasNext())
  {
    throw new Error(`No existe carpeta del año ${year}`);
  } 
  const yearFolder = yearFolderIter.next();
  
  const quarterFolderIter = yearFolder.getFoldersByName(Quarters[quarter]);
  if (!quarterFolderIter.hasNext()) 
  {
    throw new Error(`No existe carpeta del trimestre ${Quarters[quarter]}`);
  }
  return quarterFolderIter.next();
}

function CreateFolderIfNotExists(folder, folderName) {
  const exists = FolderExists(folder.getFolders(), folderName);
  if (!exists) {
    Log(`Creando carpeta: ${folderName}`);
    folder.createFolder(folderName);
  }
}

function FolderExists(subFolders, folderName) {
  while (subFolders.hasNext()) {
    if (subFolders.next().getName() === folderName) return true;
  }
  return false;
}

//----------------------------------------------
// CREACIÓN Y PROCESADO DE ARCHIVOS
//----------------------------------------------

function ParseFormData(values, header) {
  var response_data = {};
  for (var i = 0; i < values.length; i++) {

    var key = header[i];
    var value = values[i];
    switch (key) {
      case 'Fecha':
        response_data[key] =  value == null || value === "" ? "" : FormatDate(value);
        break;
      case 'Numero Rectificacion':
        response_data[key] = (value == null || value === "") ? "" : `Rectificación de la factura: ${value} `;
        break;
      case 'Fecha Rectificacion':
        response_data[key] = (value == null || value === "") ? "" : `con fecha ${FormatDate(value)}`;
        break;
      case 'NIF':
        response_data[key] = (value == null || value === "") ? "" : "NIF: " + value;
        break;
      default:
        response_data[key] = value;
    }
  }

  var IVA = Number(response_data["IVA (%)"]);
  if (isNaN(IVA)) IVA = 0;

  if (response_data["¿IVA incluido en el precio?"] === "Si") {
    var importeTotal = 0;
    var importeTotalIVA = 0;

    for (var i = 1; i <= NUMBEROFPRODUCTS; i++) {
      var cantidad = Number(response_data[`Cantidad ${i}`]);
      var precioBruto = Number(response_data[`Precio ${i}`]);

      if (!isNaN(cantidad) && !isNaN(precioBruto) && response_data[`Cantidad ${i}`] !== "" && response_data[`Precio ${i}`] !== "") {
        var precioNeto = (precioBruto / (1 + (IVA / 100)));
        var importe = (cantidad * precioNeto);
        importeTotal += importe;
        importeTotalIVA += (cantidad * precioBruto);
        response_data[`Importe ${i}`] = EURFormat.format(importe.toFixed(2));
        response_data[`Precio ${i}`] = EURFormat.format(Number(precioNeto).toFixed(2));
      } else {
        response_data[`Importe ${i}`] = "";
        if (response_data[`Precio ${i}`] !== "") {
          // Si vino texto, intenta formatear igualmente
          var n = Number(response_data[`Precio ${i}`]);
          if (!isNaN(n)) response_data[`Precio ${i}`] = EURFormat.format(n.toFixed(2));
        }
      }
    }

    var iva = importeTotalIVA - importeTotal;
    response_data["Importe Neto"] = EURFormat.format(importeTotal.toFixed(2));
    response_data["IVA Precio"] = EURFormat.format(iva.toFixed(2));
    response_data["Importe Final"] = EURFormat.format(importeTotalIVA.toFixed(2));
    response_data["IVA Number"] = IVA;

  } else {
    var importeTotal2 = 0;

    for (var j = 1; j <= NUMBEROFPRODUCTS; j++) {
      var precio = Number(response_data[`Precio ${j}`]);
      var cantidad2 = Number(response_data[`Cantidad ${j}`]);

      if (!isNaN(cantidad2) && !isNaN(precio) && response_data[`Cantidad ${j}`] !== "" && response_data[`Precio ${j}`] !== "") {
        var importe2 = cantidad2 * precio;
        importeTotal2 += importe2;
        response_data[`Importe ${j}`] = EURFormat.format(importe2.toFixed(2));
        response_data[`Precio ${j}`] = EURFormat.format(precio.toFixed(2));
      } else {
        response_data[`Importe ${j}`] = "";
        if (response_data[`Precio ${j}`] !== "") {
          var n2 = Number(response_data[`Precio ${j}`]);
          if (!isNaN(n2)) response_data[`Precio ${j}`] = EURFormat.format(n2.toFixed(2));
        }
      }
    }

    var ivaPrecio = importeTotal2 * (IVA / 100);
    var importeFinal = importeTotal2 + ivaPrecio;
    response_data["Importe Neto"] = EURFormat.format(importeTotal2.toFixed(2));
    response_data["IVA Precio"] = EURFormat.format(ivaPrecio.toFixed(2));
    response_data["Importe Final"] = EURFormat.format(importeFinal.toFixed(2));
    response_data["IVA Number"] = IVA;
  }

  return response_data;
}

function CreateFilesFromForm(folder) {
  const [excelId, filename] = CreateExcel(folder);
  Log(`Excel creado: ${filename}`);

  Utilities.sleep(500); // Pequeño retraso para evitar latencia
  CreatePDF(excelId, filename);
  Log(`PDF creado: ${filename}`);
}

function CreateExcel(folder) {
  const sheet = SpreadsheetApp.getActiveSheet();
  const values = sheet.getDataRange().getValues();
  const headers = values[0];
  const isProcessedIndex = headers.indexOf("IsProcessed");

  // Busca la última fila no procesada desde abajo
  let rowToProcess = values.length - 1;
  while (rowToProcess > 0 && values[rowToProcess][isProcessedIndex] === "true") {
    rowToProcess--;
  }

  const rowData = values[rowToProcess];
  sheet.getRange(rowToProcess + 1, isProcessedIndex + 1).setValue("true");

  const response_data = ParseFormData(rowData, headers);
  const isRect = response_data["Numero Rectificacion"] && response_data["Numero Rectificacion"] !== "";

  const nextInvoiceNumber = GetNextInvoiceNumber(folder, isRect);
  const newExcelName = `${nextInvoiceNumber} ${isRect ? "Rectificación" : "Factura"} ${response_data["Nombre"]}`;
  response_data["invoiceNumber"] = nextInvoiceNumber;

  const excelTemplate = DriveApp.getFileById(EXCELTEMPLATEID);
  if (!excelTemplate)
  {
    throw new Error("No se encontró la plantilla Excel.");
  } 

  const excelCopy = excelTemplate.makeCopy(newExcelName, folder);
  const excelId = excelCopy.getId();
  const newExcel = SpreadsheetApp.openById(excelId);
  PopulateTemplate(newExcel, response_data);

  return [excelId, newExcelName];
}

//----------------------------------------------
// GENERAR PDF
//----------------------------------------------
function CreatePDF(excelId, filename) {  
  
  const fr = 0, fc = 0, lc = 9, lr = 27;
  const url =
    "https://docs.google.com/spreadsheets/d/" + excelId + "/export" +
    "?format=pdf&size=7&fzr=true&portrait=true&fitw=true&gridlines=false&" +
    "top_margin=0.5&bottom_margin=0.25&left_margin=0.5&right_margin=0.5&" +
    "printtitle=false&sheetnames=false&pagenum=UNDEFINED&attachment=true&" +
    `r1=${fr}&c1=${fc}&r2=${lr}&c2=${lc}`;

  const params = { 
    method: "GET", 
    headers: { authorization: "Bearer " + ScriptApp.getOAuthToken() } 
  };

  try {
    const blob = UrlFetchApp.fetch(url, params).getBlob().setName(filename + ".pdf");
    const folder = GetIncomeFolder();
    return folder.createFile(blob);
  } catch (e) {
    console.error(`[FacturaGenerator] Error al generar PDF: ${e}`);
    return null;
  }
}

//----------------------------------------------
// RELLENO DE PLANTILLA
//----------------------------------------------
function PopulateTemplate(excel, response_data) {
  for (const key in response_data) {
    const finder = excel.createTextFinder(`{{${key}}}`);
    finder.replaceAllWith(response_data[key]);
  }
}
