const ACCOUNTINGFOLDERID = "1IDBHXFsM_0wLpcNeED4i7a82fmN-B5Cj"; 
const EXCELTEMPLATEID = "1frVhP9gH9UU4fj2-RrS1xPoWJ6qfg4S7RD4BsngChME";

// Folder names
const INCOMEFOLDERNAME = "Ingresos";
const EXCELFOLDERNAME = "Excels";

const NUMBEROFPRODUCTS = 5;
const IS_RUNNING = "isRuning";
const EURFormat = new Intl.NumberFormat('es-ES', { style: 'currency', currency: 'EUR' });

let year;
let quarter;

var Quarters = {
  1 : "T1 (Enero-Marzo)",
  2 : "T2 (Abril-Junio)",
  3 : "T3 (Julio-Septiembre)",
  4 : "T4 (Octubre-Diciembre)",
};

//MAIN EXECUTE FUNCTION
function CreateNewInvoice(){
  try{
    console.log(`Comienzo del proceso: ${new Date()}`);
    var isRuning = CacheService.getScriptCache().get(IS_RUNNING);
    CacheService.getScriptCache().put(IS_RUNNING, "false");

    while(isRuning === "true"){
      Utilities.sleep(5000);
      isRuning = CacheService.getScriptCache().get(IS_RUNNING);
    }
    
    CacheService.getScriptCache().put(IS_RUNNING, "true");
    SetCurrentQuarter();
    SetCurrentYear();
    let currentFolder = GetCurrentInvoiceFolder();
    console.log(`Carpeta actual encontrada: ${currentFolder.getName()}`)

    CreateFilesFromForm(currentFolder);
  }
  catch (error) {
    console.error(`Ha ocurrido un error: ${error}`);
  }
  finally{
    CacheService.getScriptCache().put(IS_RUNNING, "false");
    console.log(`Fin del proceso: ${new Date()}`);
  }
}

function SetCurrentYear(){
  year = new Date().getFullYear();
  console.log(`Año actual: ${year}`)
}

function SetCurrentQuarter(){
  let month = new Date().getMonth();
  quarter = parseInt(month/3) + 1;
  console.log(`Quarter actual: ${quarter}`)

}

function FormatDate(date){
  var currentDate = new Date(date);
  var year = currentDate.getFullYear();
  var month = ("0" + (currentDate.getMonth() + 1)).slice(-2);
  var day = ("0" + currentDate.getDate()).slice(-2);
  return `${day}/${month}/${year}`;
}

function GetNextInvoiceNumber(folder){
  let number = 1;
  while(!folder.getFiles().hasNext()){
    if(quarter == 1 ) {
      return GetNewInvoiceNumber(number)
    }
    else{
      quarter -= 1;
    }
    folder = GetCurrentInvoiceFolder();
  }
  SetCurrentQuarter()
  number = GetLastInvoiceNumber(folder)+1; //Devuelve 007
  
  return GetNewInvoiceNumber(number); //Devuelve 24N007
}

function GetNextRectInvoiceNumber(folder){
  let number = 1;
  while(!folder.getFiles().hasNext()){
    if(quarter == 1 ) {
      return GetNewInvoiceNumber(number, true)
    }
    else{
      quarter -= 1;
    }
    folder = GetCurrentInvoiceFolder();
  }
  SetCurrentQuarter()
  number = GetLastInvoiceNumber(folder, true)+1; //Devuelve 007
  
  return GetNewInvoiceNumber(number, true); //Devuelve 24R007
}

function GetNewInvoiceNumber(number, isRectification = false){
  if(isRectification)
  {
    return year.toString().substring(2) + "R" + number.toString().padStart(3, '0'); //Devuelve 24R007
  }
  else
  {
    return year.toString().substring(2) + "N" + number.toString().padStart(3, '0'); //Devuelve 24N007
  }
}

function GetLastInvoiceNumber(folder, isRectification = false){
  let folderFiles = folder.getFiles();
  let lastInvoiceNumber = -Infinity;
  while(folderFiles.hasNext()){
    let file = folderFiles.next();
    let fileName = file.getName();
    let invoiceInfo = fileName.split(" ")[0];
    let invoiceNumber = invoiceInfo.substring(3); // 24N.... coge el indice donde empieza el ...
    let letter = invoiceInfo.charAt(2); // 24N.... coge la N o R
    if(isRectification && letter != "R"){
      continue;
    }
    if(!isRectification && letter != "N"){
      continue;
    }
    if(parseInt(invoiceNumber,10) > lastInvoiceNumber){
      lastInvoiceNumber = parseInt(invoiceNumber,10);
    }
  }
  return lastInvoiceNumber;
}

function GetCurrentInvoiceFolder(){
  let incomeFolder = GetIncomeFolder();
  CreateFolderIfNotExists(incomeFolder, EXCELFOLDERNAME);
  return incomeFolder.getFoldersByName(EXCELFOLDERNAME).next();
}

function GetIncomeFolder(){
  let quarterFolder = GetQuarterFolder(year, quarter);
  return quarterFolder.getFoldersByName(INCOMEFOLDERNAME).next();
}

function GetQuarterFolder(){
  let contabilidadFolder = DriveApp.getFolderById(ACCOUNTINGFOLDERID);
  let yearFolder = contabilidadFolder.getFoldersByName(year).next();
  return yearFolder.getFoldersByName(Quarters[quarter]).next();
}

function CreateFolderIfNotExists(folder, folderName){  
  var subFolders = folder.getFolders();
  if(FolderExists(subFolders, folderName)){
    console.log(`Carpeta ${folderName} no existe en la carpeta ${folder.getName()}. Creando la carpeta`)
    folder.createFolder(folderName);
  }
}

function FolderExists(subFolders, folderName){
  while(subFolders.hasNext()){
    if(subFolders.next().getName() == folderName){
      return false;
    }
  }  
  return true;
}

function ParseFormData(values, header) {
    var response_data = {};
    for (var i = 0; i < values.length; i++) {

      var key = header[i];
      var value = values[i];
      switch (key) {
        case 'Fecha':
          response_data[key] =  FormatDate(value);
          break;
        case 'Numero Rectificacion':
          response_data[key] = value == "" || value == null ? "" : `Rectificación de la factura: ${value} `;
          break;
        case 'Fecha Rectificacion':
            response_data[key] = value == "" || value == null ? "" :  `con fecha ${FormatDate(value)}`;
            break;
        case 'NIF':
          response_data[key] = value == "" || value == null ? "" : "NIF: " + value;
          break;
        default:
          response_data[key] = value;
      }
    }
    
    var IVA = Number(response_data["IVA (%)"]);
    if(response_data["¿IVA incluido en el precio?"] == "Si"){
      var importeTotal = 0;
      var importeTotalIVA = 0;
      for(var i = 1; i <= NUMBEROFPRODUCTS; i++ ){
      
        var cantidad = response_data[`Cantidad ${i}`];
        var precioBruto = response_data[`Precio ${i}`];
        var precioNeto = (precioBruto/(1+(IVA/100)));

        if(cantidad != "" && precioBruto != "" ){
           var importe =  (Number(cantidad) * Number(precioNeto));
          importeTotal += importe;
          importeTotalIVA += Number(cantidad) * Number(precioBruto)
          response_data[`Importe ${i}`] = EURFormat.format(importe.toFixed(2));
        }
        else{
          response_data[`Importe ${i}`] = "";
        }
        if(response_data[`Precio ${i}`] != ""){
          response_data[`Precio ${i}`] = EURFormat.format(Number(precioNeto).toFixed(2));
        }
      }
      var iva = importeTotalIVA - importeTotal;
      response_data["Importe Neto"] = EURFormat.format(importeTotal.toFixed(2));
      response_data["IVA Precio"] = EURFormat.format(iva.toFixed(2));
      response_data["Importe Final"] = EURFormat.format(importeTotalIVA.toFixed(2));
      response_data["IVA Number"] = IVA ;      
    }
    else{
      var importeTotal = 0;
      for(var i = 1; i <= NUMBEROFPRODUCTS; i++ ){
      
        var precio = response_data[`Precio ${i}`];
        var cantidad = response_data[`Cantidad ${i}`];

        if(cantidad != "" && precio != "" ){
          var importe =  Number(cantidad) * Number(precio);
          importeTotal += importe;
          response_data[`Importe ${i}`] = EURFormat.format(importe.toFixed(2));
        }
        else{
          response_data[`Importe ${i}`] = "";
        }

        if(response_data[`Precio ${i}`] != ""){
          response_data[`Precio ${i}`] = EURFormat.format(Number(response_data[`Precio ${i}`]).toFixed(2));
        }
      }
      var ivaPrecio = importeTotal * (IVA/100);
      var importeFinal = importeTotal + ivaPrecio;    
      response_data["Importe Neto"] = EURFormat.format(importeTotal.toFixed(2));
      response_data["IVA Precio"] = EURFormat.format(ivaPrecio.toFixed(2));
      response_data["Importe Final"] = EURFormat.format(importeFinal.toFixed(2));
      response_data["IVA Number"] = IVA;
    } 
    return response_data;
}

function CreateFilesFromForm(folder) {
  let [excelId, filename] = CreateExcel(folder);
  console.log(`Excel creado con nombre: ${filename}`)

  Utilities.sleep(500); // Using to offset any potential latency in creating .pdf
  CreatePDF(excelId, filename);
  console.log(`PDF creado con nombre: ${filename}`)

}

function CreateExcel(folder){
  var sheet = SpreadsheetApp.getActiveSheet();
  
  var range = sheet.getDataRange(); 
  var headers = range.getValues()[0];
  var rowToProcess = sheet.getLastRow();

  var isProcessedIndex = headers.indexOf("IsProcessed");
  var data = range.getValues()[rowToProcess - 1];

  while(data[isProcessedIndex]){
    rowToProcess -= 1;
    data = range.getValues()[rowToProcess - 1];
  }

  var cellIsProcessed = ColumnToLetter(isProcessedIndex + 1, rowToProcess);
  sheet.getRange(cellIsProcessed).setValue("true");

  var response_data = ParseFormData(data, headers);
  
  if(response_data["Numero Rectificacion"] != "" && response_data["Numero Rectificacion"] != null)
  {  
    let nextRectInvoiceNumber = GetNextRectInvoiceNumber(folder);
    console.log(`Nuevo numero de factura rectificativa: ${nextRectInvoiceNumber}`)
    var newExcelName = `${nextRectInvoiceNumber} Rectificación ${response_data["Nombre"]}`;
    response_data["invoiceNumber"] = nextRectInvoiceNumber;
  }
  else
  {
    let nextInvoiceNumber = GetNextInvoiceNumber(folder);
    console.log(`Nuevo numero de factura: ${nextInvoiceNumber}`)
    var newExcelName = `${nextInvoiceNumber} Factura ${response_data["Nombre"]}`;
    response_data["invoiceNumber"] = nextInvoiceNumber;
  }

  var excelTemplate = DriveApp.getFileById(EXCELTEMPLATEID);
  var excelCopy = excelTemplate.makeCopy(newExcelName, folder);
  var excelId = excelCopy.getId();
  var newExcel = SpreadsheetApp.openById(excelId);
  PopulateTemplate(newExcel, response_data);
  return [excelId, newExcelName];
}

function ColumnToLetter(column, row) {
  var temp, letter = '';
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter + row;
}

function CreatePDF(excelId, filename) {
  const fr = 0, fc = 0, lc = 9, lr = 27;
  const url = "https://docs.google.com/spreadsheets/d/" + excelId + "/export" +
    "?format=pdf&" +
    "size=7&" +
    "fzr=true&" +
    "portrait=true&" +
    "fitw=true&" +
    "gridlines=false&" +
    "printtitle=false&" +
    "top_margin=0.5&" +
    "bottom_margin=0.25&" +
    "left_margin=0.5&" +
    "right_margin=0.5&" +
    "sheetnames=false&" +
    "pagenum=UNDEFINED&" +
    "attachment=true&" +
    "r1=" + fr + "&c1=" + fc + "&r2=" + lr + "&c2=" + lc;

  const params = { method: "GET", headers: { "authorization": "Bearer " + ScriptApp.getOAuthToken() } };
  const blob = UrlFetchApp.fetch(url, params).getBlob().setName(filename + '.pdf');

  // Gets the folder in Drive where the PDFs are stored.
  const folder = GetIncomeFolder();
  const pdfFile = folder.createFile(blob);
  return pdfFile;
}


function PopulateTemplate(excel, response_data) {
  for (var key in response_data) {
    var textFinder = excel.createTextFinder(`{{${key}}}`);
    textFinder.replaceAllWith(response_data[key]);
  }
}