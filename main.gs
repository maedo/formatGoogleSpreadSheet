/*
 FormatSpreadSheet
 https://github.com/maedo/formatGoogleSpreadSheet
 
 MIT License
 
 Copyright (c) 2020 Aedo Pino, Miguel Antonio
 
 Permission is hereby granted, free of charge, to any person obtaining a copy
 of this software and associated documentation files (the "Software"), to deal
 in the Software without restriction, including without limitation the rights
 to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
 copies of the Software, and to permit persons to whom the Software is
 furnished to do so, subject to the following conditions:
 
 The above copyright notice and this permission notice shall be included in all
 copies or substantial portions of the Software.
 
 THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
 AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
 SOFTWARE.
*/

function getFormatGoogleSpreadSheet(sheetName, Columns) {
  return new formatGSS_(sheetName,Columns);
}

class formatGSS_ {
  
  constructor(sheetName, columnList) {
    this.sheetName = sheetName || 'CIM';
    this.columnList = columnList || ['id'];
  }

  /**
  * Formatea la hoja dentro de la planilla.  
  */  
  formatSpreadSheet() {
    this.removeAll_(); // Borra todas las claves del documento.
    let sheet = this.getSheet_(); // obtiene el objeto hoja de la planilla
    this.setFormatSheet_(sheet);  // formatea las columnas de la hoja.
  }
  
  /**
  * Verifica que la hoja contenga las columnas necesarias
  * Si no las encuentra, las agrega al final de la lista de columnas.
  * Deja un registro de la posición de la columna dentro de la hoja  
  */
  setFormatSheet_(sheet) {
    const colNames = this.columnList;
    var userProperties = PropertiesService.getUserProperties(); 
    const rowInit = 1; // The row index of the cell to return; row indexing starts with 1.
    const colInit = 1; // The col index of the cell to return; col indexing starts with 1.
    var colEnd  = sheet.getLastColumn() != 0 ? sheet.getLastColumn() : 1;
    var range = sheet.getRange(rowInit,colInit,rowInit,colEnd); 
    var values = range.getValues()[0] != '' ? range.getValues()[0] : []; // lee la primera fila
    
    for (let colName of colNames){
      let index = values.indexOf(colName);
      if ( index == -1 ){ //colName not found
        values.push(colName); //agrega columna faltante al final de la fila
        index = values.length; //identifica su lugar
      }
      // identifica la columna en la que hay que almacenar la data
      // se suma 1 xq el valor del array inicia en 0 y la planilla en 1.
      userProperties.setProperty(colName, index+1); 
    }
    
    sheet.getRange(rowInit,colInit,rowInit,values.length).setValues([values]);
  }
  
  /**
  * verifica si existe la hoja. Si no existe la crea.
  * @return sheet
  */
  getSheet_(){
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    for ( let sheet of ss.getSheets() ) {
      if( sheet.getName() == this.sheetName ) {
        return sheet;
      }
    }
    
    return this.setSheet_();
  }
  
  /**
  * crea la hoja en la hoja de calculo.
  * @return sheet
  */
  setSheet_(){
    var ss = SpreadsheetApp.getActiveSpreadsheet();   
    return ss.insertSheet(this.sheetName);
  }
  
  /**
  * Remover todas las claves del documento.
  * @return Boolean
  */
  removeAll_() {
    var userProperties = PropertiesService.getUserProperties(); 
    userProperties.deleteAllProperties()
    return true;
  }  
  
}
