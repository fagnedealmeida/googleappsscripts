///////////////////////////////////////
//FUNÇÃO PARA BAIXAR GUIA DE PLANILHA//
///////////////////////////////////////
function baixarAbaPlanilha() {


var planilha = SpreadsheetApp.getActiveSpreadsheet();
var guiaDados = planilha.getSheetByName("COBRAR");
var ultimaLinha = guiaDados.getLastRow();


var area = guiaDados.getRange("B2:K" + ultimaLinha);


var guiaDadosXls = planilha.getSheetByName("DadosExcel");


if(guiaDadosXls == null){
  planilha.insertSheet("DadosExcel");
  var guiaDadosXls = planilha.getSheetByName("DadosExcel");


}else{
  guiaDadosXls.clear();
}


// Ações acima precisam ser executadas antes de continuar
SpreadsheetApp.flush();
//Codigo fica parado em tempo para que ocorra as ações.
Utilities.sleep(2000);


area.copyTo(guiaDadosXls.getRange("A1"),SpreadsheetApp.CopyPasteType.PASTE_FORMAT,false);
area.copyTo(guiaDadosXls.getRange("A1"),SpreadsheetApp.CopyPasteType.PASTE_VALUES,false);


guiaDadosXls.activate();
var getId = SpreadsheetApp.getActiveSpreadsheet().getSheetId();


var nomePlan = planilha.getName();


planilha.rename("Planilha Excel");


guiaDados.activate();
// Ações acima precisam ser executadas antes de continuar
SpreadsheetApp.flush();
//Codigo fica parado em tempo para que ocorra as ações.
Utilities.sleep(2000);


var url = planilha.getUrl().replace(/edit$/,'') + 'export?Format=xlsx' + "&gid=" + getId;


var html = "<script>window.open('"+url+"');google.script.host.close();</script>";


var userInterface = HtmlService.createHtmlOutput(html)


.setHeight(10)
.setWidth(120);


SpreadsheetApp.getUi().showModalDialog(userInterface,'Baixando em 10seg.');


// Ações acima precisam ser executadas antes de continuar
SpreadsheetApp.flush();
//Codigo fica parado em tempo para que ocorra as ações.
Utilities.sleep(10000)


var guias = planilha.getSheets();


for(i=0; i < guias.length; i++){
  var nomeGuia = guias[i].getSheetName();
  if(nomeGuia == "DadosExcel"){
    planilha.deleteSheet(guias[i]);


  }


}


planilha.rename(nomePlan);




}

