/** 
 *                  GERADOR AUTOMATICO DE NEWSLETTERS
 *                        PELO GOOGLE SHEETS
 * 
 * Esse código limpa as folhas da planilha da newsletter após o envio
 * dela pelo código em 'enviaNews.gs'. Por precaução, ele é ativado
 * manualmente através de um botão na planilha.
 * 
 * @author Nathanael Rolim <nathanael.rolim@usp.br>
 * @version 2.0
 * @since 2.0
 * @license Apache-2.0
 * 
*/

function limpaFolha() {

  // Pega o arquivo ativo e a planilha ativa
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();

  // Pega informações da última linha (usado no loop)
  var rangeData = sheet.getDataRange();
  var lastRow = rangeData.getLastRow();

  // Define a linha inicial do Loop. Isso é feito fora do
  // argumento de "for" porque separaremos o valor dela para
  // a planilha "ODEC".
  let i = 6;

  if(sheet.getSheetName() == 'ODEC'){
     i = 9;

     // Limpa o conteúdo da mensagem
     sheet.getRange("mensagem").clearContent();
  };

  for(i; i <= lastRow; i++){
    
    // Limpa a linha i
    sheet.getRange(i,2,1,6).clearContent();

  };

}

function limpaPlanilha(){

  // Pega o arquivo ativo e todas as planilhas contidas nela
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var planilhas = ss.getSheets();

  for(var s in planilhas){

    // Passa por todas as planilhas
    var sheet = planilhas[s];
    var rangeData = sheet.getDataRange();
    var lastRow = rangeData.getLastRow();

    if(sheet.getSheetName() == 'Instruções'){
      // Ignora a planilha de Instruções, já que não queremos
      // deletar nada nela.
      continue;

    } else{

      // Para todas as outras planilhas, seguir o mesmo esquema
      // da função "limpaFolha"

      let i = 6;

      if(sheet.getSheetName() == 'ODEC'){
        i = 9;
        sheet.getRange("mensagem").clearContent();
      };

      for(i; i <= lastRow; i++){
        
        sheet.getRange(i,2,1,6).clearContent();

      };

    };
  };
}