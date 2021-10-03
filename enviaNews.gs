function enviaNewsletter() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var planilhas = ss.getSheets();

  // Header da Newsletter
  var htmlHeader = HtmlService.createTemplateFromFile("news-header.html");
  var newsHeader = htmlHeader.evaluate().getContent();

  // Footer da Newsletter
  var htmlFooter = HtmlService.createTemplateFromFile("news-footer.html");
  var newsFooter = htmlFooter.evaluate().getContent();

  // Inicia a variável do arquivo final
  var newsFinal = '';


  for(var s in planilhas){ // Passar por todas as planilhas
    var sheet = planilhas[s];
    
    // Ignorar a aba "Instruções"
    if(sheet.getSheetName() == "Instruções"){
      continue;
    }

    // Para a aba "ODEC"
    else if(sheet.getSheetName() == "ODEC"){

      // Pegar a mensagem inicial e, caso ela exista, formatá-la dentro da newsletter.
      var msgInicial = sheet.getRange(6,3).getValue();
      
      if(msgInicial != ""){
        
        // Insere a mensagem na Newsletter
        var htmlMensagem = HtmlService.createTemplateFromFile("msg-odec.html");
        htmlMensagem.msgInicial = msgInicial; // Passa a variável para o arquivo
        var newsMensagem = htmlMensagem.evaluate().getContent();

        var newsFinal = newsHeader + '\n' + newsMensagem;

      } else{
        var newsFinal = newsHeader;
      };

      var newsFinal = newsFinal + '\n' + '<div class="news-entries-wrapper">';

      // Pegar as publicações do ODEC listadas
      var rangeData = sheet.getDataRange();
      var lastRow = rangeData.getLastRow();

      var eBranco = sheet.getRange(9,2).getValue();
      if(eBranco != ""){

        var newsFinal = newsFinal + '\n<div class="news-entry-section">' + '\n<h2>No ODEC:</h2>';

        for(i = 9; i <= lastRow; i++){

          var cate = sheet.getRange(i,2); // Categoria
          var dataPub = sheet.getRange(i,3); // Data de Publicação
          var titulo = sheet.getRange(i,4); // Título da Matéria
          var autor = sheet.getRange(i,5); // Autor
          var resumo = sheet.getRange(i,6); // Resumo
          var link = sheet.getRange(i,7); // Link da Notícia

          var dados = // Formata os dados na forma de um JSON para que sejam passados para o arquivo
          {
            cate: cate.getValue(),
            dataPub: Utilities.formatDate(dataPub.getValue(), "GMT-3", "dd/MM/yyyy"),
            titulo: titulo.getValue(),
            autor: autor.getValue(),
            resumo: resumo.getValue(),
            link: link.getValue()
          };

          // Passa as publicações do ODEC para o arquivo
          var htmlODECEntry = HtmlService.createTemplateFromFile("news-odec.html");
          htmlODECEntry.dados = dados;
          var newsODECEntry = htmlODECEntry.evaluate().getContent();
          
          var newsFinal = newsFinal + '\n' + newsODECEntry;
        };

        var newsFinal = newsFinal + '\n</div>'

      } else{
        continue;
      }


    } else{ // Para todas as outras abas
        var rangeData = sheet.getDataRange();
        var lastRow = rangeData.getLastRow();

        var eBranco = sheet.getRange(6,2).getValue(); // Verifica se há realmente notícias. Não é a melhor forma, mas é um atalho.
        if(eBranco != ""){

          var newsFinal = newsFinal + '\n<div class="news-entry-section">' + '\n<h2>' + sheet.getSheetName() + '</h2>'; // Nome da seção é o nome da folha

          for(i = 6; i <= lastRow; i++){

            var jornal = sheet.getRange(i,2); // Nome do Jornal
            var dataPub = sheet.getRange(i,3); // Data de Publicação
            var titulo = sheet.getRange(i,4); // Título da Matéria
            var autor = sheet.getRange(i,5); // Autor
            var resumo = sheet.getRange(i,6); // Resumo
            var link = sheet.getRange(i,7); // Link da Notícia

            var dados = 
            {
              jornal: jornal.getValue(),
              dataPub: Utilities.formatDate(dataPub.getValue(), "GMT-3", "dd/MM/yyyy"),
              titulo: titulo.getValue(),
              autor: autor.getValue(),
              resumo: resumo.getValue(),
              link: link.getValue()
            }

            // Insere as notícias no arquivo
            var htmlEntry = HtmlService.createTemplateFromFile("news-entry.html");
            htmlEntry.dados = dados;
            var newsEntry = htmlEntry.evaluate().getContent();

            var newsFinal = newsFinal + '\n' + newsEntry;

          }

          var newsFinal = newsFinal + '\n</div>'
          

        } else{
          continue;
        }
    };

  };

  var newsFinal = newsFinal + '\n' + newsFooter; // Finaliza o arquivo adicionando o rodapé.

  var parInfo = ss.getSheetByName("Instruções");
  //Pega o nome da Newsletter definido na Planilha e o utiliza para formar o título
  var newsNome = parInfo.getRange('nome').getValue() + ' ' + Utilities.formatDate(new Date(), "GMT-3", "dd/MM/yyyy");
  var newsBeamer = parInfo.getRange('beamer').getValue() // Pega o endereço secreto

  MailApp.sendEmail({ // Envia o arquivo, no formato de email, para o endereço secreto.
    to: newsBeamer,
    subject: newsNome,
    htmlBody: newsFinal
  });

}
