/** 
 *                  GERADOR AUTOMATICO DE NEWSLETTERS
 *                        PELO GOOGLE SHEETS
 * 
 * Esse código lê as folhas de uma planilha do Google Sheets, formata ela em
 * um JSON e passa essa informação para o arquivo "newsletter.html", que
 * posteriormente será interpretado, gerando a newsletter, e disparado para um
 * email secreto.
 * 
 * @author Nathanael Rolim <nathanael.rolim@usp.br>
 * @version 2.0
 * @license Apache-2.0
 * 
*/

function enviaNewsletter(){

    // Pega o arquivo ativo e as planilhas nele contidas
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var planilhas = ss.getSheets();

    // Inicia o JSON com os campos "mensagem", "odec" e "conteudo"
    var newsFinal = {mensagem: "", odec: [], conteudo: []};

    // Deixa as variáveis de nome da Newsletter e Beamer secreto prontas para
    // receberem a informação já no loop (evita retrabalho)
    var newsNome = "";
    var newsBeamer ="";

    // Inicia o Loop pelas planilhas
    for ( var s in planilhas ){
        
        // Define a planilha atual como a ativa
        var sheet = planilhas[s];

        // Ignorar a aba "Instruções"
        if (sheet.getSheetName() == "Instruções"){
            var newsNome = sheet.getRange("nome").getValue() + ' ' + Utilities.formatDate(new Date(), "GMT-3", "dd/MM/yyyy");
            var newsBeamer = sheet.getRange("beamer").getValue();
        }

        // Para a aba ODEC, verificaremos se existem uma mensagem e os posts
        else if(sheet.getSheetName() == "ODEC"){

            var msgInicial = sheet.getRange("mensagem").getValue();

            // Permite que, mais pra frente, possamos passar o conteúdo da mensagem
            // como HTML. Isso é útil, por exemplo, para adicionar quebra de linhas.
            var msgInicial = HtmlService.createHtmlOutput(msgInicial).getContent();

            // Adiciona o campo "mensagem" ao JSON
            newsFinal.mensagem = msgInicial;

            // Pega as publicações listadas
            var rangeData = sheet.getDataRange();
            var lastRow = rangeData.getLastRow();

            // Pega a primeira célula da região que os dados devem estar e
            // checa se eles realmente existem
            var primCelula = sheet.getRange(9,2).getValue(); // B9

            if(primCelula != ""){

                // Passa por todas as linhas coletando as informações
                for (i = 9; i <= lastRow; i++){

                    var odecEntry = 
                    {
                        categoria: sheet.getRange(i,2).getValue(),
                        dataPub: Utilities.formatDate( sheet.getRange(i,3).getValue(), "GMT-3", "dd/MM/yyyy" ),
                        titulo: sheet.getRange(i,4).getValue(),
                        autor: sheet.getRange(i,5).getValue(),
                        resumo: sheet.getRange(i,6).getValue(),
                        link: sheet.getRange(i,7).getValue()
                    };

                    // Adiciona a entrada em "odec"
                    newsFinal.odec.push(odecEntry);

                };
            } else{
                continue;
            }

          // Para todas as outras planilhas
        } else{

            // Pega as publicações listadas na planilha
            var rangeData = sheet.getDataRange();
            var lastRow = rangeData.getLastRow();

            // Pega a primeira célula e verifica se ela está vazia. Caso sim,
            // vamos pular a folha
            var primCelula = sheet.getRange(6,2).getValue(); // B6

            if(primCelula != ""){

                // Define a variável tema, que indica a seção, com o nome da 
                // folha. Com isso, é possível mudar o nome da seção mudando
                // apenas o nome da folha
                var tema = sheet.getSheetName();

                // Define um JSON "parcial" que será passado para dentro de
                // "matérias" em nosso JSON final.
                var parcial =
                {
                    secao: tema,
                    materias: []
                }

                // Começando na sexta linha, passar por todas as linhas buscando
                // as informações
                for(i = 6; i <= lastRow; i++){

                    var entry = 
                    {
                        jornal: sheet.getRange(i,2).getValue(),
                        dataPub: Utilities.formatDate(sheet.getRange(i,3).getValue(), "GMT-3", "dd/MM/yyyy"),
                        titulo: sheet.getRange(i,4).getValue(),
                        autor: sheet.getRange(i,5).getValue(),
                        resumo: sheet.getRange(i,6).getValue(),
                        link: sheet.getRange(i,7).getValue()
                    }

                    // Envia o conteúdo de "entry" para dentro de "materias"
                    parcial.materias.push(entry);

                };

                // Envia "parcial" para dentro de "conteudo"
                newsFinal.conteudo.push(parcial);

            } else{
                continue;
            };

        };

    };

    // Agora vamos passar o JSON para o arquivo de template, que posteriormente
    // gerará a newsletter finalizada.
    var htmlTemplate = HtmlService.createTemplateFromFile("template.html");
    htmlTemplate.dados = newsFinal; // Passa o JSON
    var templateFinal = htmlTemplate.evaluate().getContent();

    // Substitui "&lt;" por "<" e "&gt;" por ">" para permitir que código HTML
    // inserido dentro das células possa ser interpretado. CUIDADO: aqui não
    // há filtro. Todo código vai ser interpretado.
    var templateFinal = templateFinal.replaceAll("&lt;","<");
    var templateFinal = templateFinal.replaceAll("&gt;",">");
    

    // Agora chamar o API do Gmail para enviar o template finalizado para o
    // beamer secreto

    MailApp.sendEmail({
        to: newsBeamer,
        subject: newsNome,
        htmlBody: templateFinal
    });

    // Código encerrado
}