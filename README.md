# Gerador de Newsletter em Google Sheets

Olá, bem vindo!

Esse é um gerador simples de newsletter a partir de informações em uma planilha do GSheets. O GSuite é muito prático por permitir colaboração em tempo real e armazenamento na nuvem, e é o Suite utilizado pela Universidade de São Paulo como padrão.

Esse gerador vasculha a planilha em busca de notícias indexadas em cada uma das folhas, organiza elas no formato da newsletter especificado nos arquivos em HTML, e a envia para o endereço secreto especificado dentro da própria planilha.

## Requisitos básicos

O código, na forma como ele está, pressupõe a existência de pelo menos duas folhas na planilha: a folha "Instruções" - cuja principal responsabilidade dentro do código é estabelecer duas variáveis - e a folha "ODEC".

A folha "Instruções" contém as principais instruções de uso da planilha. Prefiro não colocá-las aqui porque elas são bem específicas da Newsletter do ODEC em si. O mais importante dessa planilha é que ela define duas variáveis:

| Variável | Definição | Nome do Intervalo |
| -------- | --------- | ----------------- |
| Nome da Newsletter | Identifica o nome da Newsletter. Esse nome será utilizado no assunto do email para identificá-lo. | nome |
| Beamer Secreto | Identifica o endereço de email secreto disponibilizado pelo serviço de newsletters e para o qual o arquivo final será enviado. Quando o provedor recebe uma mensagem nesse endereço secreto, ele o envia para a lista de distribuição. | beamer |

Essas variáveis utilizam **intervalos nomeados** (*named ranges*) dentro da folha "Instruções". Isso significa que as informações deles podem estar em qualquer lugar da folha, desde que haja um intervalo nomeado com o respectivo nome nela.

O endreço de email secreto do beamer estar disponível na planilha **não é a melhor prática de segurança**. Como tudo o que for enviado para esse email tem o potencial de ser distribuído na forma de newsletter, ele deveria estar oculto no código. A razão pela qual ele foi mantido na planilha é por questões de portabilidade: caso algum dia o sistema de newsletter do ODEC venha a mudar, basta apenas alterar aquele endereço na planilha, sem necessidade de que alguém altere o código e arrisque quebrá-lo.

A folha "ODEC" contém um campo especial: o campo "Mensagens". Esse campo, na folha, está na posição `C6`, e serve para veicular uma mensagem do ODEC na Newsletter. Essa mensagem aparece antes das notícias, em um campo próprio que só deve ser inserido no corpo do arquivo caso haja de fato alguma mensagem a ser veículada. Esse campo, assim, funciona independentemente de existirem ou não notícias na folha "ODEC".

Nenhuma das informações aqui é absoluta. É plenamente possível, por exemplo, alterar o nome da folha "ODEC" na sua planilha. Apenas assegure-se de alterar todas as instância dela no código (não utilize substituição automática, pois podem haver outras instâncias de "ODEC" que não se refiram à planilha).
