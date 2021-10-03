# Gerador de Newsletter em Google Sheets

Olá, bem vindo!

Esse é um gerador simples de newsletter a partir de informações em uma planilha do GSheets. O GSuite é muito prático por permitir colaboração em tempo real e armazenamento na nuvem, e é o Suite utilizado pela Universidade de São Paulo como padrão.

Esse gerador vasculha a planilha em busca de notícias indexadas em cada uma das folhas, organiza elas no formato da newsletter especificado nos arquivos em HTML, e a envia para o endereço secreto especificado dentro da própria planilha.

## Requisitos básicos

O código, na forma como ele está, pressupõe a existência de pelo menos duas folhas na planilha: a folha "Instruções" - cuja principal responsabilidade dentro do código é estabelecer duas variáveis - e a folha "ODEC".

A folha "Instruções" contém as principais instruções de uso da planilha. Prefiro não colocá-las aqui porque elas são bem específicas da Newsletter do ODEC em si. O mais importante dessa planilha é que ela define duas variáveis:
