# Gerador de Newsletter em Google Sheets

Olá, bem vindo!

Esse é um gerador simples de newsletter a partir de informações em uma planilha do GSheets. O GSuite é muito prático por permitir colaboração em tempo real e armazenamento na nuvem, e é o Suite utilizado pela Universidade de São Paulo como padrão.

Esse gerador vasculha a planilha em busca de notícias indexadas em cada uma das folhas, organiza elas no formato da newsletter especificado nos arquivos em HTML, e a envia para o endereço secreto especificado dentro da própria planilha.

O gerador foi inicialmente desenvolvido para a composição semi-automática de newsletters do [ODEC-USP](http://odec.iri.usp.br).

## Requisitos básicos

O código, na forma como ele está, pressupõe a existência de pelo menos duas folhas na planilha: a folha "Instruções" - cuja principal responsabilidade dentro do código é estabelecer duas variáveis - e a folha "ODEC".

A folha "Instruções" contém as principais instruções de uso da planilha. Prefiro não colocá-las aqui porque elas são bem específicas da Newsletter do ODEC em si. O mais importante dessa planilha é que ela define duas variáveis:

| Variável | Definição | Nome do Intervalo |
| -------- | --------- | ----------------- |
| Nome da Newsletter | Identifica o nome da Newsletter. Esse nome será utilizado no assunto do email para identificá-lo. | nome |
| Beamer Secreto | Identifica o endereço de email secreto disponibilizado pelo serviço de newsletters e para o qual o arquivo final será enviado. Quando o provedor recebe uma mensagem nesse endereço secreto, ele o envia para a lista de distribuição. | beamer |

Por padrão, o código insere a data em que ele está sendo rodado após o nome da newsletter. Assim, caso nome dela seja "[NEWSLETTER] - Dia", o nome final será: "[NEWSLETTER] - Dia dd/mm/yyyy" (o espaço é adicionado por padrão no código).

Essas variáveis utilizam **intervalos nomeados** (*named ranges*) dentro da folha "Instruções". Isso significa que as informações deles podem estar em qualquer lugar da folha, desde que haja um intervalo nomeado com o respectivo nome nela.

O endreço de email secreto do beamer estar disponível na planilha **não é a melhor prática de segurança**. Como tudo o que for enviado para esse email tem o potencial de ser distribuído na forma de newsletter, ele deveria estar oculto no código. A razão pela qual ele foi mantido na planilha é por questões de portabilidade: caso algum dia o sistema de newsletter do ODEC venha a mudar, basta apenas alterar aquele endereço na planilha, sem necessidade de que alguém altere o código e arrisque quebrá-lo.

A folha "ODEC" contém um campo especial: o campo "Mensagens". Esse campo, na folha, está na posição `C6`, e serve para veicular uma mensagem do ODEC na Newsletter. Essa mensagem aparece antes das notícias, em um campo próprio que só deve ser inserido no corpo do arquivo caso haja de fato alguma mensagem a ser veículada. Esse campo, assim, funciona independentemente de existirem ou não notícias na folha "ODEC".

Nenhuma das informações aqui é absoluta. É plenamente possível, por exemplo, alterar o nome da folha "ODEC" na sua planilha. Apenas assegure-se de alterar todas as instância dela no código (não utilize substituição automática, pois podem haver outras instâncias de "ODEC" que não se refiram à planilha).

## Como instalar

O código deve estar vinculado à planilha em que ele será rodado. Dessa forma, evita-se que ele esteja vinculado a uma conta específica, mas sim a planilha, e permite-se que outras pessoas possam usá-lo mesmo depois da desativação da conta criadora.

Para instalar o código, crie uma nova planilha do Google Sheets, e depois vá em Ferramentas > Editor de Script. O primeiro arquivo gerado é um arquivo .GS (uma versão do JavaScript utilizado pelos apps do Google). Copie o conteúdo de "enviaNews.gs" e cole nele (substituindo tudo). Em seguida, crie novos arquivos HTML - um para cada um dos contidos nesse repositório - e cole os conteúdos. O nome do arquivo contendo o código GS não importa, mas os nomes dos arquivos HTML **importa**, pois eles são utilizados no código. Caso você deseje alterá-los, altere o código.

Feito isso, crie duas folhas: "Instruções" e "ODEC", e na primeira defina os intervalos nomados "nome" e "beamer". O código estará pronto para uso.

## Como funciona

O código da planilha vasculha cada aba através de informações, aplicando as seguintes regras:

1. Ignorar a aba "Instruções"
2. Na aba "ODEC", verficar se existe conteúdo na célula `C6`:
    - Se existir, pegar esse conteúdo e inserí-lo dentro do campo especificado no arquivo "msg-odec.html".
    - Se não existir, ignorar o campo.
3. Ainda na aba "ODEC", verificar se na célula `B9` (linha 9, coluna 2) existe conteúdo:
    - Se sim, buscar todo o conteúdo, a partir dessa linha, e formatá-lo nos campos do arquivo "news-odec.html".
    - Se não, ignorar o restante da aba e prosseguir.
4. Para todas as outras abas: verificar se na célula `B6`(linha 6, coluna 2) existe conteúdo:
    - Se sim, criar uma nova seção que levará o nome da folha. Essa seção indica o tema geral das notícias dela. Em seguida, para cada linha com conteúdo a partir da sexta, buscar a informação de cada coluna e formatá-la nos moldes do arquivo "news-entry.html".
    - Se não, ignorar a folha e seguir para a próxima.

Quando o código percorrer todas as folhas, ele terminará de formatar o arquivo final da newsletter em HTML. Em seguida, ele buscará as informações de `nome` e `beamer` a partir de intervalos nomeados na folha "Instruções", e chamará a função `MailApp.SendEmail` para enviar o email para o beamer secreto.

O código foi escrito com certa versatilidade em mente. Com exceção das folhas "Instruções" e "ODEC", é possível inserir quantas folhas se desejar, cada uma contendo uma temática diferente. Como o nome da seção depende do nome da folha, o usuário final tem liberdade para decidi-lo sem precisar alterar o código.

A escolha de onde verificar se existem notícias (células B6 e B9) foi deliberada, e pensada tendo em vista a formatação estrutural da folha. É possível alterá-las no código, mas caso o faça, lembre-se sempre de verficar se alterou também as informações presentes nos loops de `for`.



Se você chegou até aqui, obrigado por ler. ❤️
