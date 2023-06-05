



# Microsoft Word
  
Módulo para trabalhar com arquivos de texto usando o Microsoft Word. Crie e edite documentos do word, trabalhe com tabelas, formate seus textos e muito mais.  

*Read this in other languages: [English](Manual_MicrosoftWord.md), [Português](Manual_MicrosoftWord.pr.md), [Español](Manual_MicrosoftWord.es.md)*
  
![banner](imgs/Banner_MicrosoftWord.png)
## Como instalar este módulo
  
Para instalar o módulo no Rocketbot Studio, pode ser feito de duas formas:
1. Manual: __Baixe__ o arquivo .zip e descompacte-o na pasta módulos. O nome da pasta deve ser o mesmo do módulo e dentro dela devem ter os seguintes arquivos e pastas: \__init__.py, package.json, docs, example e libs. Se você tiver o aplicativo aberto, atualize seu navegador para poder usar o novo módulo.
2. Automático: Ao entrar no Rocketbot Studio na margem direita você encontrará a seção **Addons**, selecione **Install Mods**, procure o módulo desejado e aperte instalar.  


## Descrição do comando

### Novo Documento
  
Criar um novo documento Word
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Sessão|sessão de arquivo|Word1|

### Abrir documento
  
Abra um documento do Word.
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Arquivo|Abra o documento especificado|arquivo.docx|
|Sessão|sessão de arquivo|Word1|

### Ler documento
  
Extraia o texto do documento Word.
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Resultado|Armazenar o resultado em uma variável|Variável|
|Sessão|sessão de arquivo|Word1|
|Adicionar detalhes|Escolha se os dados armazenados serão salvos com detalhes como estilo, alinhamento, etc.|True|

### Obter parágrafos
  
Obtenha a lista de parágrafos que compõem um documento do Word no formato de dicionário {número: texto}.
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Sessão|sessão de arquivo|Word1|
|Resultado|Armazenar o resultado em uma variável|Variável|

### Escrever no documento
  
Escreva em um documento Word.
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Sessão|sessão de arquivo|Word1|
|Escrever texto|Texto a ser escrito no documento|Lorem ipsum |
|Número do parágrafo - Opcional|Número do parágrafo de referência para inserir o texto|1|
|Inserir método - Opcional|Método a ser usado para inserir o novo texto||
|Tipo de texto|Seletor de tipo de texto que terá o texto escrito.|Subtitle|
|Nível|Nível que o texto escrito terá.|1-9|
|Tamanho da fonte|Tamanho da fonte que o texto escrito terá.|12|
|Alinhamento|Alinhamento que o texto escrito terá.|Left|
|Cor do texto|Cor que o texto escrito terá|Black|
|Negrito|Selecione se o texto ficará em negrito.|True|
|Itálico|Selecione se o texto ficará em itálico.|True|
|Sublinhar|Selecione se o texto será sublinhado.|False|

### Copie e cole o texto
  
Copie o texto entre os intervalos no documento do Word e cole-o em outro documento.
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Início do intervalo|Posição do intervalo de onde o comando começa a copiar.|0|
|fim do intervalo|Posição do intervalo para o qual o comando copia.|40|
|Sessão do arquivo a ser copiado|sessão de arquivo|Word1|
|Arquivo|Escolha o documento onde o conteúdo copiado é colado.|arquivo.docx|

### Copiar texto
  
Copiar texto para prancheta entre intervalos no documento Word
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Início do intervalo|Posição do intervalo de onde o comando começa a copiar.|0|
|Fim do intevalo|Posição do intervalo para o qual o comando copia.|40|
|Sessão|sessão de arquivo|Word1|

### Colar texto
  
Colar texto da prancheta para documento Word
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Sessão|sessão de arquivo|Word1|

### Contar caracteres
  
Contar caracteres em um parágrafo específico
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Sessão|sessão de arquivo|Word1|
|Parágrafo|Parágrafo para contar caracteres|1|
|Result|Armazenar o resultado em uma variável|Variável|

### Adicionar tabela
  
Adicionar tabela em um documento do Word.
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Numero de linhas|Número de linhas que a tabela terá|3 |
|Numero de colunas|Número de colunas que a tabela terá|4 |
|Estilo da tabela|Estilo de tabela padrão do Microsoft Word|Colorful Grid|
|Sessão|sessão de arquivo|Word1|
|Estilos de borda|Estilo de borda de tabela. Tipo e tamanho da linha.|Line type: Single wavy / Line size: 1 1/2 points|

### Ler tabelas
  
Extraia os dados das tabelas no documento
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Tabela para ler|Número da tabela a partir da qual o conteúdo será lido|1|
|Sessão|sessão de arquivo|Word1|
|Result|Armazenar o resultado em uma variável|Variável|

### Editar tabela
  
Editar uma tabela em um documento Word.
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Número da tabela|Número da tabela a ser editada|1|
|Sessão|sessão de arquivo|Word1|
|Digite o número da linha para excluir|Opcional. O número da linha inserido determina qual linha será removida da tabela.| |
|Digite o número da coluna para excluir|Opcional. O número da coluna inserido determina qual coluna será removida da tabela.| |
|Inserir linha|Se selecionado, adiciona uma linha ao final da tabela|True|
|Inserir coluna|Se selecionado, adiciona uma coluna ao final da tabela|False|
|Largura da coluna|Largura em pontos que cada coluna da tabela terá|140|
|Altura da linha|Altura em pontos que cada linha da tabela terá|25|

### Atualizar campos vinculados
  
Atualizar campos vinculados (ex. planilha do Excel)
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Número da campo|Número do campo a ser atualizado|1|
|Sessão|sessão de arquivo|Word1|

### Inserir página
  
Inserir uma nova página no documento
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Sessão|sessão de arquivo|Word1|

### Adicionar imagem
  
Adicionar uma imagem ao documento
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Sessão|sessão de arquivo|Word1|
|Rota da imagem|Direção da imagem que será adicionado abaixo do último parágrafo|imagem.jpg|

### Converter para PDF
  
Converter documento Word para PDF.
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Sessão|sessão de arquivo|Word1|
|Salvar arquivo|Caminho do arquivo onde o PDF será criado|arquivo.pdf|

### Localizar texto no parágrafo
  
Localize o parágrafo onde se encontra o texto indicado.
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Sessão|sessão de arquivo|Word1|
|Texto para pesquisar|Texto que será usado para localizar o parágrafo|Olá mundo|
|Nome variável|Armazenar o resultado em uma variável|Variável|

### Contar parágrafos
  
Conte o número de parágrafos no documento.
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Sessão|sessão de arquivo|Word1|
|Nome variável|Armazenar o número de parágrafos em uma variável|Variável|

### Substituir texto no parágrafo
  
Substituir o texto de um parágrafo.
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Sessão|sessão de arquivo|Word1|
|Texto para pesquisar|Texto a ser pesquisado nos parágrafos listados.|Olá mundo|
|Texto a substituir|Texto a ser substituído|Olá mundo|
|Lista de parágrafos|Parágrafos onde o texto especificado será pesquisado|Exemplo ',' separado por vírgula: 1,2|

### Excluir parágrafo
  
Excluir um parágrafo do documento. Se as tabelas forem incluídas, o comando Localizar texto no parágrafo deve ser usado para localizar o parágrafo a ser excluído. Retorna o texto deletado.
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Sessão|sessão de arquivo|Word1|
|Número do parágrafo|Número do parágrafo a ser excluído|1|
|Nome da variável onde o parágrafo excluído será salvo|Variável onde será salvo o texto que incluiu o parágrafo excluído|Variável|

### Adicionar texto a um bookmark
  
Adicionar texto a um bookmark.
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Sessão|sessão de arquivo|Word1|
|Texto a adicionar|Texto que será adicionado ao marcador escolhido.|Olá mundo|
|Nome do marcador|Nome do marcador onde o texto será adicionado.|Marcador 1|

### Salvar documento
  
Salve o documento Word aberto
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Sessão|sessão de arquivo|Word1|
|Salvar arquivo|Salve o arquivo no caminho especificado|arquivo.docx|

### Fechar documento
  
Feche o documento que está sendo executado
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Sessão|sessão de arquivo|Word1|
