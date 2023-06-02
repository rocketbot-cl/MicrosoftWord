



# Microsoft Word
  
Módulo para trabalhar com arquivos de texto usando o Microsoft Word. Crie e edite documentos do word, trabalhe com tabelas, formate seus textos e muito mais.  

*Read this in other languages: [English](README.md), [Português](README.pr.md), [Español](README.es.md)*

## Como instalar este módulo
  
Para instalar o módulo no Rocketbot Studio, pode ser feito de duas formas:
1. Manual: __Baixe__ o arquivo .zip e descompacte-o na pasta módulos. O nome da pasta deve ser o mesmo do módulo e dentro dela devem ter os seguintes arquivos e pastas: \__init__.py, package.json, docs, example e libs. Se você tiver o aplicativo aberto, atualize seu navegador para poder usar o novo módulo.
2. Automático: Ao entrar no Rocketbot Studio na margem direita você encontrará a seção **Addons**, selecione **Install Mods**, procure o módulo desejado e aperte instalar.  


## Overview


1. Novo Documento  
Criar um novo documento Word

2. Abrir documento  
Abra um documento do Word.

3. Ler documento  
Extraia o texto do documento Word.

4. Obter parágrafos  
Obtenha a lista de parágrafos que compõem um documento do Word no formato de dicionário {número: texto}.

5. Escrever no documento  
Escreva em um documento Word.

6. Copie e cole o texto  
Copie o texto entre os intervalos no documento do Word e cole-o em outro documento.

7. Copiar texto  
Copiar texto para prancheta entre intervalos no documento Word

8. Colar texto  
Colar texto da prancheta para documento Word

9. Contar caracteres  
Contar caracteres em um parágrafo específico

10. Adicionar tabela  
Adicionar tabela em um documento do Word.

11. Ler tabelas  
Extraia os dados das tabelas no documento

12. Editar tabela  
Editar uma tabela em um documento Word.

13. Atualizar campos vinculados  
Atualizar campos vinculados (ex. planilha do Excel)

14. Inserir página  
Inserir uma nova página no documento

15. Adicionar imagem  
Adicionar uma imagem ao documento

16. Converter para PDF  
Converter documento Word para PDF.

17. Localizar texto no parágrafo  
Localize o parágrafo onde se encontra o texto indicado.

18. Contar parágrafos  
Conte o número de parágrafos no documento.

19. Substituir texto no parágrafo  
Substituir o texto de um parágrafo.

20. Excluir parágrafo  
Excluir um parágrafo do documento. Se as tabelas forem incluídas, o comando Localizar texto no parágrafo deve ser usado para localizar o parágrafo a ser excluído. Retorna o texto deletado.

21. Adicionar texto a um bookmark  
Adicionar texto a um bookmark.

22. Salvar documento  
Salve o documento Word aberto

23. Fechar documento  
Feche o documento que está sendo executado  



### Changes
Thu Jul 21 01:32:22 2022  Merge branch qa into branch-nico

----
### OS

- windows

### Dependencies
- [**pywin32**](https://pypi.org/project/pywin32/)
### License
  
![MIT](https://camo.githubusercontent.com/107590fac8cbd65071396bb4d04040f76cde5bde/687474703a2f2f696d672e736869656c64732e696f2f3a6c6963656e73652d6d69742d626c75652e7376673f7374796c653d666c61742d737175617265)  
[MIT](http://opensource.org/licenses/mit-license.ph)