# Instruções para uso do software

Este software foi projetado para ler e identificar dados importantes em arquivos como PDFs e planilhas, e gerar um relatório em formato .docx. Ele verifica rapidamente os documentos fornecidos e extrai informações relevantes para produzir um relatório personalizado.

---

## Instalando requisitos

- Primeiro requisito é o python, caso o python não esteja instalado em seu computador instale por meio desse link: ``https://www.python.org/downloads/``

- O segundo requisito é o pip, caso o pip não esteja presente em seu computador siga as instruções da documentação para fazer o download: ``https://pip.pypa.io/en/stable/installation/``

- O terceiro requisito é o virtualenv (utilizado para criar um ambiente virtual para o projeto), para fazer o download apenas digite ``pip install virtualenv`` em seu terminal.

- O último requisito é o git (utilizado para a conexão com o github), para fazer o download entre no link: ``https://git-scm.com/``.

Obs: se você não tem certeza se os requisitos estão baixados em seu computador utilize o seguinte comando em seu terminal: ``<nome do requisito> --version`` e se o retorno for a versão do requisito, ele já está instalado.

## Acessando o software pela primeira vez 

- Abra a pasta que deseja adicionar o projeto no seu terminal (adicione ``cd <path>`` no terminal).
- Entre no repositório do github no link: ``https://github.com/anaclaraddias/ReportMaker`` e clique no botão ``"<> code"``.
- Copie o link HTTPS.
- Entre novamente no terminal e digite ``git clone <link copiado do github>``
- Entre na pasta do projeto que foi criada pelo git.
- Utilize o comando ``virtualenv venv`` em seu terminal para criar o ambiente virtual para esse projeto.
- Utilize o comando ``source venv/Scripts/activate`` para o windows ou ``source venv/bin/activate`` para o mac. Isso activa o ambiente virtual criado para esse projeto.
- Utilize o comando ``pip install -r "requirements.txt"`` para que o pip instale todas as bibliotecas necessárias para o funcionamento do projeto.

## Adicionando os arquivos necessários

Dentro da pasta raiz do projeto existem duas outras pastas, A "analysis" e a "creation".

- Dentro da analysis coloque o arquivo pdf e a planilha stock.
- Dentro da planilha created, a unica coisa que precisa ser adicionada é a alteração da planilha criada pelo sistema (a planilha criada pelo sistema precisa de dados inseridos manualmente). Mas o nome e a hora de inserir essa planilha atualizada será explicada na proxima seção.

## Aprendendo a rodar o sistema

- Sempre que for preciso rodar o software, é necessário usar o seguinte comando no terminal ``python Docx.py``
- Quando o terminal retornar com um input, coloque o nome do arquivo pdf que esta dentro da pasta ``analysis`` (o nome precisa ser exatamente igual e sem o .pdf no final). Exemplo: ``ExecCompAlignedWithROIC_ModelPortfolio_Feb2023`` 
- A próxima vez que o terminal retornar com um input, coloque o nome da planilha stock esta dentro da pasta ``analysis`` (o nome também precisa ser exatamente igual e sem o .xlsx). Exemplo: ``AZO``
- Por último, o terminal vai retornar com outro input, dessa vez coloque o nome da planilha que foi criada pelo sistema (a planilha que esta dentro da pasta ``created``), mas a versão da planilha com os dados que precisam ser inseridos manualmente já dentro da planilha. Entre na planilha e adicione os dados necessários, salve ela com o nome que quiser, mas não coloque o mesmo nome que a planilha criada pelo sistema. Após salvar a nova versão da planilha dentro da pasta ``created``, adicione o nome escolhido para o arquivo no terminal (o nome também precisa ser exatamente igual e sem o .xlsx)
- Após seguir todos os passos da maneira correta, um arquivo .docx será criado na pasta ``created`` do projeto.

