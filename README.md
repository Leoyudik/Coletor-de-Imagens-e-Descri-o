  AUTOMAÇÃO DE COLETA DE PRODUTOS - FARMÁCIAS ONLINE
  Versão 1.0
================================================================================

DESCRIÇÃO DO PROJETO
--------------------------------------------------------------------------------
Este projeto automatiza a coleta de descrições e imagens de produtos
farmacêuticos a partir de dois sites:

  - Drogaria São Paulo  (https://www.drogariasaopaulo.com.br)
  - Panvel              (https://www.panvel.com)

A busca é feita pelo código EAN de cada produto, lido de uma planilha Excel.
As descrições encontradas são gravadas de volta na planilha, e as imagens
são baixadas e organizadas em pastas por EAN.


ESTRUTURA DE ARQUIVOS
--------------------------------------------------------------------------------
  automacao_farmacia.py   → Script principal de automação
  requirements.txt        → Lista de dependências Python
  README.txt              → Este arquivo
  produtos.xlsx           → Planilha com os EANs (você fornece)
  imagens/                → Pasta criada automaticamente com as imagens
    └── 7891024030820/
          7891024030820 - 1.jpg
          7891024030820 - 2.jpg
          ...


PLANILHA DE ENTRADA (produtos.xlsx)
--------------------------------------------------------------------------------
A planilha deve ter o seguinte formato a partir da linha 2 (linha 1 = cabeçalho):

  Coluna A → EAN do produto (ex: 7891024030820)
  Coluna B → Descrição Drogaria SP  (preenchida automaticamente)
  Coluna C → Descrição Panvel       (preenchida automaticamente)

Você pode alterar as colunas no topo do script (variáveis COLUNA_EAN,
COLUNA_DESC_DSP e COLUNA_DESC_PANVEL).


REGRAS DE FUNCIONAMENTO
--------------------------------------------------------------------------------
Descrições:
  - Para cada EAN, o script busca a descrição em cada site separadamente.
  - A descrição é preenchida na coluna correspondente ao site de origem.
  - Se o produto não for encontrado em um site, a célula fica em branco.

Imagens:
  - As imagens dos dois sites são combinadas em uma única lista.
  - Imagens duplicadas são detectadas pelo conteúdo (hash MD5) e ignoradas,
    mesmo que venham de URLs diferentes.
  - São baixadas no máximo 10 imagens únicas por EAN.
  - Para cada EAN é criada uma subpasta dentro de "imagens/".
  - Os arquivos são nomeados como: EAN - 1.jpg, EAN - 2.jpg, ...

Ordem de prioridade das imagens:
  1. Imagens da Drogaria São Paulo (aparecem primeiro)
  2. Imagens da Panvel (completam até o limite de 10)


COMO OS SITES SÃO ACESSADOS
--------------------------------------------------------------------------------
Drogaria São Paulo:
  Utiliza a API interna VTEX do site, que retorna os dados do produto
  diretamente em formato JSON, sem necessidade de abrir um browser.
  É o método mais rápido e estável.

Panvel:
  O site é construído em Angular (SPA - Single Page Application), ou seja,
  o conteúdo é carregado por JavaScript após a abertura da página.
  Por isso, o script utiliza o Playwright, que abre um browser real
  (Chromium) em segundo plano, executa o JavaScript e coleta os dados
  depois que a página termina de carregar.


PRÉ-REQUISITOS
--------------------------------------------------------------------------------
  - Python 3.10 ou superior
  - Pip (gerenciador de pacotes Python)
  - Conexão com a internet


INSTALAÇÃO
--------------------------------------------------------------------------------
Siga a ordem abaixo:

  Passo 1 — Instalar as dependências Python:

      pip install -r requirements.txt

  Passo 2 — Instalar o browser Chromium do Playwright:

      playwright install chromium

      ATENÇÃO: este passo é obrigatório e separado do pip install.
      O Playwright precisa baixar o browser (~150 MB) para funcionar.


COMO USAR
--------------------------------------------------------------------------------
  1. Coloque o arquivo "produtos.xlsx" na mesma pasta do script.
  2. Preencha a coluna A com os EANs que deseja pesquisar (um por linha).
  3. Execute o script:

         python automacao_farmacia.py

  4. Acompanhe o progresso no terminal.
  5. Ao final, a planilha será salva com as descrições preenchidas,
     e as imagens estarão na pasta "imagens/".

  DICA: A planilha é salva após cada EAN processado. Se o script for
  interrompido, o progresso já realizado não será perdido.


CONFIGURAÇÕES AVANÇADAS
--------------------------------------------------------------------------------
No topo do arquivo "automacao_farmacia.py" você pode alterar:

  PLANILHA_ENTRADA   → Nome do arquivo Excel de entrada
  COLUNA_EAN         → Letra da coluna que contém os EANs
  COLUNA_DESC_DSP    → Letra da coluna para descrição da Drogaria SP
  COLUNA_DESC_PANVEL → Letra da coluna para descrição da Panvel
  PASTA_IMAGENS      → Nome da pasta onde as imagens serão salvas
  MAX_IMAGENS        → Quantidade máxima de imagens por EAN (padrão: 10)

Para visualizar o browser do Playwright abrindo durante a execução,
localize esta linha no script:

  browser = pw.chromium.launch(headless=True)

E altere para:

  browser = pw.chromium.launch(headless=False)


POSSÍVEIS ERROS E SOLUÇÕES
--------------------------------------------------------------------------------
Erro: "Planilha não encontrada"
  → Certifique-se de que o arquivo .xlsx está na mesma pasta do script
    e que o nome em PLANILHA_ENTRADA está correto.

Erro: "playwright install" não reconhecido
  → Execute: python -m playwright install chromium

Erro: Produto não encontrado em um dos sites
  → O EAN pode não estar cadastrado naquele site. A célula ficará em branco
    e o script continuará para o próximo EAN normalmente.

Erro: Timeout na Panvel
  → Pode ser lentidão temporária no site. Tente aumentar o valor de
    timeout no script (parâmetro timeout=30000, em milissegundos).


================================================================================
