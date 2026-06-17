# Automação de Coleta de Produtos — Farmácias Online

Script Python que coleta **descrições e imagens** de produtos farmacêuticos a partir do código EAN, consultando três sites simultaneamente e gravando os resultados em uma planilha Excel.

**Sites suportados:**
- [Drogaria São Paulo](https://www.drogariasaopaulo.com.br) — via API VTEX (sem browser)
- [Panvel](https://www.panvel.com) — via Playwright (SPA Angular)
- [Beleza na Web](https://www.belezanaweb.com.br) — via Playwright (com bypass Akamai)

---

## Estrutura de arquivos

```
Atual/
├── automacao_farmacia.py   # script principal
├── requirements.txt        # dependências Python
├── produtos.xlsx           # planilha de entrada (você fornece)
└── imagens/                # criada automaticamente
    └── 7891024030820/
        ├── 7891024030820-1.jpg
        ├── 7891024030820-2.jpg
        └── ...
```

---

## Planilha de entrada (`produtos.xlsx`)

A linha 1 é reservada para cabeçalhos (preenchidos automaticamente se estiverem vazios). Os EANs devem estar a partir da linha 2.

| Coluna | Conteúdo                          | Preenchimento       |
|--------|-----------------------------------|---------------------|
| A      | EAN do produto                    | Você fornece        |
| B      | Descrição — Drogaria São Paulo    | Automático          |
| C      | Descrição — Panvel                | Automático          |
| D      | Descrição — Beleza na Web         | Automático          |

> As colunas podem ser alteradas nas constantes `COLUNA_EAN`, `COLUNA_DESC_DSP`, `COLUNA_DESC_PANVEL` e `COLUNA_DESC_BLZ` no topo do script.

---

## Como os sites são acessados

**Drogaria São Paulo**
Utiliza a API interna VTEX, que retorna dados do produto em JSON sem abrir browser. É o método mais rápido e estável.

**Panvel**
Site construído em Angular (SPA). O Playwright abre um Chromium real, executa o JavaScript e lê os dados do JSON-LD após o carregamento da página. Tem fallback por extração direta do DOM.

**Beleza na Web**
Site protegido pelo Akamai Bot Manager. O script aquece a sessão visitando a home page primeiro para obter os cookies de sensor necessários (`bm_sv`, `_abck`). Se bloqueado, tenta reaquecer automaticamente uma vez. Os dados são extraídos via JSON-LD (`ProductGroup` / `Product`), e as URLs de imagem são normalizadas para resolução máxima no Cloudinary.

---

## Funcionamento

**Descrições**
- Cada site é consultado de forma independente.
- O HTML é removido automaticamente das descrições antes de gravar na planilha (tags, entidades como `&amp;`, espaços extras).
- Se o produto não for encontrado, a célula fica em branco.

**Imagens**
- As imagens dos três sites são combinadas em uma lista única, sem duplicatas por URL.
- Imagens com conteúdo idêntico são detectadas por hash MD5 e descartadas.
- São salvas no máximo **10 imagens únicas** por EAN, redimensionadas para **1000 × 1000 px**.
- Cada EAN gera uma subpasta dentro de `imagens/` com arquivos nomeados `EAN-1.jpg`, `EAN-2.jpg`, etc.

**Salvamento incremental**
A planilha é salva após cada EAN processado. Se o script for interrompido, o progresso já realizado não é perdido.

---

## Pré-requisitos

- Python 3.10 ou superior
- Pip
- Conexão com a internet

---

## Instalação

**1. Instalar dependências Python:**
```bash
pip install -r requirements.txt
```

**2. Instalar o browser Chromium do Playwright:**
```bash
playwright install chromium
```
> Este passo é obrigatório e separado do `pip install`. O Playwright precisa baixar o Chromium (~150 MB).

---

## Como usar

1. Coloque `produtos.xlsx` na mesma pasta do script.
2. Preencha a coluna A com os EANs desejados (um por linha, a partir da linha 2).
3. Execute:
   ```bash
   python automacao_farmacia.py
   ```
4. Acompanhe o progresso no terminal.
5. Ao final, a planilha estará atualizada e as imagens estarão em `imagens/`.

---

## Configurações

No topo de `automacao_farmacia.py`:

| Constante           | Padrão          | Descrição                                  |
|---------------------|-----------------|--------------------------------------------|
| `PLANILHA_ENTRADA`  | `produtos.xlsx` | Nome do arquivo Excel de entrada           |
| `COLUNA_EAN`        | `A`             | Coluna com os EANs                         |
| `COLUNA_DESC_DSP`   | `B`             | Coluna para descrição da Drogaria SP       |
| `COLUNA_DESC_PANVEL`| `C`             | Coluna para descrição da Panvel            |
| `COLUNA_DESC_BLZ`   | `D`             | Coluna para descrição da Beleza na Web     |
| `PASTA_IMAGENS`     | `imagens`       | Pasta raiz para salvar imagens             |
| `MAX_IMAGENS`       | `10`            | Limite de imagens únicas por EAN           |

O browser roda com `headless=False` (janela visível) por padrão, o que ajuda a contornar o Akamai. Para rodar sem interface gráfica, altere para `headless=True` e considere usar o pacote `playwright-stealth`.

---

## Solução de problemas

**"Planilha não encontrada"**
Verifique se o arquivo `.xlsx` está na mesma pasta do script e se o nome em `PLANILHA_ENTRADA` está correto.

**`playwright install` não reconhecido**
Execute: `python -m playwright install chromium`

**Produto não encontrado em um site**
O EAN pode não estar cadastrado naquele site. A célula ficará em branco e o script continuará normalmente.

**Timeout ou lentidão**
Aumente o valor de `timeout` nas chamadas `page.goto()` (em milissegundos, padrão `30000`).

**Akamai bloqueia a Beleza na Web repetidamente**
O script já tenta reaquecer a sessão automaticamente. Se persistir, aumente o `wait_for_timeout` na função `_blz_aquecer_sessao` ou adicione uma pausa maior entre EANs ajustando o `time.sleep` ao final do loop.
