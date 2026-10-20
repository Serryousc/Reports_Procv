# Cruzamento PROCV

Automatiza a conferência que era feita manualmente com PROCV no Excel. O script verifica se os códigos de uma planilha existem na coluna de ID de outra e gera um relatório curto em TXT.

## Como funciona

1. Lê todos os arquivos Excel (`.xlsx`, `.xlsm`, `.xls`) de uma pasta.
2. Separa os arquivos em dois grupos:
   - **Arquivo 1:** sem "id" no nome; tem a coluna `codigo`.
   - **Arquivo 2:** com "id" no nome; tem uma ou mais colunas com "ID" no cabeçalho.
3. Forma os pares comparando as partes dos nomes. Para isso, ignora a data, o prefixo `procv` e conectores como "e", "de" e "para".
   - Um Arquivo 1 pode ser comparado com vários Arquivos 2, e vice-versa.
   - Palavras parecidas contam como iguais, por exemplo "entregas" e "entrega".
   - Números como `2.1` precisam ser exatamente iguais.
4. Para cada par, verifica quais códigos do Arquivo 1 não aparecem em cada coluna ID do Arquivo 2. Em seguida, grava uma linha no relatório.

## Instalação

Requer Python 3.9 ou mais recente.

```bash
pip install -r requirements.txt
```

## Uso

```bash
python verificar_pasta.py "C:\caminho\da\pasta"
```

| Opção | O que faz |
|---|---|
| `--so-pares` | Mostra os pares encontrados, sem comparar. É útil para conferir antes de rodar. |
| `--limite 70` | Define a semelhança mínima entre os nomes para formar um par. O padrão é 60%. |
| `--saida arquivo.txt` | Muda o arquivo de saída. O padrão é `relatorio_nao_encontrados.txt`, dentro da pasta. |
| `--acrescentar` | Adiciona ao relatório existente em vez de começar um novo. |
| `--ignorar "procv,outro"` | Define as palavras desconsideradas nos nomes dos arquivos. |

## Relatório gerado

```
CRUZAMENTO 20/10/2026

1 - 2.1: 1 card com erro - Realizar Conferência Manual
1 - 2.2: sem erro
5 - 1.1: coluna de ID vazia
Entregas Especiais - Vidros: sem erro
```

Outros status possíveis:
- `sem coluna de ID` ou `sem coluna codigo`
- `coluna codigo vazia`
- `erro ao abrir os arquivos`

Nesses casos, os detalhes aparecem só no terminal.

## Regras da comparação

- **Códigos numéricos:** contam como iguais mesmo com zeros à esquerda ou formatados como texto ("00123" = 123), como faz o "Texto para colunas" do Excel.
- **Detalhes de formato:** espaços nas pontas, caracteres invisíveis e diferenças entre maiúsculas e minúsculas são ignorados.
- **Repetições e vazios:** códigos repetidos contam uma vez só, e células vazias são ignoradas.
- **Colunas de ID:** "ID" precisa aparecer como palavra no cabeçalho, como em `ID`, `ID_Produto` ou `ClienteID`. Assim, colunas como `CIDADE` ou `PEDIDO` não são confundidas com ID.

## Configuração

No topo do `verificar_pasta.py`:

| Constante | Para que serve |
|---|---|
| `NOMES_DE_EXIBICAO` | Nome de cada item no relatório (ex.: `"vdr": "Vidros"`). |
| `ACAO_CARDS_COM_ERRO` | Texto exibido depois de "N cards com erro". |
| `PALAVRAS_IGNORADAS` | Prefixos ou palavras que não ajudam a diferenciar os arquivos. |
| `LIMITE_PADRAO` | Semelhança mínima entre nomes para formar um par. |

## Estrutura

| Arquivo | Conteúdo |
|---|---|
| `verificar_pasta.py` | Script principal: lê a pasta, forma os pares e grava o relatório. |
| `comparar_excel.py` | Módulo que compara dois arquivos e devolve os códigos não encontrados. Pode ser usado sozinho: `from comparar_excel import verificar`. |
