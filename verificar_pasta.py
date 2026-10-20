"""
Verifica todos os arquivos Excel de uma pasta, comparando os pares de nomes semelhantes.

Passo a passo:
  1. Lista os arquivos Excel da pasta (não entra em subpastas).
  2. Separa em Arquivo 1 (sem "ID" no nome) e Arquivo 2 (com "ID" no nome).
  3. Monta a tabela de pares pelas partes do nome (ver montar_pares).
     Um Arquivo 1 pode ter vários pares, e um Arquivo 2 também.
  4. Para cada par, compara a coluna "codigo" do Arquivo 1 com cada coluna ID do
     Arquivo 2 e grava uma linha curta no relatório TXT:

        CRUZAMENTO 20/10/2026

        1 - 2.1: 1 card com erro - Realizar Conferência Manual
        1 - 2.2: sem erro
        5 - 1.1: coluna de ID vazia
        Entregas Especiais - Vidros: sem erro

Uso no terminal:
    python verificar_pasta.py "C:\\caminho\\da\\pasta"
    python verificar_pasta.py pasta --so-pares          (só mostra os pares, não compara)
    python verificar_pasta.py pasta --limite 70         (exige 70% de cobertura)
    python verificar_pasta.py pasta --saida relatorio.txt --acrescentar

Ou a partir de outro código:
    from verificar_pasta import verificar_pasta
    pares = verificar_pasta("pasta")

Requer o comparar_excel.py na mesma pasta deste script.
"""
from __future__ import annotations

import argparse
import math
import re
import unicodedata
from datetime import date
from difflib import SequenceMatcher
from pathlib import Path

import pandas as pd

from comparar_excel import PADRAO_ID, ColunaNaoEncontrada, tem_id, verificar

EXTENSOES = {".xlsx", ".xlsm", ".xls"}
NOME_RELATORIO = "relatorio_nao_encontrados.txt"

LIMITE_PADRAO = 0.60

# Palavras que aparecem em todos os nomes e não ajudam a diferenciar.
# Se outras pastas usarem outro prefixo, acrescente aqui (ou use --ignorar).
PALAVRAS_IGNORADAS = {"procv"}

# Conectores que não contam na comparação dos nomes.
CONECTORES = {"e", "de", "da", "do", "das", "dos", "a", "o"}
# "para" não conta na cobertura, mas é usado no desempate pela ordem
# ("entrega para vidro" x "vidro para entrega") e separa os itens do rótulo.
PALAVRA_DE_DIRECAO = "para"
# "e" separa itens de uma lista no nome ("7.1_e_8").
PALAVRA_DE_LISTA = "e"

# Como cada item aparece no relatório. A chave é o item como sai do nome do
# arquivo (minúsculo, sem acento); o que não estiver aqui aparece com a
# primeira letra maiúscula ("vendas" -> "Vendas").
NOMES_DE_EXIBICAO = {
    "entrega": "Entregas Especiais",
    "entregas": "Entregas Especiais",
    "vidro": "Vidros",
    "vdr": "Vidros",
    "marcenaria": "Marcenaria",
    "marcena": "Marcenaria",
    "solicitao de material": "Solicitação de Material",
    "solicitado material": "Solicitação de Material",
}

# Orientação que vai depois de "N cards com erro" no relatório.
# Deixe "" para não mostrar nada.
ACAO_CARDS_COM_ERRO = "Realizar Conferência Manual"

# Datas do download no nome (20-10-2026 ou 2026-10-20) e o " (1)" que o Windows
# acrescenta quando o mesmo arquivo é baixado de novo.
PADRAO_DATA = re.compile(r"\d{1,2}-\d{1,2}-\d{2,4}|\d{4}-\d{1,2}-\d{1,2}")
PADRAO_COPIA = re.compile(r"\s*\(\d+\)$")
PADRAO_PARTE = re.compile(r"\d+(?:\.\d+)*|[a-z]+")

# Palavras com pelo menos 4 letras e 80% de semelhança contam como a mesma
# ("entregas" = "entrega", "solicitado" = "solicitao").
TAMANHO_MINIMO_APROXIMAR = 4
SEMELHANCA_PALAVRAS = 0.80


# --------------------------------------------------------------------------
# 1 e 2. Listar e separar os arquivos
# --------------------------------------------------------------------------

def listar_excel(diretorio) -> list[Path]:
    """Arquivos Excel da pasta, ignorando os temporários do Excel (~$...)."""
    return sorted(
        p for p in Path(diretorio).iterdir()
        if p.is_file() and p.suffix.lower() in EXTENSOES and not p.name.startswith("~$")
    )


def separar_arquivos(arquivos: list[Path]) -> tuple[list[Path], list[Path]]:
    """Separa em (Arquivos 1, Arquivos 2). Arquivo 2 é o que tem "ID" no nome."""
    arquivos_1 = [a for a in arquivos if not tem_id(a.stem)]
    arquivos_2 = [a for a in arquivos if tem_id(a.stem)]
    return arquivos_1, arquivos_2


# --------------------------------------------------------------------------
# 3. Comparação dos nomes e tabela de pares
# --------------------------------------------------------------------------

def texto_do_nome(arquivo: Path) -> str:
    """Nome sem extensão, data, " (1)", "id" e acentos, em minúsculas.
    Os separadores (_ e espaços) são mantidos."""
    texto = PADRAO_COPIA.sub("", arquivo.stem)
    texto = PADRAO_DATA.sub(" ", texto)
    texto = PADRAO_ID.sub(" ", texto)
    return unicodedata.normalize("NFKD", texto).encode("ascii", "ignore").decode().lower()


def partes_do_nome(arquivo: Path, ignorar=PALAVRAS_IGNORADAS, com_direcao: bool = False) -> list[str]:
    """Quebra o nome do arquivo nas partes que identificam a tabela.

    "procv_7.1_e_8____solicitao_de_material_id_20-10-2026.xlsx"
        -> ["7.1", "8", "solicitao", "material"]

    Tira conectores e palavras ignoradas. Números como "2.1" ficam inteiros.
    Com com_direcao=True, mantém o "para".
    """
    descartar = set(ignorar) | CONECTORES
    if not com_direcao:
        descartar.add(PALAVRA_DE_DIRECAO)
    return [p for p in PADRAO_PARTE.findall(texto_do_nome(arquivo)) if p not in descartar]


def palavras_equivalentes(a: str, b: str) -> bool:
    """Iguais, ou palavras parecidas o bastante ("entregas" e "entrega").
    Números nunca são aproximados: "2.1" e "2.2" são partes diferentes."""
    if a == b:
        return True
    if not (a.isalpha() and b.isalpha()):
        return False
    if min(len(a), len(b)) < TAMANHO_MINIMO_APROXIMAR:
        return False
    return SequenceMatcher(None, a, b).ratio() >= SEMELHANCA_PALAVRAS


def alinhar_partes(partes_1: list[str], partes_2: list[str]) -> list[str]:
    """Troca cada parte do nome 2 pela equivalente do nome 1, quando houver,
    para que palavras parecidas contem como iguais nas comparações seguintes."""
    referencia = list(dict.fromkeys(partes_1))
    return [
        p if p in referencia else next((r for r in referencia if palavras_equivalentes(r, p)), p)
        for p in partes_2
    ]


def cobertura(arquivo_1: Path, arquivo_2: Path, ignorar=PALAVRAS_IGNORADAS) -> float:
    """Quanto das partes do nome menor aparece no outro, de 0 a 1.

    "2.2 3 4" x "2.2 1.1 3": o menor tem 3 partes e 2 aparecem no outro = 0.67
    "4 5"     x "2.2 4":     o menor tem 2 partes e 1 aparece no outro  = 0.50
    """
    partes_1 = partes_do_nome(arquivo_1, ignorar)
    partes_2 = alinhar_partes(partes_1, partes_do_nome(arquivo_2, ignorar))
    conjunto_1, conjunto_2 = set(partes_1), set(partes_2)
    if not conjunto_1 or not conjunto_2:
        return 0.0
    return len(conjunto_1 & conjunto_2) / min(len(conjunto_1), len(conjunto_2))


def semelhanca_de_ordem(arquivo_1: Path, arquivo_2: Path, ignorar=PALAVRAS_IGNORADAS) -> float:
    """Quanto as partes aparecem na mesma sequência, de 0 a 1 (usado no desempate).
    "entrega entrega para vidro" fica mais perto de "vidro entrega para vidro"
    do que "vidro vidro para entrega"."""
    partes_1 = partes_do_nome(arquivo_1, ignorar, com_direcao=True)
    partes_2 = alinhar_partes(partes_1, partes_do_nome(arquivo_2, ignorar, com_direcao=True))
    return SequenceMatcher(None, partes_1, partes_2).ratio()


# ---- Rótulo curto do par ("1 - 2.1", "Entregas Especiais - Vidros") ----

def itens_do_nome(arquivo: Path, ignorar=PALAVRAS_IGNORADAS) -> list[str]:
    """Itens do nome na ordem em que aparecem: números e expressões.

    "procv_8___solicitao_de_material_20-10-2026"  -> ["8", "solicitao de material"]
    "procv_entrega___entrega_para_vidro_20-10-2026" -> ["entrega", "entrega", "vidro"]

    Palavras ligadas por um único "_" formam uma expressão; "para", "e",
    separadores longos ("___") e números iniciam um item novo.
    """
    texto = texto_do_nome(arquivo)
    quebras = set(ignorar) | {PALAVRA_DE_DIRECAO, PALAVRA_DE_LISTA}
    itens, expressao, fim_anterior = [], [], 0

    def fechar_expressao():
        while expressao and expressao[-1] in CONECTORES:
            expressao.pop()
        if expressao:
            itens.append(" ".join(expressao))
        expressao.clear()

    for achado in PADRAO_PARTE.finditer(texto):
        parte, separador = achado.group(), texto[fim_anterior:achado.start()]
        fim_anterior = achado.end()
        if parte in quebras:
            fechar_expressao()
        elif parte[0].isdigit():
            fechar_expressao()
            itens.append(parte)
        else:
            if expressao and len(separador) > 1:
                fechar_expressao()
            if expressao or parte not in CONECTORES:
                expressao.append(parte)
    fechar_expressao()
    return itens


def itens_equivalentes(x: str, y: str) -> bool:
    """Números precisam ser iguais; expressões, ter as mesmas palavras (aproximadas)."""
    if x[0].isdigit() or y[0].isdigit():
        return x == y
    palavras_x = [p for p in x.split() if p not in CONECTORES]
    palavras_y = [p for p in y.split() if p not in CONECTORES]
    return len(palavras_x) == len(palavras_y) and all(
        palavras_equivalentes(a, b) for a, b in zip(palavras_x, palavras_y)
    )


def nome_de_exibicao(item: str) -> str:
    return NOMES_DE_EXIBICAO.get(item, item[:1].upper() + item[1:])


def rotulo_do_par(arquivo_1: Path, arquivo_2: Path, ignorar=PALAVRAS_IGNORADAS) -> str:
    """Rótulo curto da relação comparada: os itens que os dois nomes têm em comum,
    na ordem do Arquivo 1.

    "1___2.1_e_1___2.2" x "1___2.1_id"          -> "1 - 2.1"
    "2.2___3_e_4"       x "2.2_e_1.1_____3_id"  -> "2.2 - 3"
    """
    itens_1 = list(dict.fromkeys(itens_do_nome(arquivo_1, ignorar)))
    itens_2 = itens_do_nome(arquivo_2, ignorar)
    comuns = [i for i in itens_1 if any(itens_equivalentes(i, j) for j in itens_2)] or itens_1
    return " - ".join(nome_de_exibicao(i) for i in comuns)


def chave_de_ordenacao(rotulo: str) -> list:
    """Ordena números como números ("2.2" antes de "10") e depois os nomes."""
    chave = []
    for item in rotulo.split(" - "):
        if re.fullmatch(r"\d+(?:\.\d+)*", item):
            chave.append((0, tuple(int(n) for n in item.split(".")), ""))
        else:
            chave.append((1, (), item.lower()))
    return chave


def montar_pares(
    arquivos_1: list[Path],
    arquivos_2: list[Path],
    limite: float = LIMITE_PADRAO,
    ignorar=PALAVRAS_IGNORADAS,
) -> pd.DataFrame:
    """Tabela com um par por linha: rotulo, arquivo_1, arquivo_2, nota, ordem.

    Para cada Arquivo 2:
      - entram os Arquivos 1 com cobertura >= limite;
      - se vários empatam na maior cobertura, ficam só os de melhor ordem das
        palavras (os demais candidatos, com cobertura menor, continuam).
    """
    linhas = []
    for a2 in arquivos_2:
        candidatos = []
        for a1 in arquivos_1:
            nota = cobertura(a1, a2, ignorar)
            if nota >= limite:
                candidatos.append((a1, nota, semelhanca_de_ordem(a1, a2, ignorar)))
        if not candidatos:
            continue

        maior_nota = max(nota for _, nota, _ in candidatos)
        melhor_ordem = max(ordem for _, nota, ordem in candidatos if math.isclose(nota, maior_nota))
        for a1, nota, ordem in candidatos:
            empatado_no_topo = math.isclose(nota, maior_nota)
            if not empatado_no_topo or math.isclose(ordem, melhor_ordem):
                linhas.append({
                    "rotulo": rotulo_do_par(a1, a2, ignorar),
                    "arquivo_1": a1, "arquivo_2": a2,
                    "nota": round(nota, 2), "ordem": round(ordem, 2),
                })

    linhas.sort(key=lambda l: (chave_de_ordenacao(l["rotulo"]), l["arquivo_1"].name, l["arquivo_2"].name))
    return pd.DataFrame(linhas, columns=["rotulo", "arquivo_1", "arquivo_2", "nota", "ordem"])


def mostrar_pares(pares: pd.DataFrame, arquivos_1: list[Path], arquivos_2: list[Path], limite: float) -> None:
    """Imprime a tabela de pares e os arquivos que ficaram sem par."""
    print(f"Pares encontrados (cobertura mínima {limite:.0%}): {len(pares)}")
    for par in pares.itertuples(index=False):
        print(f"  {par.rotulo}   [{par.nota:.0%}]\n      {par.arquivo_1.name}\n      {par.arquivo_2.name}")

    com_par_1, com_par_2 = set(pares["arquivo_1"]), set(pares["arquivo_2"])
    sem_par_1 = [a.name for a in arquivos_1 if a not in com_par_1]
    sem_par_2 = [a.name for a in arquivos_2 if a not in com_par_2]
    if sem_par_1:
        print("Arquivos 1 sem par:\n  " + "\n  ".join(sem_par_1))
    if sem_par_2:
        print("Arquivos 2 (com ID) sem par:\n  " + "\n  ".join(sem_par_2))


# --------------------------------------------------------------------------
# 4. Comparações e relatório
# --------------------------------------------------------------------------

def descrever_status(resultado) -> str:
    """Status curto de uma comparação, no estilo do relatório manual."""
    if resultado.total_verificado == 0:
        return f"coluna {resultado.coluna_a} vazia"
    if resultado.coluna_b_vazia:
        return "coluna de ID vazia"
    if resultado.quantidade == 0:
        return "sem erro"
    cards = f"{resultado.quantidade} card{'s' if resultado.quantidade > 1 else ''} com erro"
    return f"{cards} - {ACAO_CARDS_COM_ERRO}" if ACAO_CARDS_COM_ERRO else cards


def linhas_do_par(par) -> list[str]:
    """Compara um par e devolve suas linhas do relatório. Se o Arquivo 2 tiver
    mais de uma coluna ID, cada uma vira uma linha com o nome da coluna."""
    try:
        resultados = verificar(par.arquivo_1, par.arquivo_2)
    except ColunaNaoEncontrada as erro:
        print(f"  Aviso ({par.rotulo}): {erro}")
        return [f"{par.rotulo}: {erro.resumo}"]
    except Exception as erro:
        print(f"  Erro ({par.rotulo}): {erro}")
        return [f"{par.rotulo}: erro ao abrir os arquivos"]

    varias_colunas = len(resultados) > 1
    return [
        f"{par.rotulo}{f' ({r.coluna_b})' if varias_colunas else ''}: {descrever_status(r)}"
        for r in resultados
    ]


def acrescentar_linha(caminho: Path, linha: str) -> None:
    """Acrescenta uma linha ao fim do TXT (UTF-8)."""
    with open(caminho, "a", encoding="utf-8") as arquivo:
        arquivo.write(linha + "\n")


def executar_comparacoes(pares: pd.DataFrame, saida: Path) -> list[str]:
    """Grava o cabeçalho e, para cada par, as linhas do relatório, uma a uma.
    Um par com problema vira uma linha de aviso e os demais continuam.
    Devolve as linhas gravadas (sem o cabeçalho)."""
    if saida.exists() and saida.stat().st_size > 0:
        acrescentar_linha(saida, "")
    acrescentar_linha(saida, f"CRUZAMENTO {date.today():%d/%m/%Y}")
    acrescentar_linha(saida, "")

    gravadas = []
    for par in pares.itertuples(index=False):
        for linha in linhas_do_par(par):
            acrescentar_linha(saida, linha)
            gravadas.append(linha)
    return gravadas


def verificar_pasta(
    diretorio,
    saida=None,
    limite: float = LIMITE_PADRAO,
    ignorar=PALAVRAS_IGNORADAS,
    acrescentar: bool = False,
    so_pares: bool = False,
) -> pd.DataFrame:
    """Executa o fluxo completo e devolve a tabela de pares.

    saida:        caminho do TXT (padrão: relatorio_nao_encontrados.txt dentro da pasta)
    limite:       cobertura mínima entre os nomes, de 0 a 1 (padrão 0.6)
    ignorar:      palavras desconsideradas nos nomes (padrão: {"procv"})
    acrescentar:  True mantém o conteúdo anterior do TXT; False começa um TXT novo
    so_pares:     True só monta e mostra os pares, sem comparar nem gravar
    """
    diretorio = Path(diretorio)
    saida = Path(saida) if saida else diretorio / NOME_RELATORIO

    arquivos_1, arquivos_2 = separar_arquivos(listar_excel(diretorio))
    pares = montar_pares(arquivos_1, arquivos_2, limite, ignorar)
    mostrar_pares(pares, arquivos_1, arquivos_2, limite)

    if so_pares:
        return pares

    if not acrescentar:
        saida.write_text("", encoding="utf-8")
    print()
    linhas = executar_comparacoes(pares, saida)
    print("\n" + "\n".join(linhas))
    print(f"\n{len(linhas)} linhas gravadas em {saida}")
    return pares


def ler_argumentos():
    parser = argparse.ArgumentParser(description="Compara os pares de arquivos Excel de uma pasta.")
    parser.add_argument("diretorio", help="pasta com os arquivos Excel")
    parser.add_argument("--saida", help=f"arquivo TXT de saída (padrão: {NOME_RELATORIO} dentro da pasta)")
    parser.add_argument("--limite", type=float, default=LIMITE_PADRAO,
                        help="cobertura mínima entre nomes, ex.: 0.6 ou 60 (padrão: 60%%)")
    parser.add_argument("--ignorar", default=",".join(sorted(PALAVRAS_IGNORADAS)),
                        help='palavras ignoradas nos nomes, separadas por vírgula (padrão: "procv")')
    parser.add_argument("--acrescentar", action="store_true",
                        help="mantém o que já existe no TXT em vez de começar um novo")
    parser.add_argument("--so-pares", action="store_true",
                        help="só mostra os pares encontrados, sem comparar")
    return parser.parse_args()


if __name__ == "__main__":
    args = ler_argumentos()
    verificar_pasta(
        args.diretorio,
        saida=args.saida,
        limite=args.limite / 100 if args.limite > 1 else args.limite,
        ignorar={p.strip().lower() for p in args.ignorar.split(",") if p.strip()},
        acrescentar=args.acrescentar,
        so_pares=args.so_pares,
    )
