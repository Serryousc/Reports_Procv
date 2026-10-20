"""
Módulo que verifica quais códigos de um arquivo Excel não existem em outro.

Padrão: compara a coluna "codigo" do Arquivo 1 com CADA coluna do Arquivo 2
que tenha "ID" no nome, e gera uma linha de texto por comparação.

Requisitos: pip install pandas openpyxl   (arquivos .xls antigos precisam também de xlrd)

Uso a partir de outro código:

    from comparar_excel import verificar, gerar_linha

    for resultado in verificar("vendas_jan.xlsx", "vendas_ID.xlsx"):
        print(gerar_linha(resultado))

    # vendas_jan.xlsx x vendas_ID.xlsx (ID_Produto): 3 não encontrados
    # vendas_jan.xlsx x vendas_ID.xlsx (ID_Antigo): nenhum código faltando

    # Para listar os códigos na linha: gerar_linha(resultado, com_codigos=True)
    # A lista completa também fica em resultado.nao_encontrados.

verificar() devolve uma lista com um resultado por coluna ID do Arquivo 2.
Cada resultado tem: nao_encontrados, quantidade, total_verificado,
coluna_a, coluna_b, arquivo_a, arquivo_b.
"""
from __future__ import annotations

import re
import unicodedata
from dataclasses import dataclass
from functools import lru_cache
from pathlib import Path

import pandas as pd

COLUNA_A_PADRAO = "codigo"

# "ID" como palavra: pega "ID", "ID_Produto", "Cliente ID", "id_cliente",
# "IDCliente" e "ClienteID", mas não "CIDADE", "PEDIDO", "Validade" nem "IDADE".
PADRAO_ID = re.compile(
    r"(?<![A-Za-z])[Ii][Dd](?![a-z])(?![A-Z]{2})"  # ID no início de palavra
    r"|(?<=[a-z])I[Dd](?![a-z])"                   # ID no fim de palavra (ClienteID)
)

__all__ = [
    "verificar", "gerar_linha", "tem_id",
    "ResultadoComparacao", "ColunaNaoEncontrada", "PADRAO_ID",
]


class ColunaNaoEncontrada(ValueError):
    """Coluna esperada não existe no arquivo.
    A mensagem traz os detalhes; `resumo` é a versão curta para relatórios."""

    def __init__(self, mensagem: str, resumo: str):
        super().__init__(mensagem)
        self.resumo = resumo


@dataclass
class ResultadoComparacao:
    arquivo_a: str
    arquivo_b: str
    coluna_a: str
    coluna_b: str
    total_verificado: int        # códigos distintos e não vazios no Arquivo 1
    nao_encontrados: list[str]   # códigos do Arquivo 1 ausentes nessa coluna do Arquivo 2
    coluna_b_vazia: bool = False  # a coluna ID do Arquivo 2 não tem nenhum valor

    @property
    def quantidade(self) -> int:
        return len(self.nao_encontrados)


def tem_id(texto) -> bool:
    """Diz se o texto (nome de coluna ou de arquivo) contém "ID" como palavra."""
    return bool(PADRAO_ID.search(str(texto)))


# --------------------------------------------------------------------------
# Leitura e localização das colunas
# --------------------------------------------------------------------------

@lru_cache(maxsize=32)
def _ler_excel(caminho: str, aba, modificado_em: float) -> pd.DataFrame:
    return pd.read_excel(caminho, sheet_name=aba, dtype=str)


def carregar_planilha(caminho, aba=0) -> pd.DataFrame:
    """Lê a planilha com todas as células como texto, exatamente como estão.
    A padronização (números, espaços etc.) é feita depois, em normalizar_valores.

    Guarda o resultado em memória: quando o mesmo arquivo entra em várias
    comparações, ele é lido do disco uma vez só (e de novo se for alterado).
    """
    caminho = Path(caminho)
    return _ler_excel(str(caminho.resolve()), aba, caminho.stat().st_mtime)


def simplificar_nome(texto) -> str:
    """Remove acentos e espaços nas pontas e passa para minúsculas (" Código " -> "codigo")."""
    sem_acento = unicodedata.normalize("NFKD", str(texto)).encode("ascii", "ignore").decode()
    return sem_acento.strip().lower()


def encontrar_coluna_por_nome(df: pd.DataFrame, nome: str) -> str:
    """Encontra a coluna com esse nome, sem diferenciar maiúsculas, acentos ou espaços."""
    alvo = simplificar_nome(nome)
    for coluna in df.columns:
        if simplificar_nome(coluna) == alvo:
            return coluna
    raise ValueError(f'coluna "{nome}" não encontrada. Colunas disponíveis: {list(df.columns)}')


def encontrar_colunas_id(df: pd.DataFrame) -> list[str]:
    """Devolve todas as colunas que têm "ID" como palavra no nome."""
    colunas = [c for c in df.columns if tem_id(c)]
    if not colunas:
        raise ValueError(f'nenhuma coluna com "ID" no nome. Colunas disponíveis: {list(df.columns)}')
    return colunas


# --------------------------------------------------------------------------
# Comparação
# --------------------------------------------------------------------------

# Caracteres invisíveis que às vezes vêm em exportações de sistemas
CARACTERES_INVISIVEIS = r"[​‌‍⁠﻿]"


def normalizar_valores(serie: pd.Series, como_numero: bool = True) -> pd.Series:
    """Padroniza os valores para a comparação não falhar por detalhes de formato.

    Sempre: tira espaços nas pontas (inclusive o espaço "duro" do Excel),
    caracteres invisíveis e diferença de maiúsculas/minúsculas.

    Com como_numero=True (padrão), faz o mesmo que o "Texto para colunas" do
    Excel faz com códigos numéricos: "00123", "123", 123 e 123.0 viram todos
    "123". Códigos com letras ("00ABC") ficam como estão.

    Células vazias viram NA e são ignoradas.
    """
    s = (
        serie.astype("string")
        .str.replace(CARACTERES_INVISIVEIS, "", regex=True)
        .str.strip()
        .str.upper()
    )
    if como_numero:
        numericos = s.str.fullmatch(r"\d+(?:\.0)?").fillna(False).astype(bool)
        como_inteiro = s.str.replace(r"\.0$", "", regex=True).str.lstrip("0").replace("", "0")
        s = s.mask(numericos, como_inteiro)
    else:
        s = s.str.replace(r"\.0$", "", regex=True)
    return s.replace("", pd.NA)


def verificar(
    arquivo_a,
    arquivo_b,
    coluna_a: str = COLUNA_A_PADRAO,
    coluna_b: str | None = None,
    aba_a=0,
    aba_b=0,
    como_numero: bool = True,
) -> list[ResultadoComparacao]:
    """Verifica quais códigos do Arquivo 1 (arquivo_a) não existem no Arquivo 2 (arquivo_b).

    Compara com cada coluna do Arquivo 2 que tenha "ID" no nome, e devolve um
    resultado por coluna. Para usar uma coluna específica, informe coluna_b.
    Códigos repetidos aparecem uma vez só e células vazias são ignoradas.
    Erros de coluna trazem o nome do arquivo na mensagem.

    como_numero=True trata códigos numéricos como no "Texto para colunas" do
    Excel ("00123" = 123). Use False se os zeros à esquerda forem importantes.
    """
    nome_a, nome_b = Path(arquivo_a).name, Path(arquivo_b).name
    df_a = carregar_planilha(arquivo_a, aba_a)
    df_b = carregar_planilha(arquivo_b, aba_b)

    try:
        col_a = encontrar_coluna_por_nome(df_a, coluna_a)
    except ValueError as erro:
        raise ColunaNaoEncontrada(f"{nome_a}: {erro}", f"sem coluna {coluna_a}") from None

    try:
        colunas_b = [encontrar_coluna_por_nome(df_b, coluna_b)] if coluna_b else encontrar_colunas_id(df_b)
    except ValueError as erro:
        resumo = f"sem coluna {coluna_b}" if coluna_b else "sem coluna de ID"
        raise ColunaNaoEncontrada(f"{nome_b}: {erro}", resumo) from None

    chaves_a = normalizar_valores(df_a[col_a], como_numero)
    distintos = chaves_a.notna() & ~chaves_a.duplicated()
    chaves_distintas = chaves_a[distintos]
    codigos_originais = df_a.loc[distintos, col_a].str.strip()

    resultados = []
    for col_b in colunas_b:
        chaves_b = set(normalizar_valores(df_b[col_b], como_numero).dropna())
        faltando = ~chaves_distintas.isin(chaves_b)
        resultados.append(
            ResultadoComparacao(
                arquivo_a=str(arquivo_a),
                arquivo_b=str(arquivo_b),
                coluna_a=str(col_a),
                coluna_b=str(col_b),
                total_verificado=len(chaves_distintas),
                nao_encontrados=codigos_originais[faltando].tolist(),
                coluna_b_vazia=not chaves_b,
            )
        )
    return resultados


# --------------------------------------------------------------------------
# Saída em texto
# --------------------------------------------------------------------------

def gerar_linha(resultado: ResultadoComparacao, com_codigos: bool = False) -> str:
    """Monta a linha do relatório para uma comparação. Exemplos:
    vendas_jan.xlsx x vendas_ID.xlsx (ID_Produto): 3 não encontrados
    vendas_fev.xlsx x vendas_ID.xlsx (ID_Produto): nenhum código faltando

    Com com_codigos=True, lista também os códigos no fim da linha:
    vendas_jan.xlsx x vendas_ID.xlsx (ID_Produto): 3 não encontrados - XYZ, 789, novo-1
    """
    inicio = f"{Path(resultado.arquivo_a).name} x {Path(resultado.arquivo_b).name} ({resultado.coluna_b})"
    if resultado.coluna_b_vazia:
        return f"{inicio}: coluna de ID vazia"
    if not resultado.nao_encontrados:
        return f"{inicio}: nenhum código faltando"
    rotulo = "não encontrado" if resultado.quantidade == 1 else "não encontrados"
    linha = f"{inicio}: {resultado.quantidade} {rotulo}"
    if com_codigos:
        codigos = ", ".join(" ".join(c.split()) for c in resultado.nao_encontrados)  # garante uma linha só
        linha += f" - {codigos}"
    return linha
