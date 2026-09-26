"""
Leitura das planilhas de cada loja.

Cada loja exporta de um sistema diferente, então cada uma tem o seu
leitor. Todos devolvem o mesmo formato (uma linha por item vendido) e
registram tudo o que precisaram corrigir ou descartar — esse registro vai
pro relatório final como "qualidade dos dados", pra ninguém confiar num
número sem saber de onde ele veio.

LGPD: a planilha do Centro traz o nome do cliente. Numa farmácia de
manipulação, venda pode revelar dado de saúde (dado sensível pela LGPD);
o relatório não precisa saber quem comprou, então a coluna é descartada
logo na leitura e nunca chega às saídas.
"""

import csv
import glob
import io
import os
import re
import unicodedata
import zipfile
from dataclasses import dataclass, field
from datetime import date, datetime

import pandas as pd

COLUNAS = ["data", "loja", "venda", "categoria", "produto", "qtd", "valor"]

CATEGORIAS = ["Dermatológicos", "Cápsulas e suplementos", "Florais e homeopatia", "Veterinários", "Revenda"]
OUTROS = "Outros"


def _chave(texto) -> str:
    """'  Cápsulas E Suplementos ' -> 'capsulas e suplementos' (sem acento, minúsculo)."""
    sem_acento = unicodedata.normalize("NFKD", str(texto)).encode("ascii", "ignore").decode()
    return " ".join(sem_acento.casefold().split())


SINONIMOS = {
    **{_chave(c): c for c in CATEGORIAS},
    "dermato": "Dermatológicos",
    "suplemento": "Cápsulas e suplementos",
    "florais/homeo": "Florais e homeopatia",
    "vet": "Veterinários",
}


@dataclass
class Registro:
    """O que aconteceu na leitura de uma loja."""
    loja: str
    arquivo: str = ""
    linhas_lidas: int = 0
    linhas_validas: int = 0
    devolucoes: int = 0
    valor_devolucoes: float = 0.0
    descartes: dict = field(default_factory=dict)
    avisos: list = field(default_factory=list)

    def descartar(self, motivo: str, n: int):
        if n:
            self.descartes[motivo] = self.descartes.get(motivo, 0) + int(n)


def valor_texto(texto) -> float:
    """
    'R$ 1.234,56' / '1234,56' / '1,234.56' / 1234.56 -> 1234.56.
    Com os dois separadores, o último é o decimal. Vazio ou inválido -> NaN.
    """
    if isinstance(texto, (int, float)):
        return float(texto)
    t = re.sub(r"[R$\s]", "", str(texto))
    if not t:
        return float("nan")
    if "," in t and "." in t:
        t = t.replace(".", "").replace(",", ".") if t.rfind(",") > t.rfind(".") else t.replace(",", "")
    elif "," in t:
        t = t.replace(",", ".")
    try:
        return float(t)
    except ValueError:
        return float("nan")


def para_data(valor):
    if isinstance(valor, datetime):
        return valor.date()
    if isinstance(valor, date):
        return valor
    try:
        return datetime.strptime(str(valor).strip()[:10], "%d/%m/%Y").date()
    except ValueError:
        return None


def exigir(df: pd.DataFrame, colunas: list[str]):
    """Mensagem com o nome da coluna como está na planilha, não o nome interno."""
    faltando = [c for c in colunas if c not in df.columns]
    if faltando:
        raise ValueError(f"colunas não encontradas: {', '.join(faltando)} — o formato da planilha mudou")


def normalizar_categorias(df: pd.DataFrame, reg: Registro) -> pd.DataFrame:
    df["categoria"] = df["categoria"].map(lambda c: SINONIMOS.get(_chave(c), OUTROS))
    desconhecidas = int((df["categoria"] == OUTROS).sum())
    if desconhecidas:
        reg.avisos.append(f"linhas com categoria não reconhecida, somadas em \"{OUTROS}\": {desconhecidas}")
    return df


def finalizar(df: pd.DataFrame, reg: Registro, ano: int, mes: int) -> pd.DataFrame:
    """Validações comuns a todas as lojas."""
    df["data"] = df["data"].map(para_data)
    df["valor"] = df["valor"].map(valor_texto)
    df["qtd"] = pd.to_numeric(df["qtd"], errors="coerce")

    invalidas = df["data"].isna() | df["valor"].isna() | df["qtd"].isna()
    reg.descartar("linhas com data, valor ou quantidade ilegível", invalidas.sum())
    df = df[~invalidas].copy()

    fora = df["data"].map(lambda d: (d.year, d.month) != (ano, mes))
    reg.descartar("linhas com data fora do mês do relatório", fora.sum())
    df = df[~fora].copy()

    df["qtd"] = df["qtd"].astype(int)
    # Célula vazia vira texto explícito em vez de NaN (que quebraria o Excel
    # de saída e sumiria das contagens de vendas).
    for coluna, vazio in [("venda", "(sem número)"), ("produto", "(sem descrição)")]:
        df[coluna] = df[coluna].fillna("").astype(str).str.strip().replace("", vazio)
    df["loja"] = reg.loja

    devolucoes = df["valor"] < 0
    reg.devolucoes = int(devolucoes.sum())
    reg.valor_devolucoes = round(abs(float(df.loc[devolucoes, "valor"].sum())), 2)
    reg.linhas_validas = len(df)
    return df[COLUNAS]


def ler_centro(caminho: str, reg: Registro, ano: int, mes: int) -> pd.DataFrame:
    bruto = pd.read_excel(caminho, header=None, dtype=object)
    # Procura a linha de cabeçalho em vez de pular um número fixo de linhas:
    # se alguém acrescentar uma linha de título, a leitura continua certa.
    cabecalho = next((i for i, v in bruto[0].items() if _chave(v) == "data"), None)
    if cabecalho is None:
        raise ValueError("não encontrei a linha de cabeçalho (coluna 'Data')")
    df = bruto.iloc[cabecalho + 1:].copy()
    df.columns = [str(c).strip() for c in bruto.iloc[cabecalho]]
    df = df.dropna(how="all")
    reg.linhas_lidas = len(df)
    exigir(df, ["Data", "Nº Venda", "Categoria", "Descrição", "Qtd", "Valor Total"])

    total = df["Data"].astype(str).str.contains("TOTAL", case=False, na=False)
    reg.descartar("linhas de total da própria planilha", total.sum())
    df = df[~total]

    # LGPD (ver docstring do módulo). errors="ignore": se o sistema parar de
    # exportar essa coluna, ótimo — não é motivo pra rejeitar a planilha.
    df = df.drop(columns=["Cliente"], errors="ignore")
    df = df.rename(columns={"Data": "data", "Nº Venda": "venda", "Categoria": "categoria",
                            "Descrição": "produto", "Qtd": "qtd", "Valor Total": "valor"})
    return finalizar(normalizar_categorias(df, reg), reg, ano, mes)


def ler_umarizal(caminho: str, reg: Registro, ano: int, mes: int) -> pd.DataFrame:
    with open(caminho, "rb") as f:
        conteudo = f.read()
    try:
        conteudo.decode("utf-8")
        # utf-8-sig tira o marcador invisível (BOM) que o Excel põe no início
        # ao salvar "CSV UTF-8"; sem isso a 1ª coluna viraria "﻿DATA".
        codificacao = "utf-8-sig"
    except UnicodeDecodeError:
        # Sistemas brasileiros antigos exportam em Windows-1252; lido como
        # UTF-8, "CÁPSULAS" viraria lixo e a categoria não seria reconhecida.
        codificacao = "cp1252"
        reg.avisos.append("arquivo em codificação Windows (cp1252), convertido")
    # Leitura linha a linha em vez de pd.read_csv: uma linha com ";" a mais
    # (ex.: no nome do produto, sem aspas) é só pulada e contada. No pandas,
    # se essa linha for a primeira de dados, ele trata a 1ª coluna como
    # índice e desloca o arquivo inteiro.
    linhas = list(csv.reader(io.StringIO(conteudo.decode(codificacao)), delimiter=";"))
    if not linhas:
        raise ValueError("arquivo vazio")
    cabecalho, corpo = [c.strip() for c in linhas[0]], [l for l in linhas[1:] if any(c.strip() for c in l)]
    boas = [l for l in corpo if len(l) == len(cabecalho)]
    df = pd.DataFrame(boas, columns=cabecalho, dtype=object)
    reg.linhas_lidas = len(corpo)
    reg.descartar("linhas com número errado de colunas (\";\" dentro do texto, sem aspas)", len(corpo) - len(boas))
    exigir(df, ["DATA", "CUPOM", "CATEGORIA", "PRODUTO", "QUANTIDADE", "VALOR"])
    df = df.rename(columns={"DATA": "data", "CUPOM": "venda", "CATEGORIA": "categoria",
                            "PRODUTO": "produto", "QUANTIDADE": "qtd", "VALOR": "valor"})
    return finalizar(normalizar_categorias(df, reg), reg, ano, mes)


def ler_ananindeua(caminho: str, reg: Registro, ano: int, mes: int) -> pd.DataFrame:
    df = pd.read_excel(caminho, dtype=object).dropna(how="all")
    reg.linhas_lidas = len(df)
    exigir(df, ["dt_venda", "num_pedido", "tipo_produto", "item", "qtde", "total (R$)"])
    # Linha idêntica em tudo (pedido, item, quantidade e valor) é a mesma
    # venda exportada duas vezes; o sistema dessa loja lança quantidade
    # repetida como qtde 2, não como duas linhas iguais.
    duplicadas = df.duplicated()
    reg.descartar("linhas duplicadas (mesma venda exportada duas vezes)", duplicadas.sum())
    df = df[~duplicadas]
    df = df.rename(columns={"dt_venda": "data", "num_pedido": "venda", "tipo_produto": "categoria",
                            "item": "produto", "qtde": "qtd", "total (R$)": "valor"})
    return finalizar(normalizar_categorias(df, reg), reg, ano, mes)


LOJAS = [
    ("Centro", "vendas_centro_*.xlsx", ler_centro),
    ("Umarizal", "umarizal_vendas_*.csv", ler_umarizal),
    ("Ananindeua", "Vendas Ananindeua *.xlsx", ler_ananindeua),
]


def ler_mes(pasta: str, ano: int, mes: int):
    """Lê as planilhas de todas as lojas de um mês. Devolve (dados, registros)."""
    partes, registros = [], []
    for loja, padrao, leitor in LOJAS:
        reg = Registro(loja)
        arquivos = sorted(glob.glob(os.path.join(pasta, padrao)))
        if not arquivos:
            reg.avisos.append("PLANILHA NÃO ENVIADA — o total consolidado não inclui esta loja")
        elif len(arquivos) > 1:
            reg.avisos.append(f"{len(arquivos)} arquivos encontrados; nenhum foi usado pra não somar em dobro")
        else:
            reg.arquivo = os.path.basename(arquivos[0])
            try:
                partes.append(leitor(arquivos[0], reg, ano, mes))
            except KeyError as erro:
                reg = Registro(loja, arquivo=reg.arquivo)
                reg.avisos.append(f"coluna {erro} não encontrada — o formato da planilha mudou; loja fora do total")
            except (ValueError, OSError, zipfile.BadZipFile) as erro:
                reg = Registro(loja, arquivo=reg.arquivo)
                motivo = str(erro)
                if arquivos[0].endswith(".xlsx") and (isinstance(erro, zipfile.BadZipFile) or "format" in motivo):
                    motivo = "não é um arquivo Excel válido — pode estar corrompido"
                reg.avisos.append(f"arquivo não pôde ser lido ({motivo}) — loja fora do total")
        registros.append(reg)
    dados = pd.concat(partes, ignore_index=True) if partes else pd.DataFrame(columns=COLUNAS)
    return dados, registros
