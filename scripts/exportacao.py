"""
Montagem das abas e geração dos arquivos de saída (Excel e CSV) do tratamento.
Compartilhado pelo script de linha de comando e pelo app Streamlit.
"""
from __future__ import annotations

import io
from datetime import datetime

import pandas as pd

import tratar_pedidos_vendas as t
import validacoes as v

FORMATOS = {"receita": '"R$" #,##0.00', "data_pedido": "yyyy-mm-dd"}


def filtrar_validos(tratados: pd.DataFrame) -> pd.DataFrame:
    return tratados[tratados["status_registro"] != "duplicado_confirmado"].reset_index(drop=True)


def montar_metadados(bruto, tratados, validos, log, nome_arquivo, hash_origem) -> pd.DataFrame:
    ganho = validos["status"] == "Fechado Ganho"
    linhas = [
        ("arquivo_origem", nome_arquivo),
        ("sha256_origem", hash_origem),
        ("processado_em", datetime.now().strftime("%Y-%m-%d %H:%M:%S")),
        ("script", "tratar_pedidos_vendas.py"),
        ("linhas_origem", len(bruto)),
        ("linhas_validas", len(validos)),
        ("duplicados_confirmados", int((tratados["status_registro"] == "duplicado_confirmado").sum())),
        ("possiveis_duplicados", int((tratados["status_registro"] == "possivel_duplicado").sum())),
        ("receitas_ausentes", int(tratados["receita"].isna().sum())),
        ("celulas_alteradas", len(log)),
        ("receita_fechado_ganho_validos", float(validos.loc[ganho, "receita"].sum())),
    ]
    linhas += [(f"regra: {r}", n) for r, n in log["regra"].value_counts().items()]
    return pd.DataFrame(linhas, columns=["chave", "valor"])


def montar_abas(bruto, tratados, log, nome_arquivo, hash_origem) -> dict[str, pd.DataFrame]:
    validos = filtrar_validos(tratados)
    return {
        "pedidos_tratados": tratados,
        "pedidos_validos": validos,
        "log_alteracoes": log,
        "resumo_qualidade": v.resumo_qualidade(tratados, log),
        "validacoes": v.validar(bruto, tratados, log),
        "metadados": montar_metadados(bruto, tratados, validos, log, nome_arquivo, hash_origem),
    }


def formatar(ws, df: pd.DataFrame) -> None:
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    for i, col in enumerate(df.columns, start=1):
        letra = ws.cell(row=1, column=i).column_letter
        largura = max([len(str(col))] + [len(str(x)) for x in df[col].head(200)])
        ws.column_dimensions[letra].width = min(largura + 2, 50)
        formato = FORMATOS.get(col)
        if formato:
            for c in ws[letra][1:]:
                c.number_format = formato


def escrever_abas(writer: pd.ExcelWriter, abas: dict[str, pd.DataFrame]) -> None:
    for nome, dados in abas.items():
        dados.to_excel(writer, sheet_name=nome, index=False)
        formatar(writer.book[nome], dados)


def gerar_excel(abas: dict[str, pd.DataFrame], original: pd.DataFrame) -> bytes:
    """Planilha completa em memória; a aba 'original' recebe os dados exatamente como lidos."""
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as w:
        escrever_abas(w, abas)
        original.to_excel(w, sheet_name="original", index=False)
    return buffer.getvalue()


def gerar_csv(df: pd.DataFrame, separador: str = ";") -> bytes:
    """CSV em UTF-8 com BOM (abre direto no Excel). ';' usa vírgula decimal (padrão BR)."""
    saida = df.copy()
    if "data_pedido" in saida:
        saida["data_pedido"] = pd.to_datetime(saida["data_pedido"]).dt.strftime("%Y-%m-%d")
    decimal = "," if separador == ";" else "."
    return saida.to_csv(index=False, sep=separador, decimal=decimal).encode("utf-8-sig")
