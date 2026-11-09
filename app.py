"""
App Streamlit: limpeza automática de pedidos de vendas.

O usuário envia o CSV, o app aplica o tratamento (scripts/tratar_pedidos_vendas.py),
roda as validações (scripts/validacoes.py), mostra erros e melhorias e oferece o
download da base limpa em CSV ou Excel.

Uso:
    pip install -r requirements.txt
    streamlit run app.py
"""
from __future__ import annotations

import hashlib
import sys
from pathlib import Path

import altair as alt
import pandas as pd
import streamlit as st

RAIZ = Path(__file__).resolve().parent
sys.path.insert(0, str(RAIZ / "scripts"))

import exportacao as ex  # noqa: E402
import tratar_pedidos_vendas as t  # noqa: E402
import validacoes as v  # noqa: E402

EXEMPLO = RAIZ / "data" / "exemplo_pedidos_vendas.csv"
COR_SERIE = "#2a78d6"
DESTAQUE = "background-color: rgba(42, 120, 214, 0.20)"
LIMITE_DESTAQUE = 5000          # acima disso a tabela é exibida sem realce de células

DESCRICAO_COLUNAS = {
    "ID Pedido": "identificador numérico do pedido",
    "Nome do Cliente": "nome do cliente",
    "Região": "Norte, Nordeste, Centro-Oeste, Sudeste ou Sul",
    "Representante": "vendedor responsável",
    "Data do Pedido": "dd/mm/aaaa, aaaa-mm-dd ou '5 Jan 2024'",
    "Receita": "valor do pedido, ex.: R$ 4.320,00",
    "Produto": "produto vendido",
    "Status": "Fechado Ganho / Fechado Perdido",
    "Observações": "texto livre (opcional)",
}


# --------------------------------------------------------------------------- processamento
@st.cache_data(show_spinner="Limpando e validando os dados...")
def processar(conteudo: bytes, nome: str) -> dict:
    bruto = t.ler_arquivo(conteudo, nome)
    tratados, log = t.tratar(bruto, nome)
    abas = ex.montar_abas(bruto, tratados, log, nome, hashlib.sha256(conteudo).hexdigest())
    return {
        "bruto": bruto,
        "abas": abas,
        "indicadores": v.indicadores(tratados, log),
        "excel": ex.gerar_excel(abas, bruto),
    }


def moeda(x) -> str:
    if pd.isna(x):
        return ""
    return "R$ " + f"{x:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def data_iso(x) -> str:
    return "" if pd.isna(x) else pd.Timestamp(x).strftime("%Y-%m-%d")


def tabela_destacada(df: pd.DataFrame, log: pd.DataFrame):
    """Realça as células alteradas pela limpeza (cruzando com o log por linha_origem)."""
    formatos = {"receita": moeda, "data_pedido": data_iso}
    if len(df) > LIMITE_DESTAQUE:
        return df
    posicao = {linha: i for i, linha in enumerate(df["linha_origem"])}
    alteradas = {(posicao[l], c) for l, c in zip(log["linha_origem"], log["coluna"])
                 if l in posicao and c in df.columns}

    def estilo(_):
        estilos = pd.DataFrame("", index=df.index, columns=df.columns)
        for i, c in alteradas:
            estilos.iat[i, df.columns.get_loc(c)] = DESTAQUE
        return estilos

    pd.set_option("styler.render.max_elements", max(262144, df.size + 1))
    return df.style.apply(estilo, axis=None).format(formatos, na_rep="")


def grafico_erros(resumo: pd.DataFrame) -> alt.Chart:
    dados = resumo[resumo["tipo"] == "Erro"]
    return (
        alt.Chart(dados)
        .mark_bar(color=COR_SERIE, cornerRadiusEnd=4, size=18)
        .encode(
            x=alt.X("pedidos_afetados:Q", title="Pedidos afetados", axis=alt.Axis(tickMinStep=1, grid=True)),
            y=alt.Y("categoria:N", sort="-x", title=None, axis=alt.Axis(labelLimit=260)),
            tooltip=[
                alt.Tooltip("categoria:N", title="Tipo de erro"),
                alt.Tooltip("pedidos_afetados:Q", title="Pedidos afetados"),
                alt.Tooltip("celulas_corrigidas:Q", title="Células corrigidas"),
                alt.Tooltip("pedidos_pendentes:Q", title="Pendentes de revisão"),
            ],
        )
        .properties(height=max(160, 34 * len(dados)))
    )


# --------------------------------------------------------------------------- interface
st.set_page_config(page_title="Limpeza de Pedidos de Vendas", page_icon="🧹", layout="wide")
st.title("🧹 Limpeza de Pedidos de Vendas")
st.caption("Envie o CSV de pedidos: a limpeza e as validações rodam automaticamente, "
           "e você baixa a base limpa em CSV ou Excel.")

with st.expander("Formato esperado do arquivo"):
    st.markdown("CSV com cabeçalho na primeira linha, separado por `;`, `,` ou tabulação "
                "(UTF-8 ou Windows-1252). Acentos e maiúsculas nos nomes das colunas são indiferentes.")
    st.table(pd.DataFrame(DESCRICAO_COLUNAS.items(), columns=["Coluna", "Conteúdo"]))
    modelo = ";".join(DESCRICAO_COLUNAS).encode("utf-8-sig") + b"\r\n"
    st.download_button("Baixar modelo vazio (.csv)", modelo, "modelo_pedidos_vendas.csv", "text/csv")

col_upload, col_exemplo = st.columns([4, 1], vertical_alignment="bottom")
arquivo = col_upload.file_uploader("Arquivo de pedidos", type=["csv", "txt", "xlsx"],
                                   help="CSV exportado do sistema ou do Excel. Também aceita .xlsx.")
if col_exemplo.button("Usar arquivo de exemplo", key="exemplo", width="stretch"):
    st.session_state["usar_exemplo"] = True

if arquivo is not None:
    st.session_state["usar_exemplo"] = False
    conteudo, nome = arquivo.getvalue(), arquivo.name
elif st.session_state.get("usar_exemplo") and EXEMPLO.exists():
    conteudo, nome = EXEMPLO.read_bytes(), EXEMPLO.name
    st.info(f"Usando o arquivo de exemplo **{EXEMPLO.name}**.")
else:
    st.info("Envie um arquivo CSV para começar, ou use o arquivo de exemplo.")
    st.stop()

try:
    r = processar(conteudo, nome)
except ValueError as erro:
    st.error(f"Não foi possível processar o arquivo: {erro}")
    st.stop()
except Exception as erro:  # noqa: BLE001 - mostrar qualquer falha inesperada ao usuário
    st.error("Erro inesperado ao processar o arquivo. Verifique se é um CSV de pedidos válido.")
    with st.expander("Detalhes técnicos"):
        st.exception(erro)
    st.stop()

abas, ind, bruto = r["abas"], r["indicadores"], r["bruto"]
tratados, validos, log = abas["pedidos_tratados"], abas["pedidos_validos"], abas["log_alteracoes"]
resumo, validacoes = abas["resumo_qualidade"], abas["validacoes"]

# --- indicadores ---------------------------------------------------------------
st.subheader("Resultado da limpeza")
m1, m2, m3, m4, m5 = st.columns(5)
m1.metric("Pedidos analisados", ind["pedidos"])
m2.metric("Erros encontrados", ind["erros"],
          help="Ocorrências de erro por pedido e tipo (ex.: um pedido com receita em texto "
               "e nome fora do padrão conta 2).")
m3.metric("Melhorias aplicadas", ind["melhorias"], help="Células alteradas automaticamente pela limpeza.")
m4.metric("Pendências para revisão", ind["pendencias"],
          help=f"{ind['pedidos_pendentes']} pedido(s) precisam de confirmação da área de negócio.",
          delta=f"{ind['pedidos_pendentes']} pedido(s)", delta_color="off")
m5.metric("Pedidos sem problemas", f"{ind['pct_ok_depois']:.0%}",
          delta=f"{(ind['pct_ok_depois'] - ind['pct_ok_antes']) * 100:+.0f} p.p. (antes {ind['pct_ok_antes']:.0%})",
          help="Percentual de pedidos sem nenhum erro ou pendência, antes e depois da limpeza.")

aprovadas = int(validacoes["aprovado"].sum())
total_validacoes = len(validacoes)
if aprovadas == total_validacoes:
    st.success(f"✅ {aprovadas}/{total_validacoes} verificações automáticas aprovadas.")
else:
    st.error(f"❌ {total_validacoes - aprovadas} de {total_validacoes} verificações reprovadas. "
             "Confira a aba Validações antes de usar a base.")

# --- abas ------------------------------------------------------------------------
tab_resumo, tab_dados, tab_pend, tab_valid, tab_log, tab_orig = st.tabs([
    "📊 Erros e melhorias", "✅ Dados limpos", f"⚠️ Pendências ({ind['pedidos_pendentes']})",
    "🔎 Validações", f"📝 Log ({len(log)})", "📄 Original",
])

with tab_resumo:
    erros = resumo[resumo["tipo"] == "Erro"]
    melhorias = resumo[resumo["tipo"] == "Melhoria"]
    c1, c2 = st.columns([3, 2])
    with c1:
        st.markdown("**Pedidos afetados por tipo de erro**")
        if erros.empty:
            st.write("Nenhum erro encontrado. 🎉")
        else:
            st.altair_chart(grafico_erros(resumo), width="stretch")
    with c2:
        st.markdown("**Detalhamento**")
        st.dataframe(
            resumo, hide_index=True, width="stretch",
            column_config={
                "categoria": "Categoria", "tipo": "Tipo",
                "pedidos_afetados": st.column_config.NumberColumn("Pedidos afetados"),
                "celulas_corrigidas": st.column_config.NumberColumn("Células corrigidas"),
                "pedidos_pendentes": st.column_config.NumberColumn("Pendentes"),
            },
        )
        if not melhorias.empty:
            st.caption("**Melhoria** = ajuste técnico que não indica erro na origem "
                       "(ex.: texto '2024-01-22' convertido para o tipo data).")

with tab_dados:
    so_validos = st.toggle("Mostrar apenas pedidos válidos (sem duplicatas confirmadas)", value=True)
    exibir = validos if so_validos else tratados
    st.caption(f"{len(exibir)} pedido(s). Células em azul foram alteradas pela limpeza.")
    st.dataframe(tabela_destacada(exibir, log), hide_index=True, width="stretch", height=420)

with tab_pend:
    pendentes = tratados[tratados["pendencias"].fillna("") != ""]
    if pendentes.empty:
        st.success("Nenhuma pendência: todos os pedidos foram corrigidos automaticamente.")
    else:
        st.caption("Itens que a limpeza não resolve sozinha: confirme com a área de negócio.")
        st.dataframe(
            pendentes[["id_pedido", "cliente", "receita", "status_registro", "duplicado_de", "pendencias"]],
            hide_index=True, width="stretch",
            column_config={"receita": st.column_config.NumberColumn("receita", format="R$ %.2f")},
        )

with tab_valid:
    st.dataframe(
        validacoes.assign(aprovado=validacoes["aprovado"].map({True: "✅", False: "❌"})),
        hide_index=True, width="stretch",
        column_config={"verificacao": "Verificação", "aprovado": "Resultado", "detalhe": "Detalhe"},
    )

with tab_log:
    st.caption("Cada célula alterada, com o valor original, o novo valor e a regra aplicada.")
    st.dataframe(log, hide_index=True, width="stretch", height=420)

with tab_orig:
    st.caption("Dados exatamente como foram enviados. Esta cópia também vai na aba "
               "'original' do Excel, para rastrear a origem.")
    st.dataframe(bruto, hide_index=True, width="stretch", height=420)

# --- downloads -------------------------------------------------------------------
st.subheader("Baixar planilha limpa")
base = Path(nome).stem
d1, d2 = st.columns(2)
with d1:
    st.markdown("**CSV**: só os dados, para importar em outros sistemas")
    conteudo_csv = st.radio("Conteúdo", ["Pedidos válidos", "Todos os pedidos (com duplicatas marcadas)"],
                            horizontal=True, key="conteudo_csv")
    formato_csv = st.radio("Formato", ["Excel Brasil ( ; e vírgula decimal)", "Padrão internacional ( , e ponto decimal)"],
                           horizontal=True, key="formato_csv")
    dados_csv = validos if conteudo_csv == "Pedidos válidos" else tratados
    separador = ";" if formato_csv.startswith("Excel Brasil") else ","
    st.download_button("⬇️ Baixar CSV", ex.gerar_csv(dados_csv, separador), f"{base}_limpo.csv",
                       "text/csv", type="primary", width="stretch")
with d2:
    st.markdown("**Excel**: planilha completa com rastreabilidade")
    st.caption("Abas: pedidos_tratados, pedidos_validos, log_alteracoes, resumo_qualidade, "
               "validacoes, metadados e original.")
    st.download_button(
        "⬇️ Baixar Excel", r["excel"], f"{base}_limpo.xlsx",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary", width="stretch",
    )
