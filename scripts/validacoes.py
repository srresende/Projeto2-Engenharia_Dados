"""
Validações e indicadores de qualidade do tratamento de pedidos de vendas.

Usado pelo script de linha de comando, pelo app Streamlit e pela suíte de testes:
qualquer arquivo tratado passa pelas mesmas verificações.

    validar(bruto, tratados, log)        -> tabela de verificações (aprovado/reprovado)
    resumo_qualidade(tratados, log)      -> erros e melhorias por categoria
    indicadores(tratados, log)           -> números de destaque (erros, melhorias, pendências)
"""
from __future__ import annotations

import pandas as pd

import tratar_pedidos_vendas as t

# Regra do log -> (categoria, é erro da origem?). Ordem importa: primeiro trecho que casar.
CATEGORIAS_LOG = [
    ("conversão de tipo", "Conversão de tipo", False),
    ("observação convertida", "Aviso digitado em Observações", True),
    ("status movido", "Status na coluna errada", True),
    ("inferid", "Campo obrigatório vazio", True),
    ("mesmo cliente", "Duplicidade", True),
    ("marcado como duplicado", "Duplicidade", True),
    ("receita", "Receita em formato de texto", True),
    ("data", "Data fora do padrão", True),
    ("nome padronizado", "Nome fora do padrão", True),
    ("região padronizada", "Região fora do padrão", True),
    ("status padronizado", "Status fora do padrão", True),
    ("espaços", "Espaços extras", True),
]
# Trecho da pendência -> categoria
CATEGORIAS_PENDENCIA = [
    ("receita ausente", "Receita ausente"),
    ("possível duplicado", "Duplicidade"),
    ("inferid", "Campo obrigatório vazio"),
    ("representante ausente", "Campo obrigatório vazio"),
    ("cliente ausente", "Campo obrigatório vazio"),
    ("produto ausente", "Campo obrigatório vazio"),
    ("id do pedido", "Campo obrigatório vazio"),
    ("região ausente", "Região fora do padrão"),
    ("status ausente", "Status fora do padrão"),
    ("data inválida", "Data fora do padrão"),
]
COLUNAS_RESUMO = ["categoria", "tipo", "pedidos_afetados", "celulas_corrigidas", "pedidos_pendentes"]


def categoria_log(regra: str) -> tuple[str, bool]:
    r = str(regra).lower()
    for trecho, categoria, erro in CATEGORIAS_LOG:
        if trecho in r:
            return categoria, erro
    return "Outros ajustes", True


def categoria_pendencia(item: str) -> str:
    i = str(item).lower()
    for trecho, categoria in CATEGORIAS_PENDENCIA:
        if trecho in i:
            return categoria
    return "Outras pendências"


def _pendencias(tratados: pd.DataFrame) -> list[tuple[int, str]]:
    """(linha_origem, categoria) para cada item de pendência."""
    pares = []
    for linha, texto in zip(tratados["linha_origem"], tratados["pendencias"].fillna("")):
        for item in filter(None, (p.strip() for p in str(texto).split(";"))):
            pares.append((linha, categoria_pendencia(item)))
    return pares


# --------------------------------------------------------------------------- resumo
def resumo_qualidade(tratados: pd.DataFrame, log: pd.DataFrame) -> pd.DataFrame:
    cats = log["regra"].map(categoria_log)
    log = log.assign(categoria=cats.map(lambda c: c[0]), erro=cats.map(lambda c: c[1]))
    pend = pd.DataFrame(_pendencias(tratados), columns=["linha_origem", "categoria"])

    linhas = []
    for categoria in sorted(set(log["categoria"]) | set(pend["categoria"])):
        do_log = log[log["categoria"] == categoria]
        da_pend = pend[pend["categoria"] == categoria]
        erro = bool(do_log["erro"].iloc[0]) if len(do_log) else True
        afetados = set(do_log["linha_origem"]) | set(da_pend["linha_origem"])
        linhas.append({
            "categoria": categoria,
            "tipo": "Erro" if erro else "Melhoria",
            "pedidos_afetados": len(afetados),
            "celulas_corrigidas": len(do_log),
            "pedidos_pendentes": da_pend["linha_origem"].nunique(),
        })
    resumo = pd.DataFrame(linhas, columns=COLUNAS_RESUMO)
    return resumo.sort_values(["tipo", "pedidos_afetados"], ascending=[True, False]).reset_index(drop=True)


def indicadores(tratados: pd.DataFrame, log: pd.DataFrame) -> dict:
    erros_log = {(linha, categoria_log(r)[0]) for linha, r in zip(log["linha_origem"], log["regra"])
                 if categoria_log(r)[1]}
    pend = set(_pendencias(tratados))
    ocorrencias = erros_log | pend
    total = len(tratados)
    com_erro = {linha for linha, _ in ocorrencias}
    com_pendencia = {linha for linha, _ in pend}
    return {
        "pedidos": total,
        "erros": len(ocorrencias),                     # pares (pedido, tipo de erro)
        "pedidos_com_erro": len(com_erro),
        "melhorias": len(log),                         # células alteradas
        "corrigidos_automaticamente": len(erros_log - pend),
        "pendencias": len(pend),
        "pedidos_pendentes": len(com_pendencia),
        "duplicados_confirmados": int((tratados["status_registro"] == "duplicado_confirmado").sum()),
        "pct_ok_antes": 1 - len(com_erro) / total if total else 1.0,
        "pct_ok_depois": 1 - len(com_pendencia) / total if total else 1.0,
    }


# --------------------------------------------------------------------------- validações
def validar(bruto: pd.DataFrame, tratados: pd.DataFrame, log: pd.DataFrame) -> pd.DataFrame:
    """Verifica o resultado do tratamento. Cada linha: verificação, aprovado, detalhe."""
    origem = t.normalizar_colunas(bruto).rename(columns=t.COLUNAS)
    origem.index = origem.index + 2                      # mesma chave de linha_origem
    df = tratados.set_index("linha_origem", drop=False)
    pend = df["pendencias"].fillna("").str.lower()
    ids = df["id_pedido"]
    resultados: list[dict] = []

    def checar(nome: str, falhas: list, ok: str) -> None:
        exemplo = ", ".join(map(str, falhas[:10])) + (" ..." if len(falhas) > 10 else "")
        resultados.append({
            "verificacao": nome,
            "aprovado": not falhas,
            "detalhe": ok if not falhas else f"{len(falhas)} ocorrência(s): {exemplo}",
        })

    def ref(linha) -> str:
        i = ids.get(linha)
        return f"pedido {i}" if not t.vazio(i) else f"linha {linha}"

    # 1. Linhagem: nenhuma linha perdida ou criada
    diferenca = sorted(set(origem.index) ^ set(df.index))
    checar("Nenhuma linha perdida ou criada", [f"linha {l}" for l in diferenca],
           f"{len(df)} linhas na origem e no resultado")

    # 2. IDs
    falhas = [f"linha {l}" for l in df.index[ids.isna()]]
    falhas += [f"pedido {i} repetido" for i in ids[ids.duplicated() & ids.notna()].unique()]
    checar("ID do pedido preenchido e único", falhas, "todos os IDs preenchidos e únicos")

    # 3. Receita
    numerica = pd.api.types.is_numeric_dtype(df["receita"])
    checar("Receita em formato numérico", [] if numerica else ["coluna receita não é numérica"],
           "coluna receita é numérica")
    sem_flag = df.index[df["receita"].isna() & ~df["flag_receita_ausente"].astype(bool)]
    checar("Receita ausente sinalizada", [ref(l) for l in sem_flag],
           f"{int(df['receita'].isna().sum())} receita(s) ausente(s), todas sinalizadas")

    # 4. Datas
    datas = pd.to_datetime(df["data_pedido"], errors="coerce")
    falhas = [ref(l) for l in df.index[datas.isna() & ~pend.str.contains("data inválida")]]
    falhas += [ref(l) for l in df.index[datas.notna() & (datas != datas.dt.normalize())]]
    checar("Datas válidas e sem hora", falhas, "todas as datas válidas no formato AAAA-MM-DD")

    # 5. Domínios
    for coluna, dominio, rotulo in [("regiao", t.REGIOES.values(), "Região"),
                                    ("status", t.STATUS.values(), "Status")]:
        fora = df.index[~df[coluna].isin(list(dominio))
                        & ~pend.str.contains(f"{rotulo.lower()} ausente ou fora do domínio")]
        checar(f"{rotulo} dentro do padrão (ou sinalizada)", [ref(l) for l in fora],
               f"valores aceitos: {', '.join(sorted(set(dominio)))}")

    # 6. Nomes
    falhas = [f"{ref(l)} ({c})" for c in ("cliente", "representante") for l in df.index
              if not t.vazio(df.at[l, c]) and df.at[l, c] != t.nome_proprio(df.at[l, c])]
    checar("Nomes com caixa padronizada", falhas, "cliente e representante padronizados")

    # 7. Obrigatórios
    falhas = []
    for coluna, termo in [("cliente", "cliente ausente"), ("produto", "produto ausente"),
                          ("representante", "representante ausente")]:
        falhas += [f"{ref(l)} ({coluna})" for l in df.index[df[coluna].isna() & ~pend.str.contains(termo)]]
    checar("Campos obrigatórios preenchidos (ou sinalizados)", falhas, "nenhum campo vazio sem sinalização")

    # 8. Inferências sinalizadas
    falhas = []
    for coluna, flag in [("regiao", "regiao_inferida"), ("representante", "representante_inferido")]:
        comum = origem.index.intersection(df.index)
        preenchidas = [l for l in comum if t.vazio(origem.at[l, coluna]) and not t.vazio(df.at[l, coluna])]
        falhas += [f"{ref(l)} ({coluna})" for l in preenchidas if not df.at[l, flag]]
    checar("Valores inferidos sinalizados", falhas, "todo valor preenchido automaticamente está marcado")

    # 9. Duplicatas
    validos = df[df["status_registro"] != "duplicado_confirmado"]
    chave_dup = ["cliente", "produto", "receita", "data_pedido", "regiao", "representante", "status"]
    repetidos = validos[validos.duplicated(subset=chave_dup, keep="first") & validos["receita"].notna()]
    checar("Sem duplicata exata entre os pedidos válidos", [ref(l) for l in repetidos.index],
           f"{len(df) - len(validos)} duplicata(s) confirmada(s) removida(s) da base válida")
    refs = df["duplicado_de"].dropna()
    orfas = [ref(l) for l in refs.index if refs[l] not in set(ids.dropna())]
    checar("Referências de duplicata válidas", orfas, "todo duplicado_de aponta para um pedido existente")

    # 10. Log
    registrados = set(zip(log["linha_origem"], log["coluna"]))
    comum = origem.index.intersection(df.index)
    falhas = [f"{ref(l)} ({c})" for c in t.COLUNAS.values() if c != "id_pedido" for l in comum
              if t.mudou(origem.at[l, c], df.at[l, c]) and (l, c) not in registrados]
    checar("Toda alteração registrada no log", falhas, f"{len(log)} alteração(ões) registrada(s)")

    falhas = []
    colunas_negocio = [c for c in t.COLUNAS.values() if c in origem.columns]
    for r in log[log["coluna"].isin(colunas_negocio)].itertuples():
        if r.linha_origem in origem.index:
            valor = origem.at[r.linha_origem, r.coluna]
            if str(r.valor_original) != ("" if t.vazio(valor) else str(valor)):
                falhas.append(f"{ref(r.linha_origem)} ({r.coluna})")
    checar("Log fiel ao valor original", falhas, "valor_original do log confere com a origem")

    return pd.DataFrame(resultados, columns=["verificacao", "aprovado", "detalhe"])
