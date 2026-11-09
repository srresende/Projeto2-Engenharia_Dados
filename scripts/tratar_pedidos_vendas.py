"""
Tratamento da planilha de pedidos de vendas.

Lê a planilha de origem (.xlsx ou .csv, sem alterá-la) e gera uma planilha
tratada (.xlsx) com as abas:

    pedidos_tratados  todas as linhas corrigidas + colunas de controle
    pedidos_validos   igual à anterior, sem as duplicatas confirmadas
    log_alteracoes    cada célula alterada: valor original, valor novo e regra
    resumo_qualidade  erros encontrados e melhorias aplicadas por categoria
    validacoes        verificações automáticas do resultado (scripts/validacoes.py)
    metadados         arquivo de origem, hash SHA-256, data de processamento, contagens
    original          cópia fiel dos dados de origem (linhagem)

Correções aplicadas:
    1. Receita em texto ("R$ 4320.00") -> número; receitas vazias sinalizadas
    2. Datas em formatos variados -> data única (exibida como AAAA-MM-DD)
    3. Nomes, região e representante com caixa padronizada
    4. Status padronizado (inclui "closed won") e status digitado em Observações movido
    5. Região/representante ausentes inferidos (cliente -> representante -> região)
    6. Duplicatas detectadas (cliente + produto + receita): confirmadas x possíveis
    7. Observações separadas em colunas de flag
    8. Colunas renomeadas para snake_case

Uso:
    pip install -r requirements.txt
    python scripts/tratar_pedidos_vendas.py [--entrada ARQ] [--saida ARQ]
"""
from __future__ import annotations

import argparse
import hashlib
import io
import re
import shutil
import unicodedata
from datetime import datetime
from pathlib import Path

import pandas as pd

RAIZ = Path(__file__).resolve().parents[1]
ENTRADA_PADRAO = RAIZ / "data" / "Planilha_pedidos_vendas.xlsx"
SAIDA_PADRAO = RAIZ / "data" / "Planilha_pedidos_vendas_tratada.xlsx"

COLUNAS = {
    "ID Pedido": "id_pedido",
    "Nome do Cliente": "cliente",
    "Região": "regiao",
    "Representante": "representante",
    "Data do Pedido": "data_pedido",
    "Receita": "receita",
    "Produto": "produto",
    "Status": "status",
    "Observações": "observacoes",
}

# Dicionários de domínio (chaves sem acento e em minúsculas)
REGIOES = {
    "sudeste": "Sudeste",
    "sul": "Sul",
    "nordeste": "Nordeste",
    "norte": "Norte",
    "centro-oeste": "Centro-Oeste",
    "centro oeste": "Centro-Oeste",
}
STATUS = {
    "fechado ganho": "Fechado Ganho",
    "ganho": "Fechado Ganho",
    "closed won": "Fechado Ganho",
    "won": "Fechado Ganho",
    "fechado perdido": "Fechado Perdido",
    "perdido": "Fechado Perdido",
    "closed lost": "Fechado Perdido",
    "lost": "Fechado Perdido",
}
FLAGS_OBSERVACAO = {
    "duplicado": "flag_duplicado_informado",
    "possivel duplicado": "flag_duplicado_informado",
    "receita ausente": "flag_receita_ausente",
    "sem representante": "flag_sem_representante",
}
PARTICULAS = {"da", "de", "do", "das", "dos", "e"}
CODIFICACOES_CSV = ("utf-8-sig", "cp1252", "latin-1")
SEPARADORES_CSV = (";", ",", "\t", "|")

REGRA_TIPO_NUMERO = "conversão de tipo: texto -> número"
REGRA_TIPO_DATA = "conversão de tipo: texto -> data"

ORDEM_SAIDA = [
    "id_pedido", "data_pedido", "cliente", "regiao", "representante", "produto",
    "receita", "status", "observacoes",
    "status_registro", "duplicado_de", "regiao_inferida", "representante_inferido",
    "flag_duplicado_informado", "flag_receita_ausente", "flag_sem_representante",
    "pendencias", "observacoes_original", "linha_origem", "arquivo_origem",
]


# --------------------------------------------------------------------------- utilidades
def vazio(v) -> bool:
    try:
        return v is None or bool(pd.isna(v)) or (isinstance(v, str) and not v.strip())
    except (TypeError, ValueError):
        return False


def mudou(antes, depois) -> bool:
    if vazio(antes) and vazio(depois):
        return False
    if vazio(antes) != vazio(depois):
        return True
    datas = (datetime, pd.Timestamp)
    if isinstance(antes, datas) and isinstance(depois, datas):
        return pd.Timestamp(antes) != pd.Timestamp(depois)
    if isinstance(antes, str) != isinstance(depois, str):   # ex.: "R$ 4320.00" -> 4320.0
        return True
    return bool(antes != depois)


def chave(v) -> str:
    s = unicodedata.normalize("NFKD", str(v)).encode("ascii", "ignore").decode()
    return re.sub(r"\s+", " ", s).strip().lower()


def sha256(caminho: Path) -> str:
    h = hashlib.sha256()
    with open(caminho, "rb") as f:
        for bloco in iter(lambda: f.read(65536), b""):
            h.update(bloco)
    return h.hexdigest()


class Log:
    def __init__(self) -> None:
        self.registros: list[dict] = []

    def add(self, df: pd.DataFrame, idx, coluna: str, antes, depois, regra: str) -> None:
        self.registros.append({
            "id_pedido": df.at[idx, "id_pedido"],
            "linha_origem": df.at[idx, "linha_origem"],
            "coluna": coluna,
            "valor_original": "" if vazio(antes) else str(antes),
            "valor_novo": "" if vazio(depois) else str(depois),
            "regra": regra,
        })

    def to_frame(self) -> pd.DataFrame:
        cols = ["id_pedido", "linha_origem", "coluna", "valor_original", "valor_novo", "regra"]
        return pd.DataFrame(self.registros, columns=cols)


def aplicar(df: pd.DataFrame, coluna: str, func, regra, log: Log) -> None:
    """Aplica func à coluna, registrando no log apenas as células que mudaram.
    `regra` pode ser texto fixo ou função (valor_original -> texto)."""
    novos = df[coluna].map(func)
    for idx in df.index:
        antes, depois = df.at[idx, coluna], novos.at[idx]
        if mudou(antes, depois):
            log.add(df, idx, coluna, antes, depois, regra(antes) if callable(regra) else regra)
    df[coluna] = novos


# --------------------------------------------------------------------------- leitura
def ler_arquivo(conteudo: bytes, nome: str) -> pd.DataFrame:
    """Lê .csv (codificação e separador detectados) ou .xlsx, mantendo os valores brutos."""
    if not conteudo or not conteudo.strip():
        raise ValueError("O arquivo está vazio.")
    sufixo = Path(nome).suffix.lower()
    if sufixo in (".xlsx", ".xlsm"):
        return pd.read_excel(io.BytesIO(conteudo), sheet_name=0, dtype=object)
    if sufixo not in (".csv", ".txt"):
        raise ValueError(f"Formato não suportado: {sufixo or 'sem extensão'}. Envie um arquivo .csv.")

    for codificacao in CODIFICACOES_CSV:
        try:
            texto = conteudo.decode(codificacao)
            break
        except UnicodeDecodeError:
            continue
    cabecalho = texto.lstrip().splitlines()[0]
    separador = max(SEPARADORES_CSV, key=cabecalho.count)
    df = pd.read_csv(io.StringIO(texto), sep=separador, dtype=object)
    df.columns = [str(c).strip() for c in df.columns]
    # Linhas totalmente vazias são descartadas; o índice original é mantido
    # para que linha_origem continue apontando para a linha do arquivo.
    return df.dropna(how="all")


def normalizar_colunas(bruto: pd.DataFrame) -> pd.DataFrame:
    """Reconhece cabeçalhos com variação de acento/caixa/espaço ou já em snake_case."""
    por_chave = {chave(c): c for c in COLUNAS}
    por_chave.update({chave(s): c for c, s in COLUNAS.items()})
    renome = {c: por_chave[chave(c)] for c in bruto.columns if chave(c) in por_chave}
    return bruto.rename(columns=renome)


# --------------------------------------------------------------------------- conversores
def nome_proprio(v):
    if vazio(v):
        return pd.NA
    partes = str(v).strip().lower().split()
    return " ".join(
        p if i > 0 and p in PARTICULAS else p[:1].upper() + p[1:]
        for i, p in enumerate(partes)
    )


def padronizar_regiao(v):
    if vazio(v):
        return pd.NA
    return REGIOES.get(chave(v), str(v).strip().title())


def padronizar_status(v):
    if vazio(v):
        return pd.NA
    return STATUS.get(chave(v), str(v).strip())


def para_numero(v):
    """Converte 'R$ 4320.00', 'R$ 4.320,00', '4320' etc. em float."""
    if vazio(v):
        return float("nan")
    if isinstance(v, (int, float)):
        return float(v)
    s = re.sub(r"[^\d,.\-]", "", str(v))
    if not s or not re.search(r"\d", s):
        return float("nan")
    if "," in s and "." in s:
        if s.rfind(",") > s.rfind("."):          # 4.320,00 (BR)
            s = s.replace(".", "").replace(",", ".")
        else:                                      # 4,320.00 (US)
            s = s.replace(",", "")
    elif "," in s:                                 # 4320,00
        s = s.replace(",", ".")
    elif re.fullmatch(r"\d{1,3}(\.\d{3})+", s):    # 4.320 (milhar BR)
        s = s.replace(".", "")
    try:
        return float(s)
    except ValueError:
        return float("nan")


def para_data(v):
    if vazio(v):
        return pd.NaT
    if isinstance(v, (datetime, pd.Timestamp)):
        return pd.Timestamp(v).normalize()
    if isinstance(v, (int, float)):                # número de série do Excel
        return pd.Timestamp("1899-12-30") + pd.Timedelta(days=float(v))
    s = str(v).strip()
    if re.fullmatch(r"\d{4}-\d{2}-\d{2}.*", s):
        return pd.to_datetime(s, errors="coerce").normalize()
    return pd.to_datetime(s, dayfirst=True, errors="coerce")


def regra_receita(antes) -> str:
    if isinstance(antes, str) and re.fullmatch(r"\s*-?\d+([.,]\d+)?\s*", antes):
        return REGRA_TIPO_NUMERO
    return "receita convertida de texto para número"


def regra_data(antes) -> str:
    if isinstance(antes, str):
        if re.fullmatch(r"\s*\d{4}-\d{2}-\d{2}\s*", antes):
            return REGRA_TIPO_DATA
        return "data em formato fora do padrão convertida"
    if isinstance(antes, (datetime, pd.Timestamp)):
        return "data normalizada (hora removida)"
    return "data convertida para tipo data"


# --------------------------------------------------------------------------- etapas
def converter_ids(serie: pd.Series) -> pd.Series:
    """IDs numéricos inteiros viram Int64; caso contrário são mantidos como texto."""
    preenchidos = serie.map(lambda v: not vazio(v))
    numeros = pd.to_numeric(serie, errors="coerce")
    if numeros.notna().sum() == preenchidos.sum() and (numeros.dropna() % 1 == 0).all():
        return numeros.astype("Int64")
    return serie.map(lambda v: pd.NA if vazio(v) else str(v).strip())


def tratar_observacoes(df: pd.DataFrame, log: Log) -> None:
    """Move status digitado em Observações e converte avisos em flags."""
    for col in set(FLAGS_OBSERVACAO.values()):
        df[col] = False
    for idx in df.index:
        obs = df.at[idx, "observacoes"]
        if vazio(obs):
            continue
        k = chave(obs)
        if k in STATUS and vazio(df.at[idx, "status"]):
            log.add(df, idx, "status", df.at[idx, "status"], STATUS[k], "status movido de Observações")
            log.add(df, idx, "observacoes", obs, None, "status movido para coluna Status")
            df.at[idx, "status"] = STATUS[k]
            df.at[idx, "observacoes"] = pd.NA
        elif k in FLAGS_OBSERVACAO:
            flag = FLAGS_OBSERVACAO[k]
            df.at[idx, flag] = True
            log.add(df, idx, "observacoes", obs, None, f"observação convertida em {flag}")
            df.at[idx, "observacoes"] = pd.NA


def mapa_unico(df: pd.DataFrame, origem: str, destino: str) -> dict:
    """origem -> destino, apenas quando a origem tem um único destino conhecido."""
    base = df.dropna(subset=[origem, destino]).groupby(origem)[destino].unique()
    return {k: v[0] for k, v in base.items() if len(v) == 1}


def inferir_ausentes(df: pd.DataFrame, log: Log) -> None:
    df["regiao_inferida"] = False
    df["representante_inferido"] = False

    regiao_por_cliente = mapa_unico(df, "cliente", "regiao")
    regiao_por_rep = mapa_unico(df, "representante", "regiao")
    for idx in df.index[df["regiao"].isna()]:
        cliente, rep = df.at[idx, "cliente"], df.at[idx, "representante"]
        if not vazio(cliente) and cliente in regiao_por_cliente:
            novo, regra = regiao_por_cliente[cliente], "região inferida por outro pedido do mesmo cliente"
        elif not vazio(rep) and rep in regiao_por_rep:
            novo, regra = regiao_por_rep[rep], "região inferida pelo representante (atende uma única região)"
        else:
            continue
        log.add(df, idx, "regiao", None, novo, regra)
        df.at[idx, "regiao"] = novo
        df.at[idx, "regiao_inferida"] = True

    rep_por_cliente = mapa_unico(df, "cliente", "representante")
    rep_por_regiao = mapa_unico(df, "regiao", "representante")
    for idx in df.index[df["representante"].isna()]:
        cliente, regiao = df.at[idx, "cliente"], df.at[idx, "regiao"]
        if not vazio(cliente) and cliente in rep_por_cliente:
            novo, regra = rep_por_cliente[cliente], "representante inferido por outro pedido do mesmo cliente"
        elif not vazio(regiao) and regiao in rep_por_regiao:
            novo, regra = rep_por_regiao[regiao], "representante inferido pela região (único representante)"
        else:
            continue
        log.add(df, idx, "representante", None, novo, regra)
        df.at[idx, "representante"] = novo
        df.at[idx, "representante_inferido"] = True


def marcar_duplicatas(df: pd.DataFrame, log: Log) -> None:
    """
    Agrupa por cliente + produto + receita. O menor ID do grupo é o registro de referência.
    Demais linhas: 'duplicado_confirmado' se data, região, representante e status também
    coincidem; caso contrário 'possivel_duplicado' (decisão da área de negócio).
    """
    df["status_registro"] = "ativo"
    df["duplicado_de"] = pd.Series([pd.NA] * len(df), index=df.index, dtype=df["id_pedido"].dtype)
    comparar = ["data_pedido", "regiao", "representante", "status"]

    for _, grupo in df.groupby(["cliente", "produto", "receita"], dropna=True):
        if len(grupo) < 2:
            continue
        grupo = grupo.sort_values("id_pedido")
        ref = grupo.iloc[0]
        for idx, linha in grupo.iloc[1:].iterrows():
            identico = all(not mudou(ref[c], linha[c]) for c in comparar)
            situacao = "duplicado_confirmado" if identico else "possivel_duplicado"
            df.at[idx, "status_registro"] = situacao
            df.at[idx, "duplicado_de"] = ref["id_pedido"]
            log.add(df, idx, "status_registro", "ativo", situacao,
                    "mesmo cliente/produto/receita de pedido anterior (ver duplicado_de)")

    # Duplicata apontada na origem, mas não detectada pela regra: revisar manualmente
    sem_par = df["flag_duplicado_informado"] & (df["status_registro"] == "ativo")
    for idx in df.index[sem_par]:
        df.at[idx, "status_registro"] = "possivel_duplicado"
        log.add(df, idx, "status_registro", "ativo", "possivel_duplicado",
                "marcado como duplicado na origem, sem par encontrado")


def montar_pendencias(df: pd.DataFrame) -> None:
    def pend(r) -> str:
        itens = []
        if vazio(r["id_pedido"]):
            itens.append("id do pedido ausente ou inválido")
        if vazio(r["cliente"]):
            itens.append("cliente ausente")
        if vazio(r["produto"]):
            itens.append("produto ausente")
        if vazio(r["receita"]):
            itens.append("receita ausente: confirmar valor")
        if r["status_registro"] == "possivel_duplicado":
            ref = "" if vazio(r["duplicado_de"]) else f" do pedido {r['duplicado_de']}"
            itens.append(f"possível duplicado{ref}: validar com negócio")
        if r["regiao_inferida"]:
            itens.append("região inferida: confirmar")
        if r["representante_inferido"]:
            itens.append("representante inferido: confirmar")
        if vazio(r["regiao"]) or r["regiao"] not in REGIOES.values():
            itens.append("região ausente ou fora do domínio")
        if vazio(r["representante"]):
            itens.append("representante ausente")
        if vazio(r["status"]) or r["status"] not in STATUS.values():
            itens.append("status ausente ou fora do domínio")
        if vazio(r["data_pedido"]):
            itens.append("data inválida")
        return "; ".join(itens)

    df["pendencias"] = df.apply(pend, axis=1)


# --------------------------------------------------------------------------- pipeline
def tratar(bruto: pd.DataFrame, nome_arquivo: str) -> tuple[pd.DataFrame, pd.DataFrame]:
    bruto = normalizar_colunas(bruto)
    faltando = [c for c in COLUNAS if c not in bruto.columns]
    if faltando:
        raise ValueError(f"Colunas esperadas não encontradas: {faltando}")
    if bruto.empty:
        raise ValueError("O arquivo não tem linhas de dados.")

    df = bruto.rename(columns=COLUNAS)[list(COLUNAS.values())].copy()
    df["linha_origem"] = df.index + 2            # +1 cabeçalho, +1 base 1
    df["arquivo_origem"] = nome_arquivo
    df["observacoes_original"] = df["observacoes"]
    df["id_pedido"] = converter_ids(df["id_pedido"])
    log = Log()

    aplicar(df, "cliente", nome_proprio, "nome padronizado (caixa)", log)
    aplicar(df, "representante", nome_proprio, "nome padronizado (caixa)", log)
    aplicar(df, "regiao", padronizar_regiao, "região padronizada (domínio)", log)
    aplicar(df, "produto", lambda v: pd.NA if vazio(v) else str(v).strip(), "espaços removidos", log)
    aplicar(df, "status", padronizar_status, "status padronizado (domínio)", log)
    aplicar(df, "receita", para_numero, regra_receita, log)
    aplicar(df, "data_pedido", para_data, regra_data, log)
    aplicar(df, "observacoes", lambda v: pd.NA if vazio(v) else str(v).strip(), "espaços removidos", log)

    tratar_observacoes(df, log)
    df["receita"] = df["receita"].astype(float)
    df["flag_receita_ausente"] = df["flag_receita_ausente"] | df["receita"].isna()
    df["flag_sem_representante"] = df["flag_sem_representante"] | df["representante"].isna()

    inferir_ausentes(df, log)
    marcar_duplicatas(df, log)
    montar_pendencias(df)

    df["data_pedido"] = pd.to_datetime(df["data_pedido"])
    df = df.sort_values("id_pedido", na_position="last").reset_index(drop=True)
    return df[ORDEM_SAIDA], log.to_frame()


# --------------------------------------------------------------------------- CLI
def main() -> None:
    import exportacao  # import tardio: exportacao depende deste módulo

    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--entrada", type=Path, default=ENTRADA_PADRAO)
    ap.add_argument("--saida", type=Path, default=SAIDA_PADRAO)
    args = ap.parse_args()

    entrada, saida = args.entrada.resolve(), args.saida.resolve()
    if entrada == saida:
        raise SystemExit("A saída não pode sobrescrever o arquivo de origem.")

    bruto = ler_arquivo(entrada.read_bytes(), entrada.name)
    tratados, log = tratar(bruto, entrada.name)
    abas = exportacao.montar_abas(bruto, tratados, log, entrada.name, sha256(entrada))

    saida.parent.mkdir(parents=True, exist_ok=True)
    if entrada.suffix.lower() in (".xlsx", ".xlsm"):
        # Copia o arquivo de origem e acrescenta as abas tratadas: a aba original mantém
        # os valores e formatos de célula da origem, garantindo a linhagem dos dados.
        # (Elementos gráficos/comentários do arquivo podem não ser preservados pelo openpyxl.)
        aba_origem = pd.ExcelFile(entrada, engine="openpyxl").sheet_names[0]
        shutil.copy2(entrada, saida)
        with pd.ExcelWriter(saida, engine="openpyxl", mode="a", if_sheet_exists="replace") as w:
            w.book[aba_origem].title = "original"
            exportacao.escrever_abas(w, abas)
            w.book.move_sheet("original", offset=len(w.book.sheetnames) - 1)
            w.book.active = 0
    else:
        saida.write_bytes(exportacao.gerar_excel(abas, bruto))

    print(f"Origem   : {entrada}")
    print(f"Saída    : {saida}")
    print(abas["metadados"].to_string(index=False))


if __name__ == "__main__":
    main()
