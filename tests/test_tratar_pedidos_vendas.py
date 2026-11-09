"""
Testes do tratamento da planilha de pedidos de vendas.

Três camadas:
    1. Unitários   - cada função de conversão/padronização isoladamente
    2. Pipeline    - função tratar() sobre um DataFrame sintético com casos controlados
    3. Ponta a ponta - executa o script real sobre data/Planilha_pedidos_vendas.xlsx
                       e valida a planilha gerada (correções, linhagem e formatação)

Uso:
    pip install -r requirements-dev.txt
    python -m pytest -v
"""
from __future__ import annotations

import hashlib
import math
import os
import subprocess
import sys
from datetime import datetime
from pathlib import Path

import openpyxl
import pandas as pd
import pytest

RAIZ = Path(__file__).resolve().parents[1]
SCRIPT = RAIZ / "scripts" / "tratar_pedidos_vendas.py"
ORIGEM = RAIZ / "data" / "Planilha_pedidos_vendas.xlsx"
sys.path.insert(0, str(SCRIPT.parent))

import tratar_pedidos_vendas as t  # noqa: E402


def sha256(caminho: Path) -> str:
    return hashlib.sha256(caminho.read_bytes()).hexdigest()


# =============================================================================
# 1. Testes unitários
# =============================================================================
class TestReceita:
    @pytest.mark.parametrize("entrada, esperado", [
        ("R$ 4320.00", 4320.0),        # formato encontrado na origem
        ("R$ 4.320,00", 4320.0),       # padrão brasileiro
        ("4,320.00", 4320.0),          # padrão americano com milhar
        ("4320,50", 4320.5),           # vírgula decimal
        ("4.320", 4320.0),             # ponto de milhar
        ("  R$990.00 ", 990.0),
        (1500, 1500.0),
        (1500.75, 1500.75),
    ])
    def test_converte_para_numero(self, entrada, esperado):
        assert t.para_numero(entrada) == pytest.approx(esperado)

    @pytest.mark.parametrize("entrada", [None, "", "   ", float("nan"), "R$"])
    def test_vazio_vira_nan(self, entrada):
        assert math.isnan(t.para_numero(entrada))


class TestData:
    @pytest.mark.parametrize("entrada", [
        datetime(2024, 1, 15),
        datetime(2024, 1, 15, 13, 45),   # hora é descartada
        45306,                            # número de série do Excel
        "15/01/2024",
        "2024-01-15",
    ])
    def test_converte_para_data(self, entrada):
        assert t.para_data(entrada) == pd.Timestamp("2024-01-15")

    def test_dia_primeiro_em_formato_ambiguo(self):
        assert t.para_data("03/02/2024") == pd.Timestamp("2024-02-03")

    @pytest.mark.parametrize("entrada", [None, "", "data inválida"])
    def test_invalida_vira_nat(self, entrada):
        assert pd.isna(t.para_data(entrada))


class TestTexto:
    @pytest.mark.parametrize("entrada, esperado", [
        ("PEDRO OLIVEIRA", "Pedro Oliveira"),
        ("maria souza", "Maria Souza"),
        ("  emília   dias ", "Emília Dias"),
        ("DÉRICO MENDES", "Dérico Mendes"),
        ("JOÃO DA SILVA", "João da Silva"),
        ("de souza", "De Souza"),          # partícula no início continua maiúscula
    ])
    def test_nome_proprio(self, entrada, esperado):
        assert t.nome_proprio(entrada) == esperado

    def test_nome_vazio(self):
        assert pd.isna(t.nome_proprio(None))

    @pytest.mark.parametrize("entrada, esperado", [
        ("SUL", "Sul"), ("sul", "Sul"), ("sudeste", "Sudeste"),
        ("NORDESTE", "Nordeste"), ("NORTE", "Norte"),
        ("Centro-oeste", "Centro-Oeste"), ("centro oeste", "Centro-Oeste"),
    ])
    def test_regiao(self, entrada, esperado):
        assert t.padronizar_regiao(entrada) == esperado

    def test_regiao_fora_do_dominio_nao_e_descartada(self):
        assert t.padronizar_regiao("leste") == "Leste"

    @pytest.mark.parametrize("entrada, esperado", [
        ("fechado ganho", "Fechado Ganho"), ("FECHADO GANHO", "Fechado Ganho"),
        ("closed won", "Fechado Ganho"), ("fechado perdido", "Fechado Perdido"),
        ("Closed Lost", "Fechado Perdido"),
    ])
    def test_status(self, entrada, esperado):
        assert t.padronizar_status(entrada) == esperado


class TestMudou:
    @pytest.mark.parametrize("a, b, esperado", [
        (None, float("nan"), False),
        ("", None, False),
        ("x", None, True),
        (None, "x", True),
        ("R$ 4320.00", 4320.0, True),
        ("Sul", "Sul", False),
        ("SUL", "Sul", True),
        (datetime(2024, 1, 15), pd.Timestamp("2024-01-15"), False),
        (4320.0, 4320.0, False),
    ])
    def test_mudou(self, a, b, esperado):
        assert t.mudou(a, b) is esperado


# =============================================================================
# 2. Pipeline com dados sintéticos
# =============================================================================
def linha(id_, cliente, regiao, rep, data, receita, produto="Hardware",
          status="Fechado Ganho", obs=None):
    return {"ID Pedido": id_, "Nome do Cliente": cliente, "Região": regiao,
            "Representante": rep, "Data do Pedido": data, "Receita": receita,
            "Produto": produto, "Status": status, "Observações": obs}


@pytest.fixture
def sintetico():
    d1, d2 = datetime(2024, 1, 10), datetime(2024, 2, 10)
    bruto = pd.DataFrame([
        linha(1, "ANA LIMA", "SUL", "BRUNO C", d1, "R$ 100.00"),
        linha(2, "ana lima", "sul", "bruno c", d1, "R$ 100.00", obs="duplicado"),   # dup confirmada
        linha(3, "Ana Lima", "Sul", "Bruno C", d2, "R$ 100.00"),                    # possível dup
        linha(4, "Beto Reis", None, "Bruno C", d1, "R$ 50.00"),                     # região via rep
        linha(5, "Caio Dias", "Norte", None, d1, "R$ 70.00", obs="sem representante"),
        linha(6, "Caio Dias", "Norte", "Sara J", d2, "R$ 80.00"),                    # rep via cliente
        linha(7, "Duda Melo", "Norte", "Sara J", d1, "R$ 90.00", status=None, obs="fechado perdido"),
        linha(8, "Eva Paz", "Norte", "Sara J", d1, None, obs="receita ausente"),
        linha(9, "Fábio Luz", "Norte", "Sara J", d1, "R$ 10.00", obs="possível duplicado"),  # sem par
        linha(10, "Gil Sá", "Sul", "Bruno C", d1, "R$ 20.00", status="closed won", obs="cliente recorrente"),
    ])
    tratados, log = t.tratar(bruto, "sintetico.xlsx")
    return bruto, tratados.set_index("id_pedido"), log


class TestPipeline:
    def test_colunas_snake_case_e_ordem(self, sintetico):
        _, df, _ = sintetico
        assert ["id_pedido", *df.columns] == t.ORDEM_SAIDA

    def test_nao_altera_dataframe_de_entrada(self, sintetico):
        bruto, _, _ = sintetico
        assert bruto.loc[0, "Nome do Cliente"] == "ANA LIMA"
        assert bruto.loc[0, "Receita"] == "R$ 100.00"

    def test_receita_numerica(self, sintetico):
        _, df, _ = sintetico
        assert pd.api.types.is_float_dtype(df["receita"])
        assert df.at[1, "receita"] == 100.0

    def test_receita_ausente_sinalizada(self, sintetico):
        _, df, _ = sintetico
        assert pd.isna(df.at[8, "receita"])
        assert df.at[8, "flag_receita_ausente"]
        assert "receita ausente" in df.at[8, "pendencias"]

    def test_status_movido_de_observacoes(self, sintetico):
        _, df, log = sintetico
        assert df.at[7, "status"] == "Fechado Perdido"
        assert pd.isna(df.at[7, "observacoes"])
        assert df.at[7, "observacoes_original"] == "fechado perdido"
        assert "status movido de Observações" in set(log["regra"])

    def test_status_em_ingles(self, sintetico):
        _, df, _ = sintetico
        assert df.at[10, "status"] == "Fechado Ganho"

    def test_observacao_livre_e_mantida(self, sintetico):
        _, df, _ = sintetico
        assert df.at[10, "observacoes"] == "cliente recorrente"

    def test_regiao_inferida_pelo_representante(self, sintetico):
        _, df, _ = sintetico
        assert df.at[4, "regiao"] == "Sul"
        assert df.at[4, "regiao_inferida"]

    def test_representante_inferido_pelo_cliente(self, sintetico):
        _, df, _ = sintetico
        assert df.at[5, "representante"] == "Sara J"
        assert df.at[5, "representante_inferido"]
        assert df.at[5, "flag_sem_representante"]

    def test_duplicata_confirmada(self, sintetico):
        _, df, _ = sintetico
        assert df.at[1, "status_registro"] == "ativo"
        assert df.at[2, "status_registro"] == "duplicado_confirmado"
        assert df.at[2, "duplicado_de"] == 1

    def test_possivel_duplicata_com_data_diferente(self, sintetico):
        _, df, _ = sintetico
        assert df.at[3, "status_registro"] == "possivel_duplicado"
        assert df.at[3, "duplicado_de"] == 1

    def test_duplicata_informada_sem_par(self, sintetico):
        _, df, _ = sintetico
        assert df.at[9, "status_registro"] == "possivel_duplicado"
        assert pd.isna(df.at[9, "duplicado_de"])

    def test_receita_ausente_nao_gera_duplicata(self, sintetico):
        _, df, _ = sintetico
        assert df.at[8, "status_registro"] == "ativo"

    def test_linhagem(self, sintetico):
        _, df, _ = sintetico
        assert (df["arquivo_origem"] == "sintetico.xlsx").all()
        assert df.at[1, "linha_origem"] == 2 and df.at[10, "linha_origem"] == 11

    def test_coluna_obrigatoria_ausente(self):
        bruto = pd.DataFrame([linha(1, "A", "Sul", "B", datetime(2024, 1, 1), "1")])
        with pytest.raises(ValueError, match="Receita"):
            t.tratar(bruto.drop(columns="Receita"), "x.xlsx")


# =============================================================================
# 3. Ponta a ponta sobre a planilha real
# =============================================================================
@pytest.fixture(scope="module")
def execucao(tmp_path_factory):
    saida = tmp_path_factory.mktemp("saida") / "tratada.xlsx"
    hash_antes = sha256(ORIGEM)
    env = {**os.environ, "PYTHONIOENCODING": "utf-8"}
    proc = subprocess.run([sys.executable, str(SCRIPT), "--saida", str(saida)],
                          capture_output=True, text=True, encoding="utf-8", env=env)
    assert proc.returncode == 0, proc.stderr
    return {
        "saida": saida,
        "hash_antes": hash_antes,
        "origem": pd.read_excel(ORIGEM, dtype=object),
        "tratados": pd.read_excel(saida, "pedidos_tratados").set_index("id_pedido"),
        "validos": pd.read_excel(saida, "pedidos_validos"),
        "log": pd.read_excel(saida, "log_alteracoes").fillna(""),
        "meta": pd.read_excel(saida, "metadados").set_index("chave")["valor"],
    }


class TestPontaAPonta:
    # --- estrutura e preservação da origem ----------------------------------
    def test_abas(self, execucao):
        wb = openpyxl.load_workbook(execucao["saida"], read_only=True)
        assert wb.sheetnames == ["pedidos_tratados", "pedidos_validos", "log_alteracoes",
                                 "resumo_qualidade", "validacoes", "metadados", "original"]

    def test_validacoes_aprovadas_na_planilha(self, execucao):
        validacoes = pd.read_excel(execucao["saida"], "validacoes")
        reprovadas = validacoes[~validacoes["aprovado"]]
        assert reprovadas.empty, reprovadas.to_string()

    def test_origem_nao_foi_alterada(self, execucao):
        assert sha256(ORIGEM) == execucao["hash_antes"]

    def test_hash_registrado_nos_metadados(self, execucao):
        assert execucao["meta"]["sha256_origem"] == execucao["hash_antes"]

    def test_aba_original_identica_a_origem(self, execucao):
        original = pd.read_excel(execucao["saida"], "original", dtype=object)
        pd.testing.assert_frame_equal(original, execucao["origem"])

    def test_contagem_de_linhas(self, execucao):
        assert len(execucao["tratados"]) == len(execucao["origem"]) == 30
        assert len(execucao["validos"]) == 29
        assert 1005 not in set(execucao["validos"]["id_pedido"])

    def test_nao_sobrescreve_origem(self):
        env = {**os.environ, "PYTHONIOENCODING": "utf-8"}
        proc = subprocess.run([sys.executable, str(SCRIPT), "--saida", str(ORIGEM)],
                              capture_output=True, text=True, encoding="utf-8", env=env)
        assert proc.returncode != 0
        assert "não pode sobrescrever" in proc.stderr

    # --- item 1: receita ----------------------------------------------------
    def test_receita_numerica(self, execucao):
        assert pd.api.types.is_numeric_dtype(execucao["tratados"]["receita"])

    def test_receitas_ausentes(self, execucao):
        df = execucao["tratados"]
        assert set(df.index[df["receita"].isna()]) == {1009, 1021}
        assert df.loc[[1009, 1021], "flag_receita_ausente"].all()

    def test_receita_fechado_ganho(self, execucao):
        v = execucao["validos"]
        assert v.loc[v["status"] == "Fechado Ganho", "receita"].sum() == pytest.approx(113_720)

    # --- item 2: datas ------------------------------------------------------
    def test_datas_validas(self, execucao):
        datas = execucao["tratados"]["data_pedido"]
        assert pd.api.types.is_datetime64_any_dtype(datas)
        assert datas.notna().all()
        assert (datas == datas.dt.normalize()).all()

    def test_datas_preservam_valor_da_origem(self, execucao):
        origem = execucao["origem"].set_index("ID Pedido")["Data do Pedido"]
        tratadas = execucao["tratados"]["data_pedido"]
        for id_, valor in origem.items():
            assert tratadas[int(id_)] == pd.Timestamp(valor).normalize()

    # --- itens 3 e 4: domínios e caixa -------------------------------------
    def test_dominio_regiao(self, execucao):
        assert set(execucao["tratados"]["regiao"]) <= set(t.REGIOES.values())

    def test_dominio_status(self, execucao):
        assert set(execucao["tratados"]["status"]) == {"Fechado Ganho", "Fechado Perdido"}

    def test_representantes(self, execucao):
        assert set(execucao["tratados"]["representante"]) == {
            "Bruno Carvalho", "Raquel Verde", "Sara Jorge", "Tomás Bastos"}

    @pytest.mark.parametrize("coluna", ["cliente", "representante"])
    def test_nomes_padronizados(self, execucao, coluna):
        for nome in execucao["tratados"][coluna]:
            assert nome == t.nome_proprio(nome), nome
            assert not nome.isupper() and not nome.islower()

    def test_status_1007_movido(self, execucao):
        df = execucao["tratados"]
        assert df.at[1007, "status"] == "Fechado Perdido"
        assert df.at[1007, "observacoes_original"] == "fechado perdido"
        assert pd.isna(df.at[1007, "observacoes"])

    def test_status_1019_traduzido(self, execucao):
        assert execucao["tratados"].at[1019, "status"] == "Fechado Ganho"

    # --- item 5: campos ausentes -------------------------------------------
    def test_sem_campos_obrigatorios_vazios(self, execucao):
        df = execucao["tratados"]
        assert df[["cliente", "regiao", "representante", "status", "produto"]].notna().all().all()

    @pytest.mark.parametrize("id_, coluna, valor, flag", [
        (1006, "regiao", "Centro-Oeste", "regiao_inferida"),
        (1016, "regiao", "Sul", "regiao_inferida"),
        (1027, "representante", "Sara Jorge", "representante_inferido"),
    ])
    def test_inferencias(self, execucao, id_, coluna, valor, flag):
        df = execucao["tratados"]
        assert df.at[id_, coluna] == valor
        assert df.at[id_, flag]
        assert "inferid" in df.at[id_, "pendencias"]

    def test_apenas_tres_inferencias(self, execucao):
        df = execucao["tratados"]
        assert df["regiao_inferida"].sum() + df["representante_inferido"].sum() == 3

    # --- item 6: duplicatas -------------------------------------------------
    def test_duplicatas(self, execucao):
        df = execucao["tratados"]
        nao_ativos = df[df["status_registro"] != "ativo"]
        esperado = {
            1005: ("duplicado_confirmado", 1001),
            1014: ("possivel_duplicado", 1008),
            1017: ("possivel_duplicado", 1004),
            1025: ("possivel_duplicado", 1007),
            1029: ("possivel_duplicado", 1006),
        }
        obtido = {i: (r.status_registro, int(r.duplicado_de)) for i, r in nao_ativos.iterrows()}
        assert obtido == esperado

    # --- item 7: observações -> flags --------------------------------------
    def test_flags_de_observacao(self, execucao):
        df = execucao["tratados"]
        assert set(df.index[df["flag_duplicado_informado"]]) == {1005, 1014, 1029}
        assert set(df.index[df["flag_sem_representante"]]) == {1027}
        assert set(df["observacoes"].dropna()) == {"cliente recorrente"}

    # --- item 8: colunas ----------------------------------------------------
    def test_colunas(self, execucao):
        assert ["id_pedido", *execucao["tratados"].columns] == t.ORDEM_SAIDA

    # --- linhagem: log x diferença real entre origem e tratado -------------
    def test_toda_alteracao_esta_no_log(self, execucao):
        """Cada célula que difere da origem precisa ter registro no log."""
        origem = execucao["origem"].rename(columns=t.COLUNAS).set_index("id_pedido")
        origem.index = origem.index.astype(int)
        df, log = execucao["tratados"], execucao["log"]
        registrados = set(zip(log["id_pedido"], log["coluna"]))
        faltando = [
            (id_, col)
            for col in t.COLUNAS.values() if col != "id_pedido"
            for id_ in origem.index
            if t.mudou(origem.at[id_, col], df.at[id_, col]) and (id_, col) not in registrados
        ]
        assert not faltando, f"alterações sem log: {faltando}"

    def test_log_reflete_valor_original(self, execucao):
        """valor_original do log deve ser exatamente o valor da origem."""
        origem = execucao["origem"].rename(columns=t.COLUNAS).set_index("id_pedido")
        origem.index = origem.index.astype(int)
        log = execucao["log"]
        log = log[log["coluna"].isin(origem.columns)]
        for r in log.itertuples():
            valor = origem.at[r.id_pedido, r.coluna]
            assert str(r.valor_original) == ("" if t.vazio(valor) else str(valor)), r

    def test_linha_origem(self, execucao):
        origem = execucao["origem"]
        df = execucao["tratados"]
        for pos, id_ in enumerate(origem["ID Pedido"]):
            assert df.at[int(id_), "linha_origem"] == pos + 2

    def test_metadados(self, execucao):
        m = execucao["meta"]
        assert int(m["linhas_origem"]) == 30
        assert int(m["linhas_validas"]) == 29
        assert int(m["duplicados_confirmados"]) == 1
        assert int(m["possiveis_duplicados"]) == 4
        assert int(m["receitas_ausentes"]) == 2
        assert int(m["celulas_alteradas"]) == len(execucao["log"])

    # --- formatação no Excel ------------------------------------------------
    @pytest.mark.parametrize("aba", ["pedidos_tratados", "pedidos_validos"])
    def test_formatos_de_celula(self, execucao, aba):
        ws = openpyxl.load_workbook(execucao["saida"])[aba]
        cab = {c.value: c.column_letter for c in ws[1]}
        for celula in ws[cab["data_pedido"]][1:]:
            assert celula.number_format == "yyyy-mm-dd"
        for celula in ws[cab["receita"]][1:]:
            assert celula.number_format == '"R$" #,##0.00'
        assert ws.freeze_panes == "A2"

    def test_idempotente(self, execucao, tmp_path):
        """Uma segunda execução produz os mesmos dados."""
        saida2 = tmp_path / "tratada2.xlsx"
        env = {**os.environ, "PYTHONIOENCODING": "utf-8"}
        subprocess.run([sys.executable, str(SCRIPT), "--saida", str(saida2)],
                       check=True, capture_output=True, env=env)
        pd.testing.assert_frame_equal(
            pd.read_excel(saida2, "pedidos_tratados").set_index("id_pedido"),
            execucao["tratados"])
