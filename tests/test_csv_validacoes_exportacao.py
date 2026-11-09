"""
Testes da leitura de CSV, das validações automáticas, dos indicadores de
qualidade e da exportação (CSV/Excel) usados pelo app Streamlit.
"""
from __future__ import annotations

import io
import sys
from pathlib import Path

import openpyxl
import pandas as pd
import pytest

RAIZ = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(RAIZ / "scripts"))

import exportacao as ex  # noqa: E402
import tratar_pedidos_vendas as t  # noqa: E402
import validacoes as v  # noqa: E402

XLSX = RAIZ / "data" / "Planilha_pedidos_vendas.xlsx"
CSV = RAIZ / "data" / "exemplo_pedidos_vendas.csv"
CABECALHO = "ID Pedido;Nome do Cliente;Região;Representante;Data do Pedido;Receita;Produto;Status;Observações"


def processar(conteudo: bytes, nome: str):
    bruto = t.ler_arquivo(conteudo, nome)
    tratados, log = t.tratar(bruto, nome)
    return bruto, tratados, log


@pytest.fixture(scope="module")
def csv_real():
    return processar(CSV.read_bytes(), CSV.name)


@pytest.fixture(scope="module")
def xlsx_real():
    return processar(XLSX.read_bytes(), XLSX.name)


# =============================================================================
# Leitura de arquivos
# =============================================================================
class TestLeitura:
    LINHA = "1001;JOÃO SILVA;sul;Ana;15/01/2024;R$ 10,00;Hardware;fechado ganho;"

    @pytest.mark.parametrize("codificacao", ["utf-8", "utf-8-sig", "cp1252"])
    def test_codificacoes(self, codificacao):
        conteudo = f"{CABECALHO}\n{self.LINHA}\n".encode(codificacao)
        df = t.ler_arquivo(conteudo, "x.csv")
        assert list(df.columns) == CABECALHO.split(";")
        assert df.iloc[0]["Nome do Cliente"] == "JOÃO SILVA"

    @pytest.mark.parametrize("sep", [";", ",", "\t", "|"])
    def test_separadores(self, sep):
        cab = CABECALHO.replace(";", sep)
        linha = self.LINHA.replace(";", sep).replace("R$ 10,00", "R$ 10.00")
        df = t.ler_arquivo(f"{cab}\n{linha}\n".encode(), "x.csv")
        assert df.shape == (1, 9)

    def test_valores_mantidos_como_texto(self):
        df = t.ler_arquivo(f"{CABECALHO}\n{self.LINHA}\n".encode(), "x.csv")
        assert df.iloc[0]["ID Pedido"] == "1001"
        assert df.iloc[0]["Receita"] == "R$ 10,00"

    def test_linha_vazia_descartada_preservando_numero_da_linha(self):
        texto = f"{CABECALHO}\n{self.LINHA}\n;;;;;;;;\n{self.LINHA.replace('1001', '1002')}\n"
        bruto = t.ler_arquivo(texto.encode(), "x.csv")
        tratados, _ = t.tratar(bruto, "x.csv")
        assert list(tratados["linha_origem"]) == [2, 4]

    def test_xlsx(self):
        df = t.ler_arquivo(XLSX.read_bytes(), XLSX.name)
        assert df.shape == (30, 9)

    @pytest.mark.parametrize("conteudo, nome, mensagem", [
        (b"", "x.csv", "vazio"),
        (b"   \n ", "x.csv", "vazio"),
        (b"abc", "x.pdf", "não suportado"),
    ])
    def test_arquivo_invalido(self, conteudo, nome, mensagem):
        with pytest.raises(ValueError, match=mensagem):
            t.ler_arquivo(conteudo, nome)

    def test_sem_linhas_de_dados(self):
        bruto = t.ler_arquivo(f"{CABECALHO}\n".encode(), "x.csv")
        with pytest.raises(ValueError, match="não tem linhas"):
            t.tratar(bruto, "x.csv")


class TestCabecalhos:
    @pytest.mark.parametrize("cabecalho", [
        "id pedido;NOME DO CLIENTE;Regiao;representante;data do pedido;RECEITA;produto;status;Observacoes",
        "id_pedido;cliente;regiao;representante;data_pedido;receita;produto;status;observacoes",
        " ID Pedido ; Nome do Cliente ;Região;Representante;Data do Pedido;Receita;Produto;Status;Observações",
    ])
    def test_variacoes_reconhecidas(self, cabecalho):
        texto = f"{cabecalho}\n1;Ana;Sul;Bia;01/01/2024;10;X;Fechado Ganho;\n"
        tratados, _ = t.tratar(t.ler_arquivo(texto.encode(), "x.csv"), "x.csv")
        assert tratados.at[0, "regiao"] == "Sul"

    def test_coluna_extra_ignorada(self):
        texto = f"{CABECALHO};Extra\n1;Ana;Sul;Bia;01/01/2024;10;X;Fechado Ganho;;lixo\n"
        tratados, _ = t.tratar(t.ler_arquivo(texto.encode(), "x.csv"), "x.csv")
        assert "Extra" not in tratados.columns

    def test_coluna_faltando(self):
        cab = CABECALHO.replace(";Receita", "")
        bruto = t.ler_arquivo(f"{cab}\n1;Ana;Sul;Bia;01/01/2024;X;Fechado Ganho;\n".encode(), "x.csv")
        with pytest.raises(ValueError, match="Receita"):
            t.tratar(bruto, "x.csv")


class TestIds:
    def test_ids_numericos(self):
        assert str(t.converter_ids(pd.Series(["1", "2", None])).dtype) == "Int64"

    def test_ids_alfanumericos_mantidos(self):
        ids = t.converter_ids(pd.Series(["PED-1", " PED-2 ", None]))
        assert list(ids[:2]) == ["PED-1", "PED-2"] and pd.isna(ids[2])

    def test_ids_decimais_mantidos_como_texto(self):
        assert list(t.converter_ids(pd.Series(["1.5", "2"]))) == ["1.5", "2"]


# =============================================================================
# CSV real x planilha real
# =============================================================================
class TestCsvEquivalenteAoXlsx:
    DADOS = ["id_pedido", "data_pedido", "cliente", "regiao", "representante", "produto",
             "receita", "status", "observacoes", "status_registro", "duplicado_de", "pendencias"]

    def test_mesmo_resultado(self, csv_real, xlsx_real):
        _, de_csv, _ = csv_real
        _, de_xlsx, _ = xlsx_real
        pd.testing.assert_frame_equal(de_csv[self.DADOS], de_xlsx[self.DADOS], check_dtype=False)

    def test_csv_registra_formatos_de_data(self, csv_real):
        _, _, log = csv_real
        datas = log[log["coluna"] == "data_pedido"]
        assert len(datas) == 30
        assert (datas["regra"] == t.REGRA_TIPO_DATA).sum() == 7      # já estavam em AAAA-MM-DD
        assert set(datas["regra"]) == {t.REGRA_TIPO_DATA, "data em formato fora do padrão convertida"}


# =============================================================================
# Validações
# =============================================================================
class TestValidar:
    @pytest.mark.parametrize("fixture", ["csv_real", "xlsx_real"])
    def test_tudo_aprovado(self, fixture, request):
        bruto, tratados, log = request.getfixturevalue(fixture)
        resultado = v.validar(bruto, tratados, log)
        assert len(resultado) == 14
        assert resultado["aprovado"].all(), resultado[~resultado["aprovado"]].to_string()

    def reprovadas(self, bruto, tratados, log) -> set[str]:
        r = v.validar(bruto, tratados, log)
        return set(r.loc[~r["aprovado"], "verificacao"])

    # Cada teste estraga o resultado de propósito e confere se a validação certa reprova.
    def test_detecta_alteracao_sem_log(self, xlsx_real):
        bruto, tratados, log = xlsx_real
        log2 = log[~((log["id_pedido"] == 1002) & (log["coluna"] == "cliente"))]
        assert "Toda alteração registrada no log" in self.reprovadas(bruto, tratados, log2)

    def test_detecta_log_adulterado(self, xlsx_real):
        bruto, tratados, log = xlsx_real
        log2 = log.copy()
        log2.loc[log2.index[0], "valor_original"] = "valor inventado"
        assert "Log fiel ao valor original" in self.reprovadas(bruto, tratados, log2)

    def test_detecta_linha_perdida(self, xlsx_real):
        bruto, tratados, log = xlsx_real
        assert "Nenhuma linha perdida ou criada" in self.reprovadas(bruto, tratados.iloc[1:], log)

    def test_detecta_id_repetido(self, xlsx_real):
        bruto, tratados, log = xlsx_real
        t2 = tratados.copy()
        t2.loc[1, "id_pedido"] = t2.loc[0, "id_pedido"]
        assert "ID do pedido preenchido e único" in self.reprovadas(bruto, t2, log)

    def test_detecta_receita_ausente_sem_flag(self, xlsx_real):
        bruto, tratados, log = xlsx_real
        t2 = tratados.copy()
        t2["flag_receita_ausente"] = False
        assert "Receita ausente sinalizada" in self.reprovadas(bruto, t2, log)

    def test_detecta_receita_texto(self, xlsx_real):
        bruto, tratados, log = xlsx_real
        t2 = tratados.copy()
        t2["receita"] = t2["receita"].astype(str)
        assert "Receita em formato numérico" in self.reprovadas(bruto, t2, log)

    def test_detecta_regiao_fora_do_dominio(self, xlsx_real):
        bruto, tratados, log = xlsx_real
        t2 = tratados.copy()
        t2.loc[0, "regiao"] = "SUL"
        assert "Região dentro do padrão (ou sinalizada)" in self.reprovadas(bruto, t2, log)

    def test_detecta_nome_sem_padrao(self, xlsx_real):
        bruto, tratados, log = xlsx_real
        t2 = tratados.copy()
        t2.loc[0, "cliente"] = "JOÃO SILVA"
        assert "Nomes com caixa padronizada" in self.reprovadas(bruto, t2, log)

    def test_detecta_inferencia_sem_flag(self, xlsx_real):
        bruto, tratados, log = xlsx_real
        t2 = tratados.copy()
        t2["regiao_inferida"] = False
        assert "Valores inferidos sinalizados" in self.reprovadas(bruto, t2, log)

    def test_detecta_duplicata_nao_removida(self, xlsx_real):
        bruto, tratados, log = xlsx_real
        t2 = tratados.copy()
        t2["status_registro"] = "ativo"
        assert "Sem duplicata exata entre os pedidos válidos" in self.reprovadas(bruto, t2, log)

    def test_detecta_referencia_orfa(self, xlsx_real):
        bruto, tratados, log = xlsx_real
        t2 = tratados.copy()
        t2["duplicado_de"] = t2["duplicado_de"].astype("object")
        t2.loc[t2["duplicado_de"].notna(), "duplicado_de"] = 9999
        assert "Referências de duplicata válidas" in self.reprovadas(bruto, t2, log)

    def test_detecta_data_com_hora(self, xlsx_real):
        bruto, tratados, log = xlsx_real
        t2 = tratados.copy()
        t2.loc[0, "data_pedido"] = pd.Timestamp("2024-01-15 10:30")
        assert "Datas válidas e sem hora" in self.reprovadas(bruto, t2, log)


# =============================================================================
# Indicadores e resumo
# =============================================================================
class TestQualidade:
    def test_resumo_xlsx(self, xlsx_real):
        _, tratados, log = xlsx_real
        r = v.resumo_qualidade(tratados, log).set_index("categoria")
        assert r.at["Receita em formato de texto", "pedidos_afetados"] == 28
        assert r.at["Receita ausente", "pedidos_pendentes"] == 2
        assert r.at["Duplicidade", "pedidos_afetados"] == 5
        assert r.at["Duplicidade", "pedidos_pendentes"] == 4
        assert r.at["Campo obrigatório vazio", "celulas_corrigidas"] == 3
        assert r.at["Status na coluna errada", "pedidos_afetados"] == 1
        assert (r["tipo"] == "Erro").all()            # xlsx não tem conversão de tipo pura

    def test_resumo_csv_tem_melhorias(self, csv_real):
        _, tratados, log = csv_real
        r = v.resumo_qualidade(tratados, log).set_index("categoria")
        assert r.at["Conversão de tipo", "tipo"] == "Melhoria"
        assert r.at["Conversão de tipo", "celulas_corrigidas"] == 7

    def test_celulas_do_resumo_somam_o_log(self, csv_real):
        _, tratados, log = csv_real
        assert v.resumo_qualidade(tratados, log)["celulas_corrigidas"].sum() == len(log)

    def test_indicadores(self, xlsx_real):
        _, tratados, log = xlsx_real
        ind = v.indicadores(tratados, log)
        assert ind["pedidos"] == 30
        assert ind["melhorias"] == len(log) == 78
        assert ind["pendencias"] == 9 and ind["pedidos_pendentes"] == 9
        assert ind["duplicados_confirmados"] == 1
        assert ind["pct_ok_antes"] == 0.0 and ind["pct_ok_depois"] == pytest.approx(0.7)
        assert ind["erros"] == ind["corrigidos_automaticamente"] + ind["pendencias"]

    def test_arquivo_limpo_nao_tem_erros(self):
        texto = f"{CABECALHO}\n1;Ana Lima;Sul;Bia Reis;2024-01-01;10;X;Fechado Ganho;\n"
        _, tratados, log = processar(texto.encode(), "x.csv")
        ind = v.indicadores(tratados, log)
        assert ind["pedidos_com_erro"] == 0 and ind["pendencias"] == 0
        assert ind["pct_ok_antes"] == ind["pct_ok_depois"] == 1.0
        assert (v.resumo_qualidade(tratados, log)["tipo"] == "Melhoria").all()


# =============================================================================
# Exportação
# =============================================================================
class TestExportacao:
    def test_csv_brasil(self, xlsx_real):
        _, tratados, _ = xlsx_real
        conteudo = ex.gerar_csv(ex.filtrar_validos(tratados), ";")
        assert conteudo.startswith(b"\xef\xbb\xbf")                 # BOM: acentos no Excel
        texto = conteudo.decode("utf-8-sig")
        assert texto.splitlines()[0].startswith("id_pedido;data_pedido;cliente")
        assert "1001;2024-01-15;João Silva;Sudeste;Sara Jorge;Licença de Software;4320,0;" in texto

    def test_csv_internacional(self, xlsx_real):
        _, tratados, _ = xlsx_real
        texto = ex.gerar_csv(tratados, ",").decode("utf-8-sig")
        assert "1001,2024-01-15,João Silva,Sudeste,Sara Jorge,Licença de Software,4320.0," in texto

    @pytest.mark.parametrize("sep", [";", ","])
    def test_csv_exportado_e_relido_sem_novos_erros(self, xlsx_real, sep):
        """A base limpa, reprocessada, não deve gerar nenhuma nova correção de erro."""
        _, tratados, _ = xlsx_real
        validos = ex.filtrar_validos(tratados)
        bruto, de_novo, log = processar(ex.gerar_csv(validos, sep), "limpo.csv")
        erros = [r for r in log["regra"] if v.categoria_log(r)[1]
                 and v.categoria_log(r)[0] not in ("Duplicidade", "Aviso digitado em Observações")]
        assert erros == []
        assert len(de_novo) == len(validos)
        assert v.validar(bruto, de_novo, log)["aprovado"].all()

    def test_excel_em_memoria(self, csv_real):
        bruto, tratados, log = csv_real
        abas = ex.montar_abas(bruto, tratados, log, CSV.name, "hash")
        wb = openpyxl.load_workbook(io.BytesIO(ex.gerar_excel(abas, bruto)))
        assert wb.sheetnames == ["pedidos_tratados", "pedidos_validos", "log_alteracoes",
                                 "resumo_qualidade", "validacoes", "metadados", "original"]
        original = pd.read_excel(io.BytesIO(ex.gerar_excel(abas, bruto)), "original", dtype=str)
        assert list(original["Receita"].head(2)) == ["R$ 4320.00", "R$ 1150.00"]
        assert list(original["Data do Pedido"].head(3)) == ["15/01/2024", "2024-01-22", "5 Jan 2024"]
        ws = wb["pedidos_tratados"]
        cab = {c.value: c.column_letter for c in ws[1]}
        assert ws[f"{cab['receita']}2"].number_format == '"R$" #,##0.00'
        assert ws[f"{cab['data_pedido']}2"].number_format == "yyyy-mm-dd"
