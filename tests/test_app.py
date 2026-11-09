"""Testes do app Streamlit (execução sem navegador via streamlit.testing)."""
from __future__ import annotations

from pathlib import Path

import pytest
from streamlit.testing.v1 import AppTest

APP = Path(__file__).resolve().parents[1] / "app.py"


@pytest.fixture
def app():
    return AppTest.from_file(str(APP), default_timeout=60).run()


@pytest.fixture
def app_exemplo(app):
    app.button(key="exemplo").click().run()
    assert not app.exception, app.exception
    return app


def metricas(at) -> dict:
    return {m.label: (m.value, m.delta) for m in at.metric}


def test_tela_inicial_pede_arquivo(app):
    assert not app.exception
    assert "Envie um arquivo CSV" in app.info[0].value
    assert not app.metric


def test_indicadores_do_exemplo(app_exemplo):
    m = metricas(app_exemplo)
    assert m["Pedidos analisados"][0] == "30"
    assert m["Erros encontrados"][0] == "93"
    assert m["Melhorias aplicadas"][0] == "108"
    assert m["Pendências para revisão"][0] == "9"
    assert m["Pedidos sem problemas"][0] == "70%"


def test_validacoes_aprovadas(app_exemplo):
    assert any("14/14" in s.value for s in app_exemplo.success)
    assert not app_exemplo.error


def test_abas(app_exemplo):
    rotulos = [tab.label for tab in app_exemplo.tabs]
    assert rotulos == ["📊 Erros e melhorias", "✅ Dados limpos", "⚠️ Pendências (9)",
                       "🔎 Validações", "📝 Log (108)", "📄 Original"]


def test_alternar_todos_os_pedidos(app_exemplo):
    app_exemplo.toggle[0].set_value(False).run()
    assert not app_exemplo.exception
    assert any("30 pedido(s)" in c.value for c in app_exemplo.caption)


def test_opcoes_de_download(app_exemplo):
    app_exemplo.radio(key="conteudo_csv").set_value("Todos os pedidos (com duplicatas marcadas)").run()
    app_exemplo.radio(key="formato_csv").set_value("Padrão internacional ( , e ponto decimal)").run()
    assert not app_exemplo.exception
