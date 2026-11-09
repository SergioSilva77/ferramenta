"""Browser v2 — integração real com Chromium (Playwright). Servidor HTTP local, sem internet."""

import functools
import http.server
import os
import socket
import subprocess
import threading
import time

import pytest

pytest.importorskip("playwright")

from rpaflow._exceptions import BrowserError  # noqa: E402
from rpaflow.v2.browser import (  # noqa: E402
    Browser, BrowserTimeout, _CHROME_CANDIDATOS, _porta_responde, encerrar_chromes_do_perfil,
)


def _esperar(condicao, timeout=20):
    limite = time.time() + timeout
    while not condicao():
        assert time.time() < limite, "timeout esperando condição"
        time.sleep(0.3)

PAGINA = """<html><body>
<h1 id="titulo">Olá</h1>
<a id="baixar" href="/arquivo.txt" download>baixar</a>
<div id="aviso">carregando...</div>
<script>setTimeout(() => document.getElementById('aviso').style.display = 'none', 700);</script>
<iframe id="f1" src="/frame.html" width="400" height="200"></iframe>
</body></html>"""

FRAME = """<html><body>
<button id="dentro" onclick="document.getElementById('res').innerText='clicou'">Dentro do iframe</button>
<span id="res"></span>
</body></html>"""


class _Handler(http.server.BaseHTTPRequestHandler):
    def log_message(self, *a):
        pass

    def do_GET(self):
        if self.path == "/arquivo.txt":
            corpo = b"conteudo do arquivo"
            self.send_response(200)
            self.send_header("Content-Type", "application/octet-stream")
            self.send_header("Content-Disposition", 'attachment; filename="arquivo.txt"')
        else:
            corpo = (FRAME if self.path == "/frame.html" else PAGINA).encode("utf-8")
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
        self.send_header("Content-Length", str(len(corpo)))
        self.end_headers()
        self.wfile.write(corpo)


@pytest.fixture(scope="module")
def site():
    srv = http.server.ThreadingHTTPServer(("127.0.0.1", 0), _Handler)
    threading.Thread(target=srv.serve_forever, daemon=True).start()
    yield f"http://127.0.0.1:{srv.server_address[1]}"
    srv.shutdown()


@pytest.fixture
def b(site):
    with Browser() as br:
        br.start(site, headless=True)
        yield br


def _porta_livre() -> int:
    with socket.socket() as s:
        s.bind(("127.0.0.1", 0))
        return s.getsockname()[1]


def test_start_com_erro_levanta_browser_error_e_nao_name_error():
    br = Browser()
    with pytest.raises(BrowserError):
        br.start(type="selenium")  # v1 aceitava e nada funcionava
    with pytest.raises(BrowserError, match="start"):
        br.navigate("http://x")  # sem start()


def test_start_duas_vezes_e_close_idempotente(site):
    br = Browser()
    br.start(site, headless=True)
    with pytest.raises(BrowserError, match="close"):
        br.start(site, headless=True)
    br.close()
    br.close()
    with pytest.raises(BrowserError):
        br.get_text("#titulo")


def test_texto_clique_recursivo_em_iframe(b):
    assert b.get_text("#titulo") == "Olá"
    b.click_in_frames("#dentro")
    assert b.find_element_in_frames("#res").inner_text() == "clicou"
    assert b.is_visible("#res", recursive=True)
    assert not b.is_visible("#nao_existe")


def test_wait_hidden_recursivo_nao_trava(b):
    b.wait_for_element("#aviso", state="hidden", recursive=True, timeout=5000)  # v1: ficava sem saída
    assert not b.is_visible("#aviso")
    with pytest.raises(BrowserTimeout):
        b.wait_for_element("#titulo", state="hidden", recursive=True, timeout=600)
    with pytest.raises(BrowserTimeout):
        b.find_element_in_frames("#nao_existe", timeout=500)


def test_download_depois_do_clique(b, tmp_path):
    b.click("#baixar")  # o clique vem ANTES: no v1 o evento era perdido
    destino = str(tmp_path / "pasta_nova" / "salvo.txt")
    assert b.wait_for_download(save_path=destino, timeout=10000) == destino
    assert open(destino, encoding="utf-8").read() == "conteudo do arquivo"
    with pytest.raises(BrowserTimeout):
        b.wait_for_download(timeout=500)  # cada download é entregue uma vez


def test_click_and_download(b, tmp_path):
    caminho = b.click_and_download("#baixar", save_path=str(tmp_path / "x.txt"))
    assert open(caminho, encoding="utf-8").read() == "conteudo do arquivo"


def test_abas(b, site):
    b.new_tab(site)
    b.new_tab(site)
    assert len(b.get_tabs()) == 3
    assert b.get_current_tab()["index"] == 2
    b.close_tab()  # fecha a atual, volta pra anterior
    assert b.get_current_tab()["index"] == 1
    b.close_tab(0)
    assert len(b.get_tabs()) == 1
    assert b.get_text("#titulo") == "Olá"
    with pytest.raises(IndexError):
        b.switch_to_tab(5)


def test_screenshot_cria_pasta(b, tmp_path):
    destino = tmp_path / "prints" / "tela.png"
    b.screenshot(str(destino))
    assert destino.stat().st_size > 0


# ───────────────────────── CDP (Chrome instalado) ─────────────────────────
_CHROME = next((c for c in _CHROME_CANDIDATOS if os.path.exists(c)), None)


@pytest.mark.skipif(not _CHROME, reason="Chrome não instalado")
def test_cdp_perfil_e_porta_proprios_e_close_so_mata_o_proprio(site, tmp_path):
    p1, p2 = _porta_livre(), _porta_livre()
    # perfis com nomes onde um é PREFIXO do outro (o caso que derrubava o robô vizinho)
    perfil_a = str(tmp_path / "User Data - RPA")
    perfil_b = str(tmp_path / "User Data - RPA Outro")
    # O "vizinho" é outro robô (outro processo): sobe o Chrome dele na mão, com o perfil de nome parecido.
    os.makedirs(perfil_b)
    vizinho = subprocess.Popen([
        _CHROME, f"--remote-debugging-port={p2}", f"--user-data-dir={perfil_b}", "--headless=new",
        "--no-first-run", "about:blank",
    ])
    a = Browser()
    try:
        _esperar(lambda: _porta_responde(p2))
        a.start(site, cdp_port=p1, user_data_dir=perfil_a, args=["--headless=new"])
        assert a.get_text("#titulo") == "Olá"

        with pytest.raises(BrowserError, match="em uso"):  # porta ocupada por outro robô
            Browser().start(site, cdp_port=p1, user_data_dir=str(tmp_path / "terceiro"), args=["--headless=new"])
        assert a.get_text("#titulo") == "Olá"  # a falha do terceiro não mexeu no `a`

        a.close()
        assert not _porta_responde(p1)  # o Chrome do `a` morreu
        assert _porta_responde(p2)  # o do vizinho continua de pé
    finally:
        a.close()
        vizinho.terminate()
        vizinho.wait(timeout=10)
        encerrar_chromes_do_perfil(perfil_b)
    assert not _porta_responde(p2)


def test_perfil_persistente_mantem_cookie(site, tmp_path):
    perfil = str(tmp_path / "perfil")
    with Browser() as br:
        br.start(site, headless=True, user_data_dir=perfil)
        br.page.evaluate("document.cookie = 'sessao=abc; max-age=3600'")
    with Browser() as br:
        br.start(site, headless=True, user_data_dir=perfil)
        assert "sessao=abc" in br.page.evaluate("document.cookie")
