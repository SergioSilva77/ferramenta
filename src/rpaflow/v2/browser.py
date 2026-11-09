"""Módulo Browser v2 para rpaflow (Playwright).

Diferenças em relação ao `rpaflow.browser.Browser` (v1), as principais:

- `BrowserError` agora é importado (no v1 qualquer falha no `start()` virava `NameError`).
- Suporte nativo a Chrome via CDP (`cdp_port`) com perfil próprio, e `close()` que encerra SÓ o que é desta
  sessão — nada de `taskkill /IM chrome.exe` derrubando o navegador dos outros robôs.
- `wait_for_download()` funciona: o v1 abria o `expect_download` e saía do bloco sem esperar nada, então o evento
  de download (disparado ANTES) era sempre perdido. Aqui os downloads são capturados desde o início.
- `click_and_download()` pro caso clássico (clicar e baixar).
- `wait_for_element(state="hidden"/"detached", recursive=True)` não ficava esperando pra sempre.
- `close_tab()` sem a gambiarra que fechava a aba errada por acidente; abas abertas por popup também entram no
  controle de downloads.
- `close()` idempotente, fecha o contexto e zera o estado. Selenium saiu (nenhum método funcionava com ele).
"""

from __future__ import annotations

import os
import subprocess
import sys
import tempfile
import time
import urllib.request
from typing import Optional

from rpaflow._exceptions import BrowserError


class BrowserTimeout(BrowserError, TimeoutError):
    """Elemento/evento não apareceu a tempo."""


_CHROME_CANDIDATOS = (
    r"C:\Program Files\Google\Chrome\Application\chrome.exe",
    r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
    os.path.expandvars(r"%LOCALAPPDATA%\Google\Chrome\Application\chrome.exe"),
    "/usr/bin/google-chrome",
    "/usr/bin/chromium",
    "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
)


def _porta_responde(port: int) -> bool:
    try:
        with urllib.request.urlopen(f"http://127.0.0.1:{port}/json/version", timeout=1.5):
            return True
    except Exception:
        return False


def encerrar_chromes_do_perfil(user_data_dir: str) -> None:
    """Mata só os chrome.exe que usam EXATAMENTE este --user-data-dir (Windows).

    O casamento é exato: depois do caminho só pode vir aspas, " --outra-flag" ou o fim da linha.
    "User Data - RPA" é prefixo de "User Data - RPA Fiat" e o caminho tem espaços, então um
    `-like '*dir*'` pegaria o perfil dos outros robôs.
    """
    if sys.platform != "win32":
        return
    alvo = ("--user-data-dir=" + user_data_dir).replace("'", "''")
    ps_cmd = (
        f"$alvo = [regex]::Escape('{alvo}') + '(\"|\\s--|$)'; "
        "Get-CimInstance Win32_Process -Filter \"Name='chrome.exe'\" | "
        "Where-Object { $_.CommandLine -match $alvo } | "
        "ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }"
    )
    try:
        subprocess.run(
            ["powershell", "-NoProfile", "-Command", ps_cmd],
            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=15,
        )
    except Exception:
        pass


class Browser:
    """Automação de navegador com Playwright (síncrono).

    Três jeitos de iniciar:
        b.start("https://site")                                  # Chromium do Playwright
        b.start(user_data_dir="C:/perfis/robo1")                 # Chromium com perfil persistente
        b.start(cdp_port=9555, user_data_dir="C:/perfis/robo1")  # Chrome instalado, via CDP, porta própria

    Em paralelo, cada robô deve usar sua própria porta e seu próprio perfil.
    """

    def __init__(self):
        self._page = None
        self._playwright = None
        self._browser = None
        self._context = None
        self._type = None
        self._process: Optional[subprocess.Popen] = None
        self._user_data_dir = ""
        self._cdp_port = 0
        self._downloads: list = []
        self._downloads_dir: Optional[str] = None

    @property
    def page(self):
        """Playwright Page da aba atual."""
        return self._page

    @property
    def context(self):
        """Playwright BrowserContext."""
        return self._context

    def __enter__(self) -> "Browser":
        return self

    def __exit__(self, *exc) -> bool:
        self.close()
        return False

    def _require_page(self):
        if self._page is None:
            raise BrowserError("Navegador não iniciado (ou sem aba aberta). Chame start() primeiro.")
        return self._page

    # ========== INICIAR / FECHAR ==========

    def start(
        self,
        url: str = "",
        type: str = "playwright",
        headless: bool = False,
        *,
        cdp_port: Optional[int] = None,
        user_data_dir: Optional[str] = None,
        chrome_path: Optional[str] = None,
        args: Optional[list] = None,
        init_scripts: Optional[list] = None,
        locale: str = "pt-BR",
        downloads_dir: Optional[str] = None,
    ) -> bool:
        """Inicia o navegador.

        Args:
            url: se informado, já navega.
            type: só "playwright" (o v1 aceitava "selenium", mas nenhum método funcionava com ele).
            headless: sem janela (não vale para cdp_port).
            cdp_port: abre o Chrome instalado com --remote-debugging-port e conecta nele.
            user_data_dir: pasta do perfil (cookies/login persistem). Obrigatório pensar nele com cdp_port.
            chrome_path: caminho do chrome.exe (default: procura nos locais padrão).
            args: flags extras do Chromium.
            init_scripts: JS injetado em toda página.
            downloads_dir: onde salvar downloads automaticamente (default: temp do Playwright).
        """
        if type.lower() != "playwright":
            raise BrowserError(f"Tipo '{type}' não suportado. Só 'playwright'.")
        if self._playwright or self._process:
            raise BrowserError("Navegador já iniciado. Chame close() antes de start() de novo.")

        self._type = "playwright"
        self._downloads_dir = downloads_dir
        try:
            from playwright.sync_api import sync_playwright

            if cdp_port:
                # Primeiro sobe o Chrome (e valida porta/perfil), depois o Playwright: erro de porta em uso
                # aparece antes de qualquer outra coisa.
                self._start_cdp(cdp_port, user_data_dir, chrome_path, args, sync_playwright)
            elif user_data_dir:
                self._playwright = sync_playwright().start()
                os.makedirs(user_data_dir, exist_ok=True)
                self._context = self._playwright.chromium.launch_persistent_context(
                    user_data_dir,
                    headless=headless,
                    args=self._flags(headless, args),
                    no_viewport=True,
                    locale=locale,
                    accept_downloads=True,
                )
                self._browser = self._context.browser
                self._user_data_dir = user_data_dir
            else:
                self._playwright = sync_playwright().start()
                self._browser = self._playwright.chromium.launch(headless=headless, args=self._flags(headless, args))
                self._context = self._browser.new_context(no_viewport=True, locale=locale, accept_downloads=True)

            self._context.on("page", self._observar_pagina)
            for p in self._context.pages:
                self._observar_pagina(p)
            self._page = self._context.pages[0] if self._context.pages else self._context.new_page()

            for script in init_scripts or []:
                self._context.add_init_script(script)

            if url:
                self.navigate(url)
            return True
        except Exception as e:
            self.close()
            if isinstance(e, BrowserError):
                raise
            raise BrowserError(f"Erro ao iniciar navegador: {e}") from e

    @staticmethod
    def _flags(headless: bool, extra: Optional[list]) -> list:
        flags = [] if headless else ["--start-maximized"]
        return flags + list(extra or [])

    def _start_cdp(self, port: int, user_data_dir: Optional[str], chrome_path: Optional[str],
                   extra: Optional[list], sync_playwright):
        chrome_path = chrome_path or next((c for c in _CHROME_CANDIDATOS if os.path.exists(c)), None)
        if not chrome_path or not os.path.exists(chrome_path):
            raise BrowserError("Executável do Google Chrome não encontrado. Informe chrome_path=.")

        user_data_dir = user_data_dir or os.path.join(tempfile.gettempdir(), f"rpaflow_chrome_{port}")
        os.makedirs(user_data_dir, exist_ok=True)
        self._user_data_dir = user_data_dir
        self._cdp_port = port

        # Sobra de execução anterior que travou o perfil? Limpa só o que é deste perfil.
        encerrar_chromes_do_perfil(user_data_dir)
        time.sleep(1.0)

        if _porta_responde(port):
            raise BrowserError(f"Porta CDP {port} já está em uso por outro processo (outro robô?). Use outra porta.")

        self._process = subprocess.Popen([
            chrome_path,
            f"--remote-debugging-port={port}",
            f"--user-data-dir={user_data_dir}",
            "--start-maximized",
            "--no-first-run",
            "--no-default-browser-check",
            *(extra or []),
        ])

        limite = time.time() + 30
        while not _porta_responde(port):
            if time.time() > limite:
                raise BrowserError(f"Chrome não abriu a porta CDP {port} em 30s.")
            if self._process.poll() is not None:
                raise BrowserError("O Chrome encerrou logo ao abrir (perfil em uso por outro Chrome?).")
            time.sleep(0.5)

        self._playwright = sync_playwright().start()
        self._browser = self._playwright.chromium.connect_over_cdp(f"http://127.0.0.1:{port}")
        self._context = self._browser.contexts[0] if self._browser.contexts else self._browser.new_context(
            no_viewport=True, locale="pt-BR", accept_downloads=True
        )

    def close(self) -> None:
        """Fecha tudo que ESTA sessão abriu (e só isso). Pode chamar várias vezes."""
        for fechar in (
            lambda: self._context and self._context.close(),
            lambda: self._browser and self._browser.close(),
            lambda: self._playwright and self._playwright.stop(),
        ):
            try:
                fechar()
            except Exception:
                pass

        self._page = None
        self._context = None
        self._browser = None
        self._playwright = None
        self._downloads.clear()

        if self._process:
            try:
                self._process.terminate()
                self._process.wait(timeout=3)
            except Exception:
                pass
            self._process = None
        if self._cdp_port and self._user_data_dir:
            encerrar_chromes_do_perfil(self._user_data_dir)
        self._cdp_port = 0

    # ========== NAVEGAÇÃO ==========

    def navigate(self, url: str, wait_until: str = "load", timeout: int = 30000) -> bool:
        """Navega para uma URL. wait_until: 'load', 'domcontentloaded', 'networkidle', 'commit'."""
        self._require_page().goto(url, wait_until=wait_until, timeout=timeout)
        return True

    def wait_for_load_state(self, state: str = "load", timeout: int = 30000) -> bool:
        """Espera carregamento da página. States: 'load', 'domcontentloaded', 'networkidle'."""
        self._require_page().wait_for_load_state(state, timeout=timeout)
        return True

    # ========== CLIQUE ==========

    def click(self, selector: str, frame: str = None, recursive: bool = False, timeout: int = 5000) -> bool:
        """Clica em um elemento."""
        if recursive:
            return self._find_and_click_recursive(selector, timeout)
        loc = self._get_locator(selector, frame)
        loc.click(timeout=timeout)
        return True

    def hover(self, selector: str, frame: str = None, recursive: bool = False, timeout: int = 5000) -> bool:
        """Passa o mouse sobre um elemento."""
        if recursive:
            return self._find_and_hover_recursive(selector, timeout)
        loc = self._get_locator(selector, frame)
        loc.hover(timeout=timeout)
        return True

    # ========== TEXTO ==========

    def type_text(self, selector: str, text: str, delay: int = 50, frame: str = None,
                  recursive: bool = False, timeout: int = 5000) -> bool:
        """Digita caractere por caractere (mais humano)."""
        loc = self._find_element_recursive(selector, timeout) if recursive else self._get_locator(selector, frame)
        loc.click(timeout=timeout)
        loc.fill("", timeout=timeout)
        # `type` é deprecated no Playwright; press_sequentially é o substituto.
        digitar = getattr(loc, "press_sequentially", None) or loc.type
        digitar(text, delay=delay, timeout=timeout)
        return True

    def fill_text(self, selector: str, text: str, frame: str = None, recursive: bool = False,
                  timeout: int = 5000) -> bool:
        """Preenche instantaneamente (substitui todo o conteúdo)."""
        loc = self._find_element_recursive(selector, timeout) if recursive else self._get_locator(selector, frame)
        loc.fill(text, timeout=timeout)
        return True

    def get_text(self, selector: str, frame: str = None, recursive: bool = False, timeout: int = 5000) -> str:
        """Retorna o texto visível de um elemento."""
        loc = self._find_element_recursive(selector, timeout) if recursive else self._get_locator(selector, frame)
        return loc.inner_text(timeout=timeout).strip()

    # ========== SELECT ==========

    def select_option(self, selector: str, label: str, frame: str = None, recursive: bool = False,
                      timeout: int = 5000) -> bool:
        """Seleciona opção em <select> pelo texto visível."""
        loc = self._find_element_recursive(selector, timeout) if recursive else self._get_locator(selector, frame)
        loc.select_option(label=label, timeout=timeout)
        return True

    def select_option_by_index(self, selector: str, index: int, frame: str = None, recursive: bool = False,
                               timeout: int = 5000) -> bool:
        """Seleciona opção em <select> pelo índice (1-based)."""
        loc = self._find_element_recursive(selector, timeout) if recursive else self._get_locator(selector, frame)
        loc.select_option(index=index - 1, timeout=timeout)
        return True

    # ========== ELEMENTOS ==========

    def get_all_elements(self, selector: str, frame: str = None, recursive: bool = False, timeout: int = 5000) -> list:
        """Retorna lista de elementos que matcham o seletor."""
        if recursive:
            return self._find_all_recursive(selector, timeout)
        return self._get_locator(selector, frame).all()

    def wait_for_element(self, selector: str, state: str = "visible", frame: str = None,
                         recursive: bool = False, timeout: int = 5000):
        """Espera elemento ficar num estado. States: 'visible', 'hidden', 'attached', 'detached'."""
        if recursive:
            if state in ("hidden", "detached"):
                # O v1 procurava um elemento VISÍVEL pra depois esperar ele sumir: nunca terminava bem.
                self._esperar_sumir_recursivo(selector, timeout)
                return None
            loc = self._find_element_recursive(selector, timeout, only_visible=(state == "visible"))
            loc.wait_for(state=state, timeout=timeout)
            return loc
        loc = self._get_locator(selector, frame)
        loc.wait_for(state=state, timeout=timeout)
        return loc

    def is_visible(self, selector: str, frame: str = None, recursive: bool = False) -> bool:
        """Retorna se o elemento está visível agora (sem esperar)."""
        if recursive:
            return any(self._visivel_no_frame(f, selector) for f in self._require_page().frames)
        try:
            return self._get_locator(selector, frame).first.is_visible()
        except Exception:
            return False

    # ========== TECLADO ==========

    def send_shortcut(self, key: str) -> bool:
        """Envia atalho de teclado. Ex: 'Enter', 'Control+A', 'Escape'."""
        self._require_page().keyboard.press(key)
        return True

    # ========== FRAME ==========

    def switch_to_frame(self, selector: str):
        """Entra no iframe. Retorna o frame_locator."""
        return self._require_page().frame_locator(selector)

    def switch_to_root(self):
        """Sai do iframe (volta ao contexto raiz)."""
        return self._require_page()

    # ========== DOWNLOAD ==========

    def _observar_pagina(self, page) -> None:
        """Registra os downloads de toda aba (inclusive as abertas por popup)."""
        page.on("download", self._ao_baixar)

    def _ao_baixar(self, download) -> None:
        self._downloads.append(download)

    def clear_downloads(self) -> None:
        """Descarta downloads já capturados (chame antes de disparar o clique, se quiser garantir)."""
        self._downloads.clear()

    def wait_for_download(self, save_path: str = None, timeout: int = 30000) -> Optional[str]:
        """Espera um download e retorna o caminho do arquivo.

        Pode ser chamado DEPOIS do clique: o download já capturado é devolvido na hora.
        Cada download é entregue uma vez.
        """
        self._require_page()
        limite = time.time() + timeout / 1000
        while not self._downloads:
            if time.time() > limite:
                raise BrowserTimeout(f"Nenhum download iniciou em {timeout}ms.")
            self._page.wait_for_timeout(200)  # bombeia os eventos do Playwright
        download = self._downloads.pop(0)
        return self._salvar_download(download, save_path, timeout)

    def click_and_download(self, selector: str, save_path: str = None, frame: str = None,
                           recursive: bool = False, timeout: int = 30000) -> Optional[str]:
        """Clica no elemento e devolve o arquivo baixado (forma mais segura de baixar)."""
        self._require_page()
        with self._page.expect_download(timeout=timeout) as info:
            self.click(selector, frame=frame, recursive=recursive, timeout=min(timeout, 10000))
        download = info.value
        if download in self._downloads:
            self._downloads.remove(download)
        return self._salvar_download(download, save_path, timeout)

    @staticmethod
    def _salvar_download(download, save_path: Optional[str], timeout: int) -> str:
        if save_path:
            pasta = os.path.dirname(os.path.abspath(save_path))
            os.makedirs(pasta, exist_ok=True)
            download.save_as(save_path)
            return save_path
        return str(download.path())

    # ========== JANELA ==========

    def maximize(self) -> bool:
        """Maximiza a janela do navegador."""
        page = self._require_page()
        try:
            cdp = self._context.new_cdp_session(page)
            janela = cdp.send("Browser.getWindowForTarget")
            cdp.send("Browser.setWindowBounds", {"windowId": janela["windowId"], "bounds": {"windowState": "maximized"}})
            return True
        except Exception:
            size = page.evaluate("() => ({ width: screen.width, height: screen.height })")
            page.set_viewport_size(size)
            return True

    # ========== ABAS ==========

    def get_tabs(self) -> list:
        """Lista todas as abas abertas. Retorna lista de dicts com index, title e url."""
        tabs = []
        for i, p in enumerate(self._context.pages):
            try:
                title = p.title()
            except Exception:
                title = ""
            tabs.append({"index": i, "title": title, "url": p.url})
        return tabs

    def get_current_tab(self) -> dict:
        """Retorna info da aba atual (index, title, url)."""
        page = self._require_page()
        pages = self._context.pages
        idx = pages.index(page) if page in pages else -1
        try:
            title = page.title()
        except Exception:
            title = ""
        return {"index": idx, "title": title, "url": page.url}

    def switch_to_tab(self, index: int) -> bool:
        """Muda o contexto para a aba pelo índice (0-based)."""
        pages = self._context.pages
        if index < 0 or index >= len(pages):
            raise IndexError(f"Índice {index} inválido. Existem {len(pages)} abas (0 a {len(pages) - 1}).")
        self._page = pages[index]
        self._page.bring_to_front()
        return True

    def new_tab(self, url: str = "") -> bool:
        """Abre uma nova aba. Se url for informado, navega até ela."""
        self._page = self._context.new_page()
        if url:
            self._page.goto(url)
        self._page.bring_to_front()
        return True

    def close_tab(self, index: int = None) -> bool:
        """Fecha a aba pelo índice (0-based). Se index=None, fecha a aba atual.

        Fechando a aba atual, a seleção vai para a aba anterior (ou a primeira).
        """
        pages = self._context.pages
        if not pages:
            return False

        if index is None:
            index = pages.index(self._page) if self._page in pages else 0
        if index < 0 or index >= len(pages):
            raise IndexError(f"Índice {index} inválido. Existem {len(pages)} abas (0 a {len(pages) - 1}).")

        alvo = pages[index]
        era_atual = alvo == self._page
        alvo.close()

        restantes = self._context.pages
        if era_atual or self._page is None or self._page.is_closed():
            self._page = restantes[max(0, index - 1)] if restantes else None
            if self._page is not None:
                self._page.bring_to_front()
        return True

    # ========== SCREENSHOT ==========

    def screenshot(self, filepath: str) -> bool:
        """Captura screenshot."""
        pasta = os.path.dirname(os.path.abspath(filepath))
        os.makedirs(pasta, exist_ok=True)
        self._require_page().screenshot(path=filepath)
        return True

    # ========== HELPERS INTERNOS ==========

    def _get_locator(self, selector: str, frame: str = None):
        """Retorna o locator, com suporte a frame explícito."""
        page = self._require_page()
        if frame:
            return page.frame_locator(frame).locator(selector)
        return page.locator(selector)

    @staticmethod
    def _visivel_no_frame(frame, selector: str) -> bool:
        try:
            return frame.locator(selector).first.is_visible()
        except Exception:
            return False

    def _find_element_recursive(self, selector: str, timeout: int = 5000, only_visible: bool = True):
        """Busca elemento em todos os iframes (e na página principal)."""
        page = self._require_page()
        limite = time.time() + timeout / 1000
        while True:
            for frame in page.frames:
                try:
                    loc = frame.locator(selector).first
                    if (loc.is_visible() if only_visible else loc.count() > 0):
                        return loc
                except Exception:
                    continue
            if time.time() >= limite:
                break
            page.wait_for_timeout(300)
        raise BrowserTimeout(f"Elemento não encontrado após {timeout}ms: {selector}")

    def _esperar_sumir_recursivo(self, selector: str, timeout: int) -> None:
        page = self._require_page()
        limite = time.time() + timeout / 1000
        while True:
            if not any(self._visivel_no_frame(f, selector) for f in page.frames):
                return
            if time.time() >= limite:
                raise BrowserTimeout(f"Elemento continua visível após {timeout}ms: {selector}")
            page.wait_for_timeout(300)

    def _find_all_recursive(self, selector: str, timeout: int = 5000) -> list:
        """Busca múltiplos elementos em todos os iframes. Lista vazia se não achar."""
        page = self._require_page()
        limite = time.time() + timeout / 1000
        while True:
            for frame in page.frames:
                try:
                    loc = frame.locator(selector)
                    if loc.count() > 0:
                        return loc.all()
                except Exception:
                    continue
            if time.time() >= limite:
                return []
            page.wait_for_timeout(300)

    def _find_and_click_recursive(self, selector: str, timeout: int = 5000) -> bool:
        self._find_element_recursive(selector, timeout).click(timeout=timeout)
        return True

    def _find_and_hover_recursive(self, selector: str, timeout: int = 5000) -> bool:
        self._find_element_recursive(selector, timeout).hover(timeout=timeout)
        return True

    def _find_and_select_recursive(self, selector: str, label: str, timeout: int = 5000) -> bool:
        self._find_element_recursive(selector, timeout).select_option(label=label, timeout=timeout)
        return True

    # ========== MÉTODOS RECURSIVOS PÚBLICOS ==========

    def find_element_in_frames(self, selector: str, timeout: int = 5000):
        """Busca elemento em todos os iframes e retorna o Locator."""
        return self._find_element_recursive(selector, timeout)

    def click_in_frames(self, selector: str, timeout: int = 5000) -> bool:
        """Clica em elemento em qualquer iframe."""
        return self._find_and_click_recursive(selector, timeout)

    def select_in_frames(self, selector: str, label: str, timeout: int = 5000) -> bool:
        """Seleciona opção em qualquer iframe."""
        return self._find_and_select_recursive(selector, label, timeout)

    def find_all_in_frames(self, selector: str, timeout: int = 5000) -> list:
        """Busca múltiplos elementos em qualquer iframe."""
        return self._find_all_recursive(selector, timeout)
