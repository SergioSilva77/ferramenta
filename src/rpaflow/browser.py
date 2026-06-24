"""Módulo Browser para rpaflow."""

import time
from typing import Optional


class Browser:
    """Classe para automação de navegador (Playwright ou Selenium)."""

    def __init__(self):
        self._page = None
        self._playwright = None
        self._browser = None
        self._type = None

    @property
    def page(self):
        """Retorna o Playwright Page subjacente."""
        return self._page

    # ========== INICIAR / FECHAR ==========

    def start(self, url: str = "", type: str = "playwright", headless: bool = False) -> bool:
        """Inicia o navegador."""
        try:
            self._type = type.lower()
            if self._type == "playwright":
                from playwright.sync_api import sync_playwright
                self._playwright = sync_playwright().start()
                self._browser = self._playwright.chromium.launch(headless=headless)
                self._page = self._browser.new_page()
            elif self._type == "selenium":
                from selenium import webdriver
                self._page = webdriver.Chrome()
            else:
                raise ValueError(f"Tipo '{type}' não suportado. Use 'playwright' ou 'selenium'.")

            if url:
                self.navigate(url)
            return True
        except Exception as e:
            raise BrowserError(f"Erro ao iniciar navegador: {e}")

    def close(self) -> None:
        """Fecha o navegador."""
        if self._page:
            try:
                self._page.close()
            except Exception:
                pass
        if self._browser:
            try:
                self._browser.close()
            except Exception:
                pass
        if self._playwright:
            try:
                self._playwright.stop()
            except Exception:
                pass

    # ========== NAVEGAÇÃO ==========

    def navigate(self, url: str) -> bool:
        """Navega para uma URL."""
        self._page.goto(url)
        return True

    def wait_for_load_state(self, state: str = "load", timeout: int = 30000) -> bool:
        """Espera carregamento da página. States: 'load', 'domcontentloaded', 'networkidle'."""
        self._page.wait_for_load_state(state, timeout=timeout)
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

    def type_text(self, selector: str, text: str, delay: int = 50, frame: str = None, recursive: bool = False, timeout: int = 5000) -> bool:
        """Digita caractere por caractere (mais humano)."""
        if recursive:
            loc = self._find_element_recursive(selector, timeout)
            loc.click()
            loc.fill("")
            loc.type(text, delay=delay, timeout=timeout)
            return True
        loc = self._get_locator(selector, frame)
        loc.click()
        loc.fill("")
        loc.type(text, delay=delay, timeout=timeout)
        return True

    def fill_text(self, selector: str, text: str, frame: str = None, recursive: bool = False, timeout: int = 5000) -> bool:
        """Preenche instantaneamente (substitui todo o conteúdo)."""
        if recursive:
            loc = self._find_element_recursive(selector, timeout)
            loc.fill(text, timeout=timeout)
            return True
        loc = self._get_locator(selector, frame)
        loc.fill(text, timeout=timeout)
        return True

    def get_text(self, selector: str, frame: str = None, recursive: bool = False, timeout: int = 5000) -> str:
        """Retorna o texto visível de um elemento."""
        if recursive:
            loc = self._find_element_recursive(selector, timeout)
            return loc.inner_text(timeout=timeout).strip()
        loc = self._get_locator(selector, frame)
        return loc.inner_text(timeout=timeout).strip()

    # ========== SELECT ==========

    def select_option(self, selector: str, label: str, frame: str = None, recursive: bool = False, timeout: int = 5000) -> bool:
        """Seleciona opção em <select> pelo texto visível."""
        if recursive:
            return self._find_and_select_recursive(selector, label, timeout)
        loc = self._get_locator(selector, frame)
        loc.select_option(label=label, timeout=timeout)
        return True

    def select_option_by_index(self, selector: str, index: int, frame: str = None, recursive: bool = False, timeout: int = 5000) -> bool:
        """Seleciona opção em <select> pelo índice (1-based)."""
        if recursive:
            loc = self._find_element_recursive(selector, timeout)
            loc.select_option(index=index - 1, timeout=timeout)
            return True
        loc = self._get_locator(selector, frame)
        loc.select_option(index=index - 1, timeout=timeout)
        return True

    # ========== ELEMENTOS ==========

    def get_all_elements(self, selector: str, frame: str = None, recursive: bool = False, timeout: int = 5000) -> list:
        """Retorna lista de elementos que matcham o seletor."""
        if recursive:
            return self._find_all_recursive(selector, timeout)
        loc = self._get_locator(selector, frame)
        return loc.all()

    def wait_for_element(self, selector: str, state: str = "visible", frame: str = None, recursive: bool = False, timeout: int = 5000):
        """Espera elemento ficar num estado. States: 'visible', 'hidden', 'attached', 'detached'."""
        if recursive:
            loc = self._find_element_recursive(selector, timeout)
            loc.wait_for(state=state, timeout=timeout)
            return loc
        loc = self._get_locator(selector, frame)
        loc.wait_for(state=state, timeout=timeout)
        return loc

    # ========== TECLADO ==========

    def send_shortcut(self, key: str) -> bool:
        """Envia atalho de teclado. Ex: 'Enter', 'Control+A', 'Escape'."""
        self._page.keyboard.press(key)
        return True

    # ========== FRAME ==========

    def switch_to_frame(self, selector: str):
        """Entra no iframe. Retorna o frame_locator."""
        return self._page.frame_locator(selector)

    def switch_to_root(self):
        """Sai do iframe (volta ao contexto raiz)."""
        return self._page

    # ========== DOWNLOAD ==========

    def wait_for_download(self, save_path: str = None, timeout: int = 30000) -> Optional[str]:
        """Espera download concluir e retorna o caminho do arquivo."""
        with self._page.expect_download(timeout=timeout) as download_info:
            pass
        download = download_info.value
        if save_path:
            download.save_as(save_path)
            return save_path
        return str(download.path())

    # ========== SCREENSHOT ==========

    def screenshot(self, filepath: str) -> bool:
        """Captura screenshot."""
        self._page.screenshot(path=filepath)
        return True

    # ========== HELPERS INTERNOS ==========

    def _get_locator(self, selector: str, frame: str = None):
        """Retorna o locator, com suporte a frame explícito."""
        if frame:
            return self._page.frame_locator(frame).locator(selector)
        return self._page.locator(selector)

    def _find_element_recursive(self, selector: str, timeout: int = 5000):
        """Busca elemento em todos os iframes."""
        start_time = time.time()
        while (time.time() - start_time) < (timeout / 1000):
            for frame in self._page.frames:
                try:
                    loc = frame.locator(selector).first
                    if loc.is_visible():
                        return loc
                except Exception:
                    continue
            time.sleep(0.5)
        raise Exception(f"Elemento não encontrado após {timeout}ms: {selector}")

    def _find_all_recursive(self, selector: str, timeout: int = 5000) -> list:
        """Busca múltiplos elementos em todos os iframes."""
        start_time = time.time()
        while (time.time() - start_time) < (timeout / 1000):
            for frame in self._page.frames:
                try:
                    loc = frame.locator(selector)
                    if loc.count() > 0:
                        return loc.all()
                except Exception:
                    continue
            time.sleep(0.5)
        return []

    def _find_and_click_recursive(self, selector: str, timeout: int = 5000) -> bool:
        """Busca e clica em qualquer iframe."""
        loc = self._find_element_recursive(selector, timeout)
        loc.click(timeout=timeout)
        return True

    def _find_and_hover_recursive(self, selector: str, timeout: int = 5000) -> bool:
        """Busca e passa o mouse em qualquer iframe."""
        loc = self._find_element_recursive(selector, timeout)
        loc.hover(timeout=timeout)
        return True

    def _find_and_select_recursive(self, selector: str, label: str, timeout: int = 5000) -> bool:
        """Busca e seleciona opção em qualquer iframe."""
        loc = self._find_element_recursive(selector, timeout)
        loc.select_option(label=label, timeout=timeout)
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
