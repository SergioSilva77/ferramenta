"""Módulo Browser para rpaflow."""

from typing import Optional


class Browser:
    """Classe para automação de navegador."""

    def __init__(self):
        self._driver = None
        self._playwright = None
        self._browser = None

    def start(self, url: str = "", type: str = "playwright") -> bool:
        """Inicia o navegador."""
        try:
            if type == "playwright":
                from playwright.sync_api import sync_playwright
                self._playwright = sync_playwright().start()
                self._browser = self._playwright.chromium.launch(headless=False)
                self._driver = self._browser.new_page()
            elif type == "selenium":
                from selenium import webdriver
                self._driver = webdriver.Chrome()
            else:
                raise ValueError(f"Tipo '{type}' não suportado. Use 'playwright' ou 'selenium'.")

            if url:
                self._driver.goto(url)
            return True
        except Exception as e:
            raise BrowserError(f"Erro ao iniciar navegador: {e}")

    def click(self, selector: str) -> bool:
        """Clica em um elemento."""
        try:
            self._driver.click(selector)
            return True
        except Exception as e:
            raise BrowserError(f"Erro ao clicar: {e}")

    def type(self, selector: str, text: str) -> bool:
        """Digita em um campo."""
        try:
            self._driver.fill(selector, text)
            return True
        except Exception as e:
            raise BrowserError(f"Erro ao digitar: {e}")

    def get_text(self, selector: str) -> str:
        """Retorna o texto de um elemento."""
        try:
            return self._driver.text_content(selector) or ""
        except Exception as e:
            raise BrowserError(f"Erro ao obter texto: {e}")

    def screenshot(self, filepath: str) -> bool:
        """Captura uma screenshot."""
        try:
            self._driver.screenshot(path=filepath)
            return True
        except Exception as e:
            raise BrowserError(f"Erro ao capturar screenshot: {e}")

    def close(self) -> None:
        """Fecha o navegador."""
        if self._driver:
            self._driver.close()
        if self._browser:
            self._browser.close()
        if self._playwright:
            self._playwright.stop()
