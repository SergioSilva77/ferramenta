"""Módulo Desktop para rpaflow - Automação via reconhecimento de imagem."""

import shutil
import tempfile
import time
from pathlib import Path
from typing import Optional


class Desktop:
    """Classe para automação desktop via reconhecimento de imagem.

    Requer: pyautogui, opencv-python, pillow
    """

    def __init__(self):
        try:
            import pyautogui
            self._pyautogui = pyautogui
        except ImportError:
            raise ImportError(
                "pyautogui é necessário para o módulo desktop. "
                "Instale com: pip install rpaflow[desktop]"
            )

    # ========== LOCALIZAR ==========

    def find_image(
        self,
        image: str,
        confidence: float = 0.80,
        limit: int = 10,
        time_ms: int = 1000,
    ) -> Optional[dict]:
        """Localiza imagem na tela e retorna coordenadas.

        Args:
            image: Caminho completo da imagem de referência.
            confidence: Nível de confiança (0.80 a 1.00).
            limit: Máximo de tentativas.
            time_ms: Tempo entre tentativas (ms).

        Returns:
            dict com x, y, width, height, confidence, attempts, elapsed_ms
            ou None se não encontrar.
        """
        image_path, temp_path = self._prepare_image(image)
        try:
            regions, attempts, elapsed_ms = self._locate_all(
                image_path, confidence, limit, time_ms
            )
            if not regions:
                return None

            region = regions[0]
            x, y = self._compute_center(region)
            return {
                "x": x,
                "y": y,
                "width": int(region.width),
                "height": int(region.height),
                "confidence": confidence,
                "attempts": attempts,
                "elapsed_ms": elapsed_ms,
            }
        finally:
            self._cleanup_temp(temp_path)

    # ========== CLICAR ==========

    def click_image(
        self,
        image: str,
        confidence: float = 0.80,
        offset: str = "center",
        offset_x: int = 0,
        offset_y: int = 0,
        match_index: int = 0,
        highlight_ms: int = 0,
        limit: int = 10,
        time_ms: int = 1000,
    ) -> Optional[dict]:
        """Localiza e clica na imagem.

        Args:
            image: Caminho completo da imagem de referência.
            confidence: Nível de confiança (0.80 a 1.00).
            offset: Ponto de clique: 'left', 'right', 'top', 'bottom', 'center'.
            offset_x: Deslocamento X adicional (pixels).
            offset_y: Deslocamento Y adicional (pixels).
            match_index: Índice da ocorrência (0-based).
            highlight_ms: Tempo para mostrar borda vermelha (ms).
            limit: Máximo de tentativas.
            time_ms: Tempo entre tentativas (ms).

        Returns:
            dict com x, y, width, height, confidence, attempts, elapsed_ms
            ou None se não encontrar.
        """
        image_path, temp_path = self._prepare_image(image)
        try:
            regions, attempts, elapsed_ms = self._locate_all(
                image_path, confidence, limit, time_ms
            )
            if not regions:
                return None

            if match_index >= len(regions):
                return None

            region = regions[match_index]
            left = int(region.left)
            top = int(region.top)
            width = int(region.width)
            height = int(region.height)

            x, y = self._compute_target_point(
                left, top, width, height, offset, offset_x, offset_y
            )

            if highlight_ms > 0:
                self._highlight_region(left, top, width, height, highlight_ms)

            self._pyautogui.click(x=x, y=y)

            return {
                "x": x,
                "y": y,
                "width": width,
                "height": height,
                "confidence": confidence,
                "attempts": attempts,
                "elapsed_ms": elapsed_ms,
            }
        finally:
            self._cleanup_temp(temp_path)

    # ========== ENCONTRAR TODAS ==========

    def find_all_images(
        self,
        image: str,
        confidence: float = 0.80,
        limit: int = 10,
        time_ms: int = 1000,
    ) -> list:
        """Localiza todas as ocorrências da imagem na tela.

        Args:
            image: Caminho completo da imagem de referência.
            confidence: Nível de confiança (0.80 a 1.00).
            limit: Máximo de tentativas.
            time_ms: Tempo entre tentativas (ms).

        Returns:
            Lista de dicts com index, x, y, width, height.
        """
        image_path, temp_path = self._prepare_image(image)
        try:
            regions, attempts, elapsed_ms = self._locate_all(
                image_path, confidence, limit, time_ms
            )
            if not regions:
                return []

            results = []
            for i, region in enumerate(regions):
                x, y = self._compute_center(region)
                results.append({
                    "index": i,
                    "x": x,
                    "y": y,
                    "width": int(region.width),
                    "height": int(region.height),
                })
            return results
        finally:
            self._cleanup_temp(temp_path)

    # ========== HELPERS INTERNOS ==========

    def _prepare_image(self, image_path: str) -> tuple:
        """Prepara o caminho da imagem (lida com caracteres especiais)."""
        path = Path(image_path)
        if not path.exists() or not path.is_file():
            raise FileNotFoundError(f"Imagem não encontrada: {image_path}")

        try:
            image_path.encode("ascii")
            return image_path, None
        except UnicodeEncodeError:
            suffix = path.suffix or ".png"
            with tempfile.NamedTemporaryFile(
                prefix="rpa_desktop_", suffix=suffix, delete=False
            ) as tmp:
                temp_path = Path(tmp.name)
            shutil.copyfile(path, temp_path)
            return str(temp_path), temp_path

    def _cleanup_temp(self, temp_path) -> None:
        """Remove arquivo temporário se existir."""
        if temp_path:
            try:
                temp_path.unlink(missing_ok=True)
            except Exception:
                pass

    def _locate_all(self, image: str, confidence: float, limit: int, sleep_ms: int):
        """Localiza todas as ocorrências da imagem na tela."""
        start = time.perf_counter()
        attempts = 0
        regions = []

        while attempts < limit:
            attempts += 1
            regions = list(
                self._pyautogui.locateAllOnScreen(image, confidence=confidence)
            )
            if regions:
                break
            if attempts < limit and sleep_ms > 0:
                time.sleep(sleep_ms / 1000.0)

        elapsed_ms = int((time.perf_counter() - start) * 1000)
        return regions, attempts, elapsed_ms

    def _compute_center(self, region) -> tuple:
        """Calcula o centro de uma região."""
        x = int(region.left) + int(region.width) // 2
        y = int(region.top) + int(region.height) // 2
        return x, y

    def _compute_target_point(
        self, left, top, width, height, offset_mode, offset_x, offset_y
    ) -> tuple:
        """Calcula o ponto de clique baseado no offset."""
        center_x = left + (width // 2)
        center_y = top + (height // 2)
        right_x = left + max(width - 1, 0)
        bottom_y = top + max(height - 1, 0)

        if offset_mode == "left":
            base_x, base_y = left, center_y
        elif offset_mode == "right":
            base_x, base_y = right_x, center_y
        elif offset_mode == "top":
            base_x, base_y = center_x, top
        elif offset_mode == "bottom":
            base_x, base_y = center_x, bottom_y
        else:
            base_x, base_y = center_x, center_y

        return base_x + offset_x, base_y + offset_y

    def _highlight_region(self, left, top, width, height, highlight_ms) -> None:
        """Exibe borda vermelha temporária ao redor da região encontrada."""
        try:
            import tkinter as tk

            safe_width = max(width, 2)
            safe_height = max(height, 2)
            root = tk.Tk()
            root.overrideredirect(True)
            root.attributes("-topmost", True)

            transparent_color = "white"
            root.configure(bg=transparent_color)
            try:
                root.wm_attributes("-transparentcolor", transparent_color)
            except Exception:
                pass

            root.geometry(f"{safe_width}x{safe_height}+{left}+{top}")
            canvas = tk.Canvas(
                root,
                width=safe_width,
                height=safe_height,
                highlightthickness=0,
                bg=transparent_color,
            )
            canvas.pack()
            canvas.create_rectangle(
                1,
                1,
                max(safe_width - 2, 1),
                max(safe_height - 2, 1),
                outline="red",
                width=3,
            )
            root.after(highlight_ms, root.destroy)
            root.mainloop()
        except Exception:
            pass
