"""Módulo Files para rpaflow."""

import shutil
import os


class Files:
    """Classe para operações com arquivos."""

    def read(self, filepath: str) -> str:
        """Lê o conteúdo de um arquivo."""
        try:
            with open(filepath, "r", encoding="utf-8") as f:
                return f.read()
        except Exception as e:
            raise FileError(f"Erro ao ler arquivo: {e}")

    def write(self, filepath: str, content: str) -> bool:
        """Escreve conteúdo em um arquivo."""
        try:
            with open(filepath, "w", encoding="utf-8") as f:
                f.write(content)
            return True
        except Exception as e:
            raise FileError(f"Erro ao escrever arquivo: {e}")

    def copy(self, src: str, dst: str) -> bool:
        """Copia um arquivo."""
        try:
            shutil.copy2(src, dst)
            return True
        except Exception as e:
            raise FileError(f"Erro ao copiar arquivo: {e}")

    def move(self, src: str, dst: str) -> bool:
        """Move um arquivo."""
        try:
            shutil.move(src, dst)
            return True
        except Exception as e:
            raise FileError(f"Erro ao mover arquivo: {e}")

    def delete(self, filepath: str) -> bool:
        """Deleta um arquivo."""
        try:
            os.remove(filepath)
            return True
        except Exception as e:
            raise FileError(f"Erro ao deletar arquivo: {e}")
