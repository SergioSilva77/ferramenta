"""Módulo Files v2 para rpaflow - Operações com arquivos e diretórios.

Diferenças em relação ao `rpaflow.files.Files` (v1), as principais:

- Escrita/cópia/movimentação criam a pasta de destino (o v1 estourava `FileNotFoundError`).
- Leitura de texto aceita arquivo com BOM (`utf-8-sig`); o v1 devolvia '\\ufeff' no começo.
- `dir_get_files`/`dir_get_all_files` devolvem só arquivos (o v1 misturava pastas) e não quebram com `[` no caminho.
- `dir_get_dirs` devolve caminho completo, como `dir_get_files` (o v1 devolvia só o nome).
- `get_base_directory` é a pasta do script, não o diretório atual — no Agendador de Tarefas o cwd é System32.
- `set_attributes(hidden=...)` funciona (o v1 ignorava), `change_extension('x.txt', 'log')` não estoura,
  `trim_separator`/`ends_with_separator` entendem `/` no Windows, `get_path_root` devolve 'C:\\' como diz a doc,
  `get_random_filename` sem `mktemp` (deprecated e com race).
"""

import glob
import os
import shutil
import stat
import sys
import tempfile
import uuid
from datetime import datetime
from pathlib import Path
from typing import Optional


def _garantir_pasta(filepath: str) -> None:
    pasta = os.path.dirname(os.path.abspath(filepath))
    os.makedirs(pasta, exist_ok=True)


class Files:
    """Classe para operações com arquivos e diretórios.

    Inspirado nos métodos estáticos de System.IO.File e System.IO.Path do C#.
    """

    # ========== LEITURA/ESCRITA TEXTO ==========

    def read_text(self, filepath: str, encoding: str = "utf-8-sig") -> str:
        """Lê conteúdo do arquivo como texto (tolera BOM)."""
        with open(filepath, "r", encoding=encoding) as f:
            return f.read()

    def write_text(self, filepath: str, content: str, encoding: str = "utf-8") -> bool:
        """Escreve texto no arquivo (sobrescreve). Cria a pasta se não existir."""
        _garantir_pasta(filepath)
        with open(filepath, "w", encoding=encoding) as f:
            f.write(content)
        return True

    def append_text(self, filepath: str, content: str, encoding: str = "utf-8") -> bool:
        """Adiciona texto ao final do arquivo."""
        _garantir_pasta(filepath)
        with open(filepath, "a", encoding=encoding) as f:
            f.write(content)
        return True

    def read_lines(self, filepath: str, encoding: str = "utf-8-sig") -> list:
        """Lê linhas do arquivo como lista (sem \\n nem \\r)."""
        with open(filepath, "r", encoding=encoding) as f:
            return f.read().splitlines()

    def write_lines(self, filepath: str, lines: list, encoding: str = "utf-8") -> bool:
        """Escreve lista de linhas no arquivo."""
        _garantir_pasta(filepath)
        with open(filepath, "w", encoding=encoding) as f:
            for line in lines:
                f.write(line + "\n")
        return True

    def append_lines(self, filepath: str, lines: list, encoding: str = "utf-8") -> bool:
        """Adiciona lista de linhas ao final do arquivo."""
        _garantir_pasta(filepath)
        with open(filepath, "a", encoding=encoding) as f:
            for line in lines:
                f.write(line + "\n")
        return True

    # ========== LEITURA/ESCRITA BYTES ==========

    def read_bytes(self, filepath: str) -> bytes:
        """Lê conteúdo do arquivo como bytes."""
        with open(filepath, "rb") as f:
            return f.read()

    def write_bytes(self, filepath: str, data: bytes) -> bool:
        """Escreve bytes no arquivo."""
        _garantir_pasta(filepath)
        with open(filepath, "wb") as f:
            f.write(data)
        return True

    def append_bytes(self, filepath: str, data: bytes) -> bool:
        """Adiciona bytes ao final do arquivo."""
        _garantir_pasta(filepath)
        with open(filepath, "ab") as f:
            f.write(data)
        return True

    # ========== OPERAÇÕES DE ARQUIVO ==========

    def exists(self, filepath: str) -> bool:
        """Verifica se arquivo ou diretório existe."""
        return os.path.exists(filepath)

    def copy(self, src: str, dst: str, overwrite: bool = False) -> bool:
        """Copia arquivo."""
        if not overwrite and os.path.exists(dst):
            raise FileExistsError(f"Arquivo destino já existe: {dst}")
        _garantir_pasta(dst)
        shutil.copy2(src, dst)
        return True

    def move(self, src: str, dst: str, overwrite: bool = False) -> bool:
        """Move arquivo."""
        if not overwrite and os.path.exists(dst):
            raise FileExistsError(f"Arquivo destino já existe: {dst}")
        _garantir_pasta(dst)
        if overwrite and os.path.exists(dst):
            os.remove(dst)  # shutil.move não sobrescreve arquivo existente no Windows
        shutil.move(src, dst)
        return True

    def delete(self, filepath: str, missing_ok: bool = False) -> bool:
        """Deleta arquivo. missing_ok=True não reclama se ele já não existe (como o File.Delete do C#)."""
        try:
            os.remove(filepath)
        except FileNotFoundError:
            if not missing_ok:
                raise
            return False
        return True

    def replace(self, source: str, destination: str, backup: str = None) -> bool:
        """Substitui arquivo destino pelo arquivo source, criando backup se especificado."""
        if backup and os.path.exists(destination):
            _garantir_pasta(backup)
            shutil.copy2(destination, backup)
        _garantir_pasta(destination)
        try:
            os.replace(source, destination)  # atômico quando é no mesmo volume
        except OSError:
            shutil.copy2(source, destination)  # volumes diferentes
            os.remove(source)
        return True

    # ========== INFORMAÇÕES DO ARQUIVO ==========

    def get_creation_time(self, filepath: str) -> datetime:
        """Retorna data de criação do arquivo."""
        return datetime.fromtimestamp(os.path.getctime(filepath))

    def get_last_write_time(self, filepath: str) -> datetime:
        """Retorna data da última escrita do arquivo."""
        return datetime.fromtimestamp(os.path.getmtime(filepath))

    def get_last_access_time(self, filepath: str) -> datetime:
        """Retorna data do último acesso do arquivo."""
        return datetime.fromtimestamp(os.path.getatime(filepath))

    def get_attributes(self, filepath: str) -> dict:
        """Retorna atributos do arquivo."""
        st = os.stat(filepath)
        return {
            "size": st.st_size,
            "mode": st.st_mode,
            "created": datetime.fromtimestamp(st.st_ctime),
            "modified": datetime.fromtimestamp(st.st_mtime),
            "accessed": datetime.fromtimestamp(st.st_atime),
            "is_file": os.path.isfile(filepath),
            "is_dir": os.path.isdir(filepath),
            "is_readonly": not os.access(filepath, os.W_OK),
        }

    def set_attributes(self, filepath: str, readonly: bool = None, hidden: bool = None) -> bool:
        """Define atributos do arquivo. `hidden` só tem efeito no Windows (no v1 era ignorado)."""
        if readonly is not None:
            self.set_readonly(filepath, readonly)
        if hidden is not None and sys.platform == "win32":
            import ctypes
            atual = ctypes.windll.kernel32.GetFileAttributesW(str(filepath))
            if atual == -1:
                raise FileNotFoundError(filepath)
            novo = (atual | 0x2) if hidden else (atual & ~0x2)  # FILE_ATTRIBUTE_HIDDEN
            if not ctypes.windll.kernel32.SetFileAttributesW(str(filepath), novo):
                raise OSError(f"Não foi possível alterar os atributos de {filepath}")
        return True

    def create_symlink(self, link: str, target: str) -> bool:
        """Cria link simbólico."""
        os.symlink(target, link)
        return True

    # ========== PATH ==========

    def get_filename(self, filepath: str) -> str:
        """Retorna nome do arquivo com extensão. Ex: 'C:/pasta/arquivo.txt' -> 'arquivo.txt'"""
        return os.path.basename(filepath)

    def get_filename_without_ext(self, filepath: str) -> str:
        """Retorna nome do arquivo sem extensão. Ex: 'arquivo.txt' -> 'arquivo'"""
        return Path(filepath).stem

    def get_extension(self, filepath: str) -> str:
        """Retorna extensão do arquivo. Ex: 'arquivo.txt' -> '.txt'"""
        return Path(filepath).suffix

    def get_directory(self, filepath: str) -> str:
        """Retorna diretório pai. Ex: 'C:/pasta/arquivo.txt' -> 'C:/pasta'"""
        return str(Path(filepath).parent)

    def combine(self, *paths: str) -> str:
        """Junta caminhos. Ex: combine('C:/pasta', 'arquivo.txt') -> 'C:/pasta/arquivo.txt'"""
        return str(Path(*paths))

    def join(self, *paths: str) -> str:
        """Junta caminhos (mais moderno)."""
        return os.path.join(*paths)

    def get_full_path(self, filepath: str) -> str:
        """Retorna caminho absoluto."""
        return os.path.abspath(filepath)

    def get_relative_path(self, base: str, filepath: str) -> str:
        """Retorna caminho relativo. Ex: get_relative_path('C:/pasta', 'C:/pasta/arquivo.txt') -> 'arquivo.txt'"""
        return os.path.relpath(filepath, base)

    def is_rooted(self, filepath: str) -> bool:
        """Verifica se é caminho absoluto."""
        return os.path.isabs(filepath)

    def is_fully_qualified(self, filepath: str) -> bool:
        """Verifica se é caminho completo (no Windows exige drive ou UNC: '\\\\pasta' sozinho não conta)."""
        if os.name == "nt":
            return bool(os.path.splitdrive(filepath)[0]) and os.path.isabs(filepath)
        return os.path.isabs(filepath)

    def has_extension(self, filepath: str) -> bool:
        """Verifica se tem extensão."""
        return bool(Path(filepath).suffix)

    def change_extension(self, filepath: str, new_ext: str) -> str:
        """Troca extensão. Ex: change_extension('arquivo.txt', '.log') -> 'arquivo.log'"""
        if new_ext and not new_ext.startswith("."):
            new_ext = "." + new_ext  # with_suffix('log') estoura ValueError
        return str(Path(filepath).with_suffix(new_ext))

    def _separadores(self) -> tuple:
        return (os.sep, os.altsep) if os.altsep else (os.sep,)

    def trim_separator(self, filepath: str) -> str:
        """Remove separador final do caminho (preserva a raiz: 'C:\\\\' continua 'C:\\\\')."""
        if filepath == Path(filepath).anchor:
            return filepath
        return filepath.rstrip("".join(self._separadores()))

    def ends_with_separator(self, filepath: str) -> bool:
        """Verifica se termina com separador de diretório ('/' ou '\\\\' no Windows)."""
        return filepath.endswith(self._separadores())

    def get_temp_path(self) -> str:
        """Retorna diretório temporário do sistema."""
        return tempfile.gettempdir()

    def get_temp_file(self) -> str:
        """Cria arquivo temporário e retorna o caminho."""
        fd, path = tempfile.mkstemp()
        os.close(fd)
        return path

    def get_random_filename(self) -> str:
        """Retorna nome de arquivo aleatório."""
        return uuid.uuid4().hex

    def get_invalid_filename_chars(self) -> list:
        """Retorna caracteres inválidos no nome do arquivo."""
        return list('<>:"/\\|?*')

    def get_invalid_path_chars(self) -> list:
        """Retorna caracteres inválidos no caminho."""
        return list('<>:"|?*')

    def get_path_root(self, filepath: str) -> str:
        """Retorna raiz do caminho. Ex: 'C:/pasta/arquivo.txt' -> 'C:\\'"""
        return Path(filepath).anchor or os.sep

    def get_base_directory(self) -> str:
        """Retorna diretório base: pasta do executável (PyInstaller) ou do script que iniciou o programa.

        O v1 devolvia o diretório ATUAL, que no Agendador de Tarefas é C:\\\\Windows\\\\System32.
        """
        if getattr(sys, "frozen", False):
            return os.path.dirname(sys.executable)
        main = sys.modules.get("__main__")
        main_file = getattr(main, "__file__", None)
        if main_file:
            return os.path.dirname(os.path.abspath(main_file))
        return os.getcwd()

    # ========== DIRECTORY ==========

    def dir_create(self, path: str) -> bool:
        """Cria diretório (e pais se necessário)."""
        os.makedirs(path, exist_ok=True)
        return True

    def dir_delete(self, path: str, recursive: bool = False) -> bool:
        """Deleta diretório."""
        if recursive:
            shutil.rmtree(path)
        else:
            os.rmdir(path)
        return True

    def dir_exists(self, path: str) -> bool:
        """Verifica se diretório existe."""
        return os.path.isdir(path)

    def dir_move(self, src: str, dst: str) -> bool:
        """Move diretório."""
        shutil.move(src, dst)
        return True

    def dir_get_files(self, path: str, pattern: str = "*") -> list:
        """Lista arquivos do diretório (caminho completo; só arquivos, sem pastas)."""
        achados = glob.glob(os.path.join(glob.escape(path), pattern))  # escape: '[' no caminho quebra o glob
        return sorted(p for p in achados if os.path.isfile(p))

    def dir_get_dirs(self, path: str) -> list:
        """Lista subdiretórios do diretório (caminho completo, como dir_get_files)."""
        return sorted(
            os.path.join(path, d) for d in os.listdir(path) if os.path.isdir(os.path.join(path, d))
        )

    def dir_get_all_files(self, path: str, pattern: str = "*") -> list:
        """Lista todos os arquivos recursivamente (só arquivos, sem pastas)."""
        achados = glob.glob(os.path.join(glob.escape(path), "**", pattern), recursive=True)
        return sorted(p for p in achados if os.path.isfile(p))

    # ========== FILE INFO ==========

    def get_size(self, filepath: str) -> int:
        """Retorna tamanho do arquivo em bytes."""
        return os.path.getsize(filepath)

    def is_readonly(self, filepath: str) -> bool:
        """Verifica se arquivo é somente leitura."""
        return not os.access(filepath, os.W_OK)

    def set_readonly(self, filepath: str, readonly: bool = True) -> bool:
        """Define arquivo como somente leitura (preserva as demais permissões)."""
        modo = os.stat(filepath).st_mode
        if readonly:
            os.chmod(filepath, modo & ~(stat.S_IWUSR | stat.S_IWGRP | stat.S_IWOTH))
        else:
            os.chmod(filepath, modo | stat.S_IWUSR)
        return True

    # ========== COMPATIBILIDADE (métodos antigos) ==========

    def read(self, filepath: str) -> str:
        """Lê conteúdo do arquivo (compatibilidade)."""
        return self.read_text(filepath)

    def write(self, filepath: str, content: str) -> bool:
        """Escreve conteúdo no arquivo (compatibilidade)."""
        return self.write_text(filepath, content)
