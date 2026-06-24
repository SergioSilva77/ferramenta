"""Módulo Files para rpaflow - Operações com arquivos e diretórios."""

import glob
import os
import shutil
import stat
import sys
import tempfile
from datetime import datetime
from pathlib import Path
from typing import Optional


class Files:
    """Classe para operações com arquivos e diretórios.

    Inspirado nos métodos estáticos de System.IO.File e System.IO.Path do C#.
    """

    # ========== LEITURA/ESCRITA TEXTO ==========

    def read_text(self, filepath: str, encoding: str = "utf-8") -> str:
        """Lê conteúdo do arquivo como texto."""
        with open(filepath, "r", encoding=encoding) as f:
            return f.read()

    def write_text(self, filepath: str, content: str, encoding: str = "utf-8") -> bool:
        """Escreve texto no arquivo (sobrescreve)."""
        with open(filepath, "w", encoding=encoding) as f:
            f.write(content)
        return True

    def append_text(self, filepath: str, content: str, encoding: str = "utf-8") -> bool:
        """Adiciona texto ao final do arquivo."""
        with open(filepath, "a", encoding=encoding) as f:
            f.write(content)
        return True

    def read_lines(self, filepath: str, encoding: str = "utf-8") -> list:
        """Lê linhas do arquivo como lista."""
        with open(filepath, "r", encoding=encoding) as f:
            return [line.rstrip("\n") for line in f.readlines()]

    def write_lines(self, filepath: str, lines: list, encoding: str = "utf-8") -> bool:
        """Escreve lista de linhas no arquivo."""
        with open(filepath, "w", encoding=encoding) as f:
            for line in lines:
                f.write(line + "\n")
        return True

    def append_lines(self, filepath: str, lines: list, encoding: str = "utf-8") -> bool:
        """Adiciona lista de linhas ao final do arquivo."""
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
        with open(filepath, "wb") as f:
            f.write(data)
        return True

    def append_bytes(self, filepath: str, data: bytes) -> bool:
        """Adiciona bytes ao final do arquivo."""
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
        shutil.copy2(src, dst)
        return True

    def move(self, src: str, dst: str, overwrite: bool = False) -> bool:
        """Move arquivo."""
        if not overwrite and os.path.exists(dst):
            raise FileExistsError(f"Arquivo destino já existe: {dst}")
        shutil.move(src, dst)
        return True

    def delete(self, filepath: str) -> bool:
        """Deleta arquivo."""
        os.remove(filepath)
        return True

    def replace(self, source: str, destination: str, backup: str = None) -> bool:
        """Substitui arquivo destino pelo arquivo source, criando backup se especificado."""
        if backup and os.path.exists(destination):
            shutil.copy2(destination, backup)
        shutil.copy2(source, destination)
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
        """Define atributos do arquivo."""
        if readonly is not None:
            if readonly:
                os.chmod(filepath, stat.S_IREAD)
            else:
                os.chmod(filepath, stat.S_IREAD | stat.S_IWRITE)
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
        """Verifica se é caminho completo."""
        return os.path.isabs(filepath)

    def has_extension(self, filepath: str) -> bool:
        """Verifica se tem extensão."""
        return bool(Path(filepath).suffix)

    def change_extension(self, filepath: str, new_ext: str) -> str:
        """Troca extensão. Ex: change_extension('arquivo.txt', '.log') -> 'arquivo.log'"""
        return str(Path(filepath).with_suffix(new_ext))

    def trim_separator(self, filepath: str) -> str:
        """Remove separador final do caminho."""
        return filepath.rstrip(os.sep)

    def ends_with_separator(self, filepath: str) -> bool:
        """Verifica se termina com separador de diretório."""
        return filepath.endswith(os.sep)

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
        return tempfile.mktemp().split(os.sep)[-1]

    def get_invalid_filename_chars(self) -> list:
        """Retorna caracteres inválidos no nome do arquivo."""
        return list('<>:"/\\|?*')

    def get_invalid_path_chars(self) -> list:
        """Retorna caracteres inválidos no caminho."""
        return list('<>:"|?*')

    def get_path_root(self, filepath: str) -> str:
        """Retorna raiz do caminho. Ex: 'C:/pasta/arquivo.txt' -> 'C:\\'"""
        return os.path.splitdrive(filepath)[0] or os.sep

    def get_base_directory(self) -> str:
        """Retorna diretório base (equivalente a AppDomain.CurrentDomain.BaseDirectory)."""
        if getattr(sys, "frozen", False):
            return os.path.dirname(sys.executable)
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
        """Lista arquivos do diretório."""
        return glob.glob(os.path.join(path, pattern))

    def dir_get_dirs(self, path: str) -> list:
        """Lista subdiretórios do diretório."""
        return [d for d in os.listdir(path) if os.path.isdir(os.path.join(path, d))]

    def dir_get_all_files(self, path: str, pattern: str = "*") -> list:
        """Lista todos os arquivos recursivamente."""
        return glob.glob(os.path.join(path, "**", pattern), recursive=True)

    # ========== FILE INFO ==========

    def get_size(self, filepath: str) -> int:
        """Retorna tamanho do arquivo em bytes."""
        return os.path.getsize(filepath)

    def is_readonly(self, filepath: str) -> bool:
        """Verifica se arquivo é somente leitura."""
        return not os.access(filepath, os.W_OK)

    def set_readonly(self, filepath: str, readonly: bool = True) -> bool:
        """Define arquivo como somente leitura."""
        if readonly:
            os.chmod(filepath, stat.S_IREAD)
        else:
            os.chmod(filepath, stat.S_IREAD | stat.S_IWRITE)
        return True

    # ========== COMPATIBILIDADE (métodos antigos) ==========

    def read(self, filepath: str) -> str:
        """Lê conteúdo do arquivo (compatibilidade)."""
        return self.read_text(filepath)

    def write(self, filepath: str, content: str) -> bool:
        """Escreve conteúdo no arquivo (compatibilidade)."""
        return self.write_text(filepath, content)
