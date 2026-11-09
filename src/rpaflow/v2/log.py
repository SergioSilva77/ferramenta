"""Módulo Log v2 para rpaflow - Logging profissional (loguru, inspirado no Serilog).

Diferenças em relação ao `rpaflow.log.Log` (v1), as principais:

- Criar um segundo `Log` NÃO apaga os handlers do primeiro. O v1 chamava `logger.remove()` — que remove
  TODOS os handlers do loguru, globais ao processo — então cada `Log(...)` novo calava os anteriores.
  Aqui cada instância registra e remove só os próprios handlers, e só recebe as próprias mensagens
  (dois robôs no mesmo processo não misturam arquivos).
- `close()` / `with Log(...)` removem os handlers desta instância e liberam o arquivo (importante p/ rotação,
  compressão e para apagar/mover o .log).
- Sem console (`pythonw.exe`, serviço do Windows: `sys.stderr` é None) não estoura ao registrar o handler.
- O `{extra[context]}` vazio não deixa " |  | " no meio da linha.
- A pasta do arquivo de log é criada se não existir.
"""

from __future__ import annotations

import itertools
import os
import sys
from typing import Optional

_ids = itertools.count(1)

_FMT_CONSOLE = (
    "<green>{time:YYYY-MM-DD HH:mm:ss}</green> | <level>{level: <8}</level> | "
    "{extra[prefixo]}<level>{message}</level>"
)
_FMT_ARQUIVO = "{time:YYYY-MM-DD HH:mm:ss} | {level: <8} | {extra[prefixo]}{message}"


class Log:
    """Classe para logging com loguru: console colorido e arquivo (texto ou JSON) com rotação."""

    def __init__(
        self,
        path: Optional[str] = None,
        level: str = "DEBUG",
        rotation: str = "10 MB",
        retention: str = "7 days",
        compression: str = "zip",
        json: bool = False,
        console: bool = True,
        format_console: Optional[str] = None,
        format_file: Optional[str] = None,
    ):
        """Configura o logger.

        Args:
            path: Caminho do arquivo de log (None = só console).
            level: Nível mínimo (DEBUG, INFO, WARNING, ERROR, CRITICAL).
            rotation: Rotação do arquivo ("10 MB", "00:00", "1 week").
            retention: Manter logs antigos ("7 days", "30 days").
            compression: Compressão dos logs rotados ("zip", "gz", "tar").
            json: Se True, output JSON no arquivo.
            console: Se True, exibe no console (ignorado se não houver console).
            format_console: Formato customizado do console (variáveis loguru + {extra[prefixo]}).
            format_file: Formato customizado do arquivo.
        """
        try:
            from loguru import logger
        except ImportError:
            raise ImportError(
                "loguru é necessário para o módulo log. "
                "Instale com: pip install rpaflow[log]"
            )

        self._id = next(_ids)
        self._root = logger
        self._handlers: list = []
        self._context: dict = {}

        # Cada instância só recebe as próprias mensagens (marcadas com _rpa_id).
        meu = self._id
        so_meu = lambda record: record["extra"].get("_rpa_id") == meu  # noqa: E731

        # O handler padrão do loguru (id 0, stderr) duplicaria o console e não conhece {extra[prefixo]}.
        try:
            logger.remove(0)
        except ValueError:
            pass  # já foi removido

        if console and sys.stderr is not None:
            self._handlers.append(logger.add(
                sys.stderr, level=level, format=format_console or _FMT_CONSOLE, colorize=True, filter=so_meu,
            ))

        if path:
            pasta = os.path.dirname(os.path.abspath(path))
            os.makedirs(pasta, exist_ok=True)
            opcoes = dict(level=level, rotation=rotation, retention=retention, compression=compression,
                          filter=so_meu, encoding="utf-8")
            if json:
                self._handlers.append(logger.add(path, serialize=True, **opcoes))
            else:
                self._handlers.append(logger.add(path, format=format_file or _FMT_ARQUIVO, **opcoes))

        self._logger = logger.bind(_rpa_id=self._id, prefixo="")

    # ========== CONTEXTO / CICLO DE VIDA ==========

    def bind(self, **kwargs) -> "Log":
        """Adiciona contexto ao logger (tipo ForContext do Serilog). Devolve um Log filho que compartilha os handlers."""
        filho = Log.__new__(Log)
        filho._id = self._id
        filho._root = self._root
        filho._handlers = []  # quem fecha os handlers é o Log original
        filho._context = {**self._context, **kwargs}
        texto = " | ".join(f"{k}={v}" for k, v in filho._context.items())
        filho._logger = self._root.bind(_rpa_id=self._id, prefixo=f"{texto} | " if texto else "")
        return filho

    def close(self) -> None:
        """Remove os handlers desta instância (fecha o arquivo). Os de outras instâncias não são afetados."""
        for h in self._handlers:
            try:
                self._root.remove(h)
            except ValueError:
                pass
        self._handlers = []

    def __enter__(self) -> "Log":
        return self

    def __exit__(self, *exc) -> bool:
        self.close()
        return False

    # ========== MENSAGENS ==========

    @staticmethod
    def _montar(message, kwargs: dict) -> str:
        message = str(message)
        if kwargs:
            message = f"{message} | " + " | ".join(f"{k}={v}" for k, v in kwargs.items())
        return message

    # `opt(depth=1)` faz o loguru atribuir a linha/módulo ao CHAMADOR, não a este wrapper.
    def debug(self, message: str, **kwargs) -> None:
        """Log DEBUG."""
        self._logger.opt(depth=1).debug("{}", self._montar(message, kwargs))

    def info(self, message: str, **kwargs) -> None:
        """Log INFO."""
        self._logger.opt(depth=1).info("{}", self._montar(message, kwargs))

    def warning(self, message: str, **kwargs) -> None:
        """Log WARNING."""
        self._logger.opt(depth=1).warning("{}", self._montar(message, kwargs))

    def error(self, message: str, **kwargs) -> None:
        """Log ERROR."""
        self._logger.opt(depth=1).error("{}", self._montar(message, kwargs))

    def critical(self, message: str, **kwargs) -> None:
        """Log CRITICAL."""
        self._logger.opt(depth=1).critical("{}", self._montar(message, kwargs))

    def success(self, message: str, **kwargs) -> None:
        """Log SUCCESS."""
        self._logger.opt(depth=1).success("{}", self._montar(message, kwargs))

    def exception(self, message: str, **kwargs) -> None:
        """Log ERROR com traceback (chame dentro do `except`)."""
        self._logger.opt(depth=1, exception=True).error("{}", self._montar(message, kwargs))
