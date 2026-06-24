"""Módulo Log para rpaflow - Logging profissional (inspirado no Serilog)."""

import sys
from typing import Optional


class Log:
    """Classe para logging profissional com loguru.

    Suporta console colorido e arquivo JSON com rotação.
    """

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
            console: Se True, exibe no console.
            format_console: Formato customizado do console.
            format_file: Formato customizado do arquivo.
        """
        try:
            from loguru import logger
        except ImportError:
            raise ImportError(
                "loguru é necessário para o módulo log. "
                "Instale com: pip install rpaflow[log]"
            )

        self._logger = logger
        self._context = {}

        # Remover handler padrão
        self._logger.remove()

        # Formato padrão do console
        if format_console is None:
            format_console = (
                "<green>{time:YYYY-MM-DD HH:mm:ss}</green> | "
                "<level>{level: <8}</level> | "
                "<cyan>{extra[context]}</cyan> | "
                "<level>{message}</level>"
            )

        # Formato padrão do arquivo
        if format_file is None:
            format_file = (
                "{time:YYYY-MM-DD HH:mm:ss} | {level: <8} | "
                "{extra[context]} | {message}"
            )

        # Adicionar console
        if console:
            self._logger.add(
                sys.stderr,
                level=level,
                format=format_console,
                colorize=True,
            )

        # Adicionar arquivo
        if path:
            if json:
                self._logger.add(
                    path,
                    level=level,
                    rotation=rotation,
                    retention=retention,
                    compression=compression,
                    serialize=True,
                )
            else:
                self._logger.add(
                    path,
                    level=level,
                    rotation=rotation,
                    retention=retention,
                    compression=compression,
                    format=format_file,
                )

        # Bind contexto vazio inicial
        self._logger = self._logger.bind(context="")

    def bind(self, **kwargs) -> "Log":
        """Adiciona contexto ao logger (tipo ForContext do Serilog).

        Args:
            **kwargs: Campos de contexto. Ex: user="admin", bot="vendas"

        Returns:
            Nova instância de Log com contexto.
        """
        new_log = Log.__new__(Log)
        new_log._context = {**self._context, **kwargs}
        context_str = " | ".join(f"{k}={v}" for k, v in new_log._context.items())
        new_log._logger = self._logger.bind(context=context_str)
        return new_log

    def debug(self, message: str, **kwargs) -> None:
        """Log DEBUG."""
        if kwargs:
            extra = " | ".join(f"{k}={v}" for k, v in kwargs.items())
            message = f"{message} | {extra}"
        self._logger.debug(message)

    def info(self, message: str, **kwargs) -> None:
        """Log INFO."""
        if kwargs:
            extra = " | ".join(f"{k}={v}" for k, v in kwargs.items())
            message = f"{message} | {extra}"
        self._logger.info(message)

    def warning(self, message: str, **kwargs) -> None:
        """Log WARNING."""
        if kwargs:
            extra = " | ".join(f"{k}={v}" for k, v in kwargs.items())
            message = f"{message} | {extra}"
        self._logger.warning(message)

    def error(self, message: str, **kwargs) -> None:
        """Log ERROR."""
        if kwargs:
            extra = " | ".join(f"{k}={v}" for k, v in kwargs.items())
            message = f"{message} | {extra}"
        self._logger.error(message)

    def critical(self, message: str, **kwargs) -> None:
        """Log CRITICAL."""
        if kwargs:
            extra = " | ".join(f"{k}={v}" for k, v in kwargs.items())
            message = f"{message} | {extra}"
        self._logger.critical(message)

    def success(self, message: str, **kwargs) -> None:
        """Log SUCCESS."""
        if kwargs:
            extra = " | ".join(f"{k}={v}" for k, v in kwargs.items())
            message = f"{message} | {extra}"
        self._logger.success(message)

    def exception(self, message: str, **kwargs) -> None:
        """Log EXCEPTION com traceback."""
        if kwargs:
            extra = " | ".join(f"{k}={v}" for k, v in kwargs.items())
            message = f"{message} | {extra}"
        self._logger.exception(message)
