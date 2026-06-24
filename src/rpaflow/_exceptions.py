"""Exceções do rpaflow."""


class RPAFlowError(Exception):
    """Exceção base do rpaflow."""


class ConnectionError(RPAFlowError):
    """Erro de conexão com banco de dados ou serviço."""


class QueryError(RPAFlowError):
    """Erro ao executar query SQL."""


class FileError(RPAFlowError):
    """Erro em operações com arquivos."""


class BrowserError(RPAFlowError):
    """Erro em automação de navegador."""


class APIError(RPAFlowError):
    """Erro em requisições HTTP."""


class EmailError(RPAFlowError):
    """Erro em operações de email."""
