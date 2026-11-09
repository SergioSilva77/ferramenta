"""Módulo Email v2 para rpaflow.

Diferenças em relação ao `rpaflow.email.Email` (v1), as principais:

- `EmailError` agora é importado (no v1 qualquer falha de SMTP virava `NameError`).
- Vários destinatários (lista ou "a@x.com; b@y.com"), `cc` e `bcc`. No v1 uma string com vírgula ia como UM
  endereço só e o servidor rejeitava.
- Nome do anexo correto no Windows (o v1 fazia `split('/')` e mandava "C:\\pasta\\arq.xlsx" inteiro como nome)
  e codificado em UTF-8 (acentos).
- `html=True` para corpo HTML; porta 465 usa SMTP_SSL; `timeout`; STARTTLS opcional.
- Conexão reaproveitada por `connect()` é validada antes do envio (reconecta se o servidor derrubou).
"""

from __future__ import annotations

import mimetypes
import os
import re
import smtplib
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from typing import Optional, Union

from rpaflow._exceptions import EmailError

Destinatarios = Union[str, list, tuple, None]


def _lista(valor: Destinatarios) -> list:
    """'a@x; b@y, c@z' ou ['a@x', 'b@y'] -> ['a@x', 'b@y', 'c@z']."""
    if not valor:
        return []
    if isinstance(valor, str):
        valor = re.split(r"[;,]", valor)
    return [v.strip() for v in valor if v and v.strip()]


class Email:
    """Classe para envio de emails via SMTP."""

    def __init__(
        self,
        smtp_host: str = "",
        smtp_port: int = 587,
        user: str = "",
        password: str = "",
        use_tls: bool = True,
        timeout: int = 30,
        sender: Optional[str] = None,
    ):
        """
        Args:
            use_tls: STARTTLS (porta 587). Ignorado na 465 (SSL direto). False para relay interno sem TLS.
            sender: endereço do "De" se diferente do login.
        """
        self.smtp_host = smtp_host
        self.smtp_port = smtp_port
        self.user = user
        self.password = password
        self.use_tls = use_tls
        self.timeout = timeout
        self.sender = sender or user
        self._server = None

    # ========== CONEXÃO ==========

    def _abrir(self):
        if self.smtp_port == 465:
            server = smtplib.SMTP_SSL(self.smtp_host, self.smtp_port, timeout=self.timeout)
        else:
            server = smtplib.SMTP(self.smtp_host, self.smtp_port, timeout=self.timeout)
            if self.use_tls:
                server.starttls()
        if self.user:
            server.login(self.user, self.password)
        return server

    def connect(self) -> bool:
        """Conecta ao servidor SMTP."""
        try:
            self.disconnect()
            self._server = self._abrir()
            return True
        except Exception as e:
            raise EmailError(f"Erro ao conectar: {e}") from e

    def disconnect(self) -> None:
        """Desconecta do servidor SMTP."""
        server, self._server = self._server, None
        if server:
            try:
                server.quit()
            except Exception:
                pass

    def __enter__(self) -> "Email":
        self.connect()
        return self

    def __exit__(self, *exc) -> bool:
        self.disconnect()
        return False

    # ========== ENVIO ==========

    def send(
        self,
        to: Destinatarios,
        subject: str,
        body: str,
        attachments: Optional[list] = None,
        html: bool = False,
        cc: Destinatarios = None,
        bcc: Destinatarios = None,
    ) -> bool:
        """Envia um email. `to`/`cc`/`bcc` aceitam string (separada por ; ou ,) ou lista."""
        destinos = _lista(to)
        copia = _lista(cc)
        oculta = _lista(bcc)
        if not (destinos or copia or oculta):
            raise EmailError("Nenhum destinatário informado.")

        try:
            msg = MIMEMultipart()
            msg["From"] = self.sender
            msg["To"] = ", ".join(destinos)
            if copia:
                msg["Cc"] = ", ".join(copia)
            msg["Subject"] = subject  # o BCC fica de fora do cabeçalho de propósito
            msg.attach(MIMEText(body, "html" if html else "plain", "utf-8"))

            for filepath in attachments or []:
                with open(filepath, "rb") as f:
                    dados = f.read()
                nome = os.path.basename(filepath)  # funciona com / e \ (o v1 só tratava /)
                tipo, _ = mimetypes.guess_type(nome)
                principal, _, sub = (tipo or "application/octet-stream").partition("/")
                parte = MIMEBase(principal, sub)
                parte.set_payload(dados)
                encoders.encode_base64(parte)
                parte.add_header("Content-Disposition", "attachment", filename=("utf-8", "", nome))
                msg.attach(parte)

            todos = destinos + copia + oculta
            if self._server is not None:
                try:
                    self._server.noop()
                except Exception:
                    self._server = self._abrir()  # servidor derrubou a conexão ociosa
                self._server.sendmail(self.sender, todos, msg.as_string())
            else:
                with self._abrir() as server:
                    server.sendmail(self.sender, todos, msg.as_string())
            return True
        except EmailError:
            raise
        except Exception as e:
            raise EmailError(f"Erro ao enviar email: {e}") from e
