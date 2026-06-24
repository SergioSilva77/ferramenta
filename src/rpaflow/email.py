"""Módulo Email para rpaflow."""

import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from typing import Optional


class Email:
    """Classe para envio e leitura de emails."""

    def __init__(
        self,
        smtp_host: str = "",
        smtp_port: int = 587,
        user: str = "",
        password: str = "",
    ):
        self.smtp_host = smtp_host
        self.smtp_port = smtp_port
        self.user = user
        self.password = password
        self._server = None

    def connect(self) -> bool:
        """Conecta ao servidor SMTP."""
        try:
            self._server = smtplib.SMTP(self.smtp_host, self.smtp_port)
            self._server.starttls()
            self._server.login(self.user, self.password)
            return True
        except Exception as e:
            raise EmailError(f"Erro ao conectar: {e}")

    def send(
        self,
        to: str,
        subject: str,
        body: str,
        attachments: Optional[list] = None,
    ) -> bool:
        """Envia um email."""
        try:
            msg = MIMEMultipart()
            msg["From"] = self.user
            msg["To"] = to
            msg["Subject"] = subject
            msg.attach(MIMEText(body, "plain"))

            if attachments:
                for filepath in attachments:
                    with open(filepath, "rb") as f:
                        part = MIMEBase("application", "octet-stream")
                        part.set_payload(f.read())
                    encoders.encode_base64(part)
                    part.add_header(
                        "Content-Disposition",
                        f"attachment; filename={filepath.split('/')[-1]}",
                    )
                    msg.attach(part)

            if self._server:
                self._server.sendmail(self.user, to, msg.as_string())
            else:
                with smtplib.SMTP(self.smtp_host, self.smtp_port) as server:
                    server.starttls()
                    server.login(self.user, self.password)
                    server.sendmail(self.user, to, msg.as_string())

            return True
        except Exception as e:
            raise EmailError(f"Erro ao enviar email: {e}")

    def disconnect(self) -> None:
        """Desconecta do servidor SMTP."""
        if self._server:
            self._server.quit()
            self._server = None
