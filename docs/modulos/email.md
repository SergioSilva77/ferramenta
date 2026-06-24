# Email

Envio de emails via SMTP.

## Instalar

```bash
pip install rpaflow[email]
```

## Exemplo

```python
from rpaflow.email import Email

email = Email(smtp_host="smtp.gmail.com", smtp_port=587, user="meu@email.com", password="senha")

# Conectar
email.connect()

# Enviar
email.send(
    to="dest@email.com",
    subject="Assunto",
    body="Mensagem do email",
    attachments=["arquivo.pdf"]
)

# Desconectar
email.disconnect()
```

## Métodos

| Método | Parâmetros | Descrição |
|--------|-----------|-----------|
| `connect()` | — | Conecta ao SMTP |
| `send()` | to, subject, body, attachments | Envia email |
| `disconnect()` | — | Desconecta |
