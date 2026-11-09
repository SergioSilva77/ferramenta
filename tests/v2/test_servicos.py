"""SQL, Excel (openpyxl), Email, API e Log v2."""

import email as _email
import smtplib

import pytest

from rpaflow._exceptions import APIError, EmailError, FileError, QueryError
from rpaflow.v2.api import API
from rpaflow.v2.email import Email
from rpaflow.v2.excel import Excel
from rpaflow.v2.log import Log
from rpaflow.v2.sql import SQL


# ───────────────────────── SQL ─────────────────────────
@pytest.fixture
def db(tmp_path):
    with SQL(type="sqlite", database=str(tmp_path / "t.db")) as d:
        d.execute("CREATE TABLE pessoa (id INTEGER PRIMARY KEY, nome TEXT, idade INTEGER)")
        yield d


def test_sql_crud(db):
    db.insert("pessoa", {"id": 1, "nome": "Ana", "idade": 30})
    db.insert("pessoa", {"id": 2, "nome": "Bia", "idade": 25})
    assert db.select("pessoa", {"nome": "Ana"}) == [(1, "Ana", 30)]
    db.update("pessoa", {"idade": 31}, {"id": 1})
    assert db.last_rowcount == 1
    assert db.select("pessoa", {"id": 1})[0][2] == 31
    db.delete("pessoa", {"id": 2})
    assert len(db.select("pessoa")) == 1


def test_sql_erro_vira_query_error_e_nao_name_error(db):
    with pytest.raises(QueryError):
        db.execute("SELECT * FROM tabela_que_nao_existe")


def test_sql_placeholder_interrogacao_e_dicts(db):
    db.execute("INSERT INTO pessoa VALUES (?, ?, ?)", [1, "Ana", 30])
    assert db.fetch_dicts("SELECT nome FROM pessoa WHERE id = ?", [1]) == [{"nome": "Ana"}]
    assert db.execute("UPDATE pessoa SET idade = 1") == []  # não devolve linhas, e não estoura


def test_sql_update_delete_sem_where_recusados(db):
    with pytest.raises(ValueError):
        db.update("pessoa", {"idade": 1}, {})
    with pytest.raises(ValueError):
        db.delete("pessoa", {})


def test_sql_valida_identificador(db):
    with pytest.raises(ValueError):
        db.select("pessoa; DROP TABLE pessoa")
    with pytest.raises(ValueError):
        db.insert("pessoa", {"nome) VALUES ('x'); --": 1})


def test_sql_transacao_faz_rollback(db):
    with pytest.raises(RuntimeError):
        with db.transaction():
            db.insert("pessoa", {"id": 1, "nome": "Ana", "idade": 1})
            raise RuntimeError("falhou no meio")
    assert db.select("pessoa") == []


def test_sql_execute_many(db):
    n = db.execute_many("INSERT INTO pessoa VALUES (?, ?, ?)", [[1, "A", 1], [2, "B", 2], [3, "C", 3]])
    assert n == 3 and len(db.select("pessoa")) == 3


def test_sql_traduz_placeholder_ignorando_aspas():
    d = SQL(type="mysql")
    assert d._traduzir("SELECT '?' , a FROM t WHERE x = ? AND y = ?") == "SELECT '?' , a FROM t WHERE x = %s AND y = %s"


def test_sql_sem_conexao_mensagem_clara():
    with pytest.raises(Exception, match="connect"):
        SQL(type="sqlite", database=":memory:").execute("SELECT 1")


# ───────────────────────── Excel (openpyxl) ─────────────────────────
def test_excel_write_e_read_celula_e_bloco(tmp_path):
    caminho = str(tmp_path / "x" / "p.xlsx")
    with Excel(caminho, create_if_missing=True) as xl:
        xl.write(None, "B2", [[1, 2], [3, 4]])
        xl.write(None, "A1", ["a", "b"])
        assert xl.read(None, "B2") == [[1]]
        assert xl.read(None, "B2:C3") == [[1, 2], [3, 4]]
        assert xl.read(None, "A1:B1") == [["a", "b"]]
        xl.save()
    with Excel(caminho) as xl:
        assert xl.read()[1][1:] == [1, 2]


def test_excel_arquivo_inexistente_nao_cria_em_silencio(tmp_path):
    with pytest.raises(FileError):
        Excel(str(tmp_path / "typo.xlsx")).open()


def test_excel_sheet_inexistente_e_sem_open(tmp_path):
    xl = Excel(str(tmp_path / "p.xlsx"), create_if_missing=True)
    with pytest.raises(FileError, match="open"):
        xl.read("Dados")
    xl.open()
    with pytest.raises(FileError, match="Dados"):
        xl.read("Dados")


# ───────────────────────── Email ─────────────────────────
class FakeSMTP:
    enviados = []
    falhar = False

    def __init__(self, host, port, timeout=None):
        self.host, self.port = host, port

    def starttls(self): pass
    def login(self, u, p): pass
    def noop(self): pass
    def quit(self): pass
    def __enter__(self): return self
    def __exit__(self, *a): return False

    def sendmail(self, de, para, texto):
        if FakeSMTP.falhar:
            raise smtplib.SMTPException("recusado")
        FakeSMTP.enviados.append((de, para, texto))


@pytest.fixture
def smtp(monkeypatch):
    FakeSMTP.enviados = []
    FakeSMTP.falhar = False
    monkeypatch.setattr(smtplib, "SMTP", FakeSMTP)
    return FakeSMTP


def test_email_varios_destinatarios_cc_bcc_e_anexo_windows(smtp, tmp_path):
    anexo = tmp_path / "relatório final.xlsx"
    anexo.write_bytes(b"dados")
    Email("smtp.x", 587, "robo@x.com", "senha").send(
        to="a@x.com; b@x.com", cc=["c@x.com"], bcc="d@x.com",
        subject="Olá", body="<b>oi</b>", html=True, attachments=[str(anexo).replace("/", "\\")],
    )
    de, para, texto = smtp.enviados[0]
    assert de == "robo@x.com"
    assert para == ["a@x.com", "b@x.com", "c@x.com", "d@x.com"]
    msg = _email.message_from_string(texto)
    assert "d@x.com" not in (msg["To"] + (msg["Cc"] or ""))  # bcc não aparece no cabeçalho
    anexos = [p for p in msg.walk() if p.get_filename()]
    assert anexos[0].get_filename() == "relatório final.xlsx"  # só o nome, nunca "C:\...\"
    assert msg.get_payload(0).get_content_type() == "text/html"


def test_email_falha_vira_email_error(smtp):
    smtp.falhar = True
    with pytest.raises(EmailError, match="recusado"):
        Email("smtp.x", 587, "robo@x.com", "s").send("a@x.com", "assunto", "corpo")


def test_email_sem_destinatario():
    with pytest.raises(EmailError):
        Email("smtp.x").send("", "a", "b")


# ───────────────────────── API ─────────────────────────
@pytest.fixture
def api_mock():
    import httpx

    def handler(request: httpx.Request) -> httpx.Response:
        if request.url.path == "/erro":
            return httpx.Response(500, text="boom")
        if request.url.path == "/redir":
            return httpx.Response(302, headers={"location": "/ok"})
        return httpx.Response(200, json={"q": dict(request.url.params)})

    api = API(base_url="http://teste", raise_for_status=True)
    api._client = httpx.Client(
        transport=httpx.MockTransport(handler), base_url="http://teste", follow_redirects=True, timeout=5,
    )
    return api


def test_api_params_redirect_e_raise_for_status(api_mock):
    assert api_mock.get("/ok", params={"a": "1"}).json() == {"q": {"a": "1"}}
    assert api_mock.get("/redir").status_code == 200
    with pytest.raises(APIError, match="500"):
        api_mock.get("/erro")


def test_api_falha_de_rede_vira_api_error():
    api = API(timeout=0.5)
    with pytest.raises(APIError):
        api.get("http://127.0.0.1:9/nada")


# ───────────────────────── Log ─────────────────────────
def test_log_instancias_independentes_e_close(tmp_path):
    a = tmp_path / "a.log"
    b = tmp_path / "b" / "b.log"  # pasta inexistente
    la = Log(path=str(a), console=False)
    lb = Log(path=str(b), console=False)  # o v1 apagaria o handler de `la` aqui
    la.info("msg-a")
    lb.info("msg-b", chave=1)
    la.bind(robo="fiat").warning("com {chaves} literais")
    la.close()
    lb.close()
    ta, tb = a.read_text(encoding="utf-8"), b.read_text(encoding="utf-8")
    assert "msg-a" in ta and "msg-b" not in ta
    assert "msg-b | chave=1" in tb and "msg-a" not in tb
    assert "robo=fiat | com {chaves} literais" in ta
    la.info("depois do close")  # não explode nem escreve
    assert "depois do close" not in a.read_text(encoding="utf-8")
    a.unlink()  # arquivo liberado (no Windows falharia com handler aberto)
