"""Ini, Json e Files v2: cada teste cobre um bug real do v1."""

import datetime
import decimal
import os
import sys

import pytest

from rpaflow.v2.files import Files
from rpaflow.v2.ini import Ini
from rpaflow.v2.json import Json


# ───────────────────────── Ini ─────────────────────────
def test_ini_senha_com_percent_nao_estoura(tmp_path):
    p = tmp_path / "secrets.ini"
    p.write_text("[db]\nsenha=abc%123%\n", encoding="utf-8")
    assert Ini(str(p)).get("db", "senha") == "abc%123%"


def test_ini_com_bom(tmp_path):
    p = tmp_path / "settings.ini"
    p.write_bytes(b"\xef\xbb\xbf[browser]\ncdp_port=9555\n")
    assert Ini(str(p)).get_int("browser", "cdp_port") == 9555


def test_ini_valor_vazio_usa_fallback(tmp_path):
    p = tmp_path / "a.ini"
    p.write_text("[x]\nporta=\nativo=\nlimite=\n", encoding="utf-8")
    ini = Ini(str(p))
    assert ini.get_int("x", "porta", fallback=7) == 7
    assert ini.get_bool("x", "ativo", fallback=True) is True
    assert ini.get_float("x", "limite", fallback=1.5) == 1.5
    assert ini.get_int("x", "inexistente", fallback=3) == 3


def test_ini_bool_invalido_diz_qual_chave(tmp_path):
    p = tmp_path / "a.ini"
    p.write_text("[x]\nativo=talvez\n", encoding="utf-8")
    with pytest.raises(ValueError, match="ativo"):
        Ini(str(p)).get_bool("x", "ativo")


def test_ini_save_roundtrip_sem_bom_e_cria_pasta(tmp_path):
    destino = tmp_path / "nova" / "pasta" / "c.ini"
    ini = Ini(str(destino))
    ini.set("Email", "Remetente", "robo@x.com")
    ini.set("Email", "senha", "100%")
    ini.save()
    assert destino.read_bytes()[:3] != b"\xef\xbb\xbf"
    assert Ini(str(destino)).get("Email", "Remetente") == "robo@x.com"  # case da chave preservado
    assert Ini(str(destino)).get("Email", "senha") == "100%"


def test_ini_must_exist_e_get_or_raise(tmp_path):
    with pytest.raises(FileNotFoundError):
        Ini(str(tmp_path / "nao_existe.ini"), must_exist=True)
    p = tmp_path / "a.ini"
    p.write_text("[x]\nvazio=\n", encoding="utf-8")
    with pytest.raises(KeyError, match="vazio"):
        Ini(str(p)).get_or_raise("x", "vazio")


# ───────────────────────── Json ─────────────────────────
def test_json_lista_no_topo(tmp_path):
    p = tmp_path / "l.json"
    p.write_text('[{"nome": "Ana"}, {"nome": "Bia"}]', encoding="utf-8")
    dados = Json().load(str(p))
    assert dados[1].nome == "Bia"


def test_json_com_bom(tmp_path):
    p = tmp_path / "b.json"
    p.write_bytes(b"\xef\xbb\xbf" + '{"a": 1}'.encode("utf-8"))
    assert Json().load(str(p)).a == 1


def test_json_acento_igual_em_save_e_dumps(tmp_path):
    j = Json()
    dados = j.loads('{"cidade": "São Paulo"}')
    p = tmp_path / "x" / "a.json"  # pasta inexistente
    j.save(str(p), dados)
    assert "São Paulo" in p.read_text(encoding="utf-8")
    assert "São Paulo" in j.dumps(dados)


def test_json_serializa_tipos_exoticos(tmp_path):
    j = Json()
    txt = j.dumps({"d": datetime.date(2026, 11, 9), "v": decimal.Decimal("1.5"), "s": {"b", "a"}})
    assert '"2026-11-09"' in txt and '"1.5"' in txt and '"a"' in txt


def test_json_estrito_nao_inventa_chave():
    box = Json(default_box=False).loads('{"a": 1}')
    with pytest.raises(Exception):
        box.nao_existe


def test_json_get_value_caminho_pontuado():
    j = Json()
    d = j.loads('{"a": {"b": [10, 20]}, "x.y": 5}')
    assert j.get_value(d, "a.b.1") == 20
    assert j.get_value(d, "x.y") == 5
    assert j.get_value(d, "a.z", default="nada") == "nada"


# ───────────────────────── Files ─────────────────────────
def test_files_write_cria_pasta_e_le_bom(tmp_path):
    f = Files()
    alvo = str(tmp_path / "a" / "b" / "t.txt")
    f.write_text(alvo, "olá")
    assert f.read_text(alvo) == "olá"
    bom = tmp_path / "bom.txt"
    bom.write_bytes(b"\xef\xbb\xbfoi\r\nlinha2\r\n")
    assert f.read_text(str(bom)) == "oi\nlinha2\n"  # sem '\ufeff' no começo
    assert f.read_lines(str(bom)) == ["oi", "linha2"]


def test_files_listagem_so_arquivos_e_colchete_no_caminho(tmp_path):
    f = Files()
    base = tmp_path / "pasta [1]"
    (base / "sub").mkdir(parents=True)
    (base / "a.txt").write_text("x")
    (base / "sub" / "b.txt").write_text("x")
    assert [os.path.basename(p) for p in f.dir_get_files(str(base))] == ["a.txt"]
    assert sorted(os.path.basename(p) for p in f.dir_get_all_files(str(base), "*.txt")) == ["a.txt", "b.txt"]
    assert f.dir_get_dirs(str(base)) == [str(base / "sub")]


def test_files_delete_missing_ok_e_replace(tmp_path):
    f = Files()
    assert f.delete(str(tmp_path / "nao.txt"), missing_ok=True) is False
    with pytest.raises(FileNotFoundError):
        f.delete(str(tmp_path / "nao.txt"))
    (tmp_path / "novo.txt").write_text("novo")
    (tmp_path / "velho.txt").write_text("velho")
    f.replace(str(tmp_path / "novo.txt"), str(tmp_path / "velho.txt"), backup=str(tmp_path / "bkp" / "velho.bak"))
    assert (tmp_path / "velho.txt").read_text() == "novo"
    assert (tmp_path / "bkp" / "velho.bak").read_text() == "velho"
    assert not (tmp_path / "novo.txt").exists()


def test_files_move_overwrite_funciona(tmp_path):
    f = Files()
    (tmp_path / "a.txt").write_text("A")
    (tmp_path / "b.txt").write_text("B")
    f.move(str(tmp_path / "a.txt"), str(tmp_path / "b.txt"), overwrite=True)
    assert (tmp_path / "b.txt").read_text() == "A"


def test_files_path_helpers():
    f = Files()
    assert f.change_extension("arq.txt", "log") == "arq.log"
    assert f.ends_with_separator("pasta/")
    assert f.trim_separator("pasta/sub//") == "pasta/sub"
    assert len(f.get_random_filename()) == 32
    if sys.platform == "win32":
        assert f.get_path_root("C:\\x\\y.txt") == "C:\\"
        assert f.trim_separator("C:\\") == "C:\\"
        assert not f.is_fully_qualified("\\sem_drive")


def test_files_base_directory_e_script_nao_cwd():
    base = Files().get_base_directory()
    assert os.path.isdir(base)


@pytest.mark.skipif(sys.platform != "win32", reason="atributo hidden é do Windows")
def test_files_hidden_windows(tmp_path):
    import ctypes
    f = Files()
    p = tmp_path / "h.txt"
    p.write_text("x")
    f.set_attributes(str(p), hidden=True)
    assert ctypes.windll.kernel32.GetFileAttributesW(str(p)) & 0x2
    f.set_attributes(str(p), hidden=False)
    assert not ctypes.windll.kernel32.GetFileAttributesW(str(p)) & 0x2
