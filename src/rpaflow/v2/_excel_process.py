"""Excel isolado por robô (Windows).

Por que existe: `Dispatch("Excel.Application")` agarra QUALQUER Excel que já esteja rodando na máquina.
Com robôs em paralelo, o `Quit()` de um fecha as planilhas do outro, e a abertura seguinte pega um Excel
que ainda está morrendo ("disconnected from its clients" / "rejected by callee").
O remédio clássico — `taskkill /IM EXCEL.EXE` — derruba o Excel de todo mundo, inclusive o do usuário.

Aqui cada instância sobe o SEU Excel (`DispatchEx`, processo novo) e só encerra o PID dele. Nunca por nome.
Se o processo Python morrer sem fechar o Excel, a próxima execução limpa só aquele PID (dono morto,
mesmo exe, mesma hora de criação — senão o PID foi reaproveitado por outro programa e não mexe).
"""

from __future__ import annotations

import atexit
import ctypes
import json
import os
import subprocess
import tempfile
import time
from ctypes import wintypes
from pathlib import Path
from typing import Optional

_STILL_ACTIVE = 259
_QUERY_LIMITED = 0x1000
_ESPERA_ENCERRAR_S = 5

_k32 = None


def _kernel32():
    """Carrega o kernel32 só quando precisa (o módulo precisa importar em qualquer SO)."""
    global _k32
    if _k32 is None:
        k = ctypes.WinDLL("kernel32", use_last_error=True)
        k.OpenProcess.argtypes = [wintypes.DWORD, wintypes.BOOL, wintypes.DWORD]
        k.OpenProcess.restype = wintypes.HANDLE
        k.CloseHandle.argtypes = [wintypes.HANDLE]
        k.GetExitCodeProcess.argtypes = [wintypes.HANDLE, ctypes.POINTER(wintypes.DWORD)]
        k.GetProcessTimes.argtypes = [wintypes.HANDLE] + [ctypes.POINTER(wintypes.FILETIME)] * 4
        k.QueryFullProcessImageNameW.argtypes = [
            wintypes.HANDLE, wintypes.DWORD, wintypes.LPWSTR, ctypes.POINTER(wintypes.DWORD),
        ]
        _k32 = k
    return _k32


def _pasta_pids() -> Path:
    """Pasta com o controle de órfãos. Compartilhada entre robôs: a limpeza só mexe em dono morto."""
    base = os.environ.get("RPAFLOW_PIDS_DIR")
    if base:
        return Path(base)
    raiz = os.environ.get("LOCALAPPDATA") or tempfile.gettempdir()
    return Path(raiz) / "rpaflow" / "excel_pids"


# id(xl) -> (xl, pid). Guardar o xl evita o id ser reaproveitado antes do encerramento.
_INSTANCIAS: dict = {}
_orfaos_limpos = False


# ─── Processos ───────────────────────────────────
def info_processo(pid: int) -> Optional[tuple]:
    """(nome do exe, hora de criação) de um processo VIVO. None se não existe mais."""
    k = _kernel32()
    h = k.OpenProcess(_QUERY_LIMITED, False, pid)
    if not h:
        return None
    try:
        codigo = wintypes.DWORD()
        if not k.GetExitCodeProcess(h, ctypes.byref(codigo)) or codigo.value != _STILL_ACTIVE:
            return None
        c, e, kt, u = (wintypes.FILETIME() for _ in range(4))
        if not k.GetProcessTimes(h, ctypes.byref(c), ctypes.byref(e), ctypes.byref(kt), ctypes.byref(u)):
            return None
        criacao = (c.dwHighDateTime << 32) | c.dwLowDateTime
        buf = ctypes.create_unicode_buffer(1024)
        tam = wintypes.DWORD(1024)
        nome = buf.value if k.QueryFullProcessImageNameW(h, 0, buf, ctypes.byref(tam)) else ""
        return os.path.basename(nome).lower(), criacao
    finally:
        k.CloseHandle(h)


def _matar_pid(pid: int) -> None:
    """taskkill por PID (e só por PID)."""
    subprocess.run(["taskkill", "/PID", str(pid), "/F"], capture_output=True, text=True)


def _pid_da_janela(xl) -> Optional[int]:
    try:
        pid = wintypes.DWORD()
        ctypes.windll.user32.GetWindowThreadProcessId(int(xl.Hwnd), ctypes.byref(pid))
        return int(pid.value) or None
    except Exception:
        return None


# ─── Registro de PIDs deste processo ─────────────
def _salvar_registro() -> None:
    pasta = _pasta_pids()
    arquivo = pasta / f"{os.getpid()}.json"
    pids = [pid for _, pid in _INSTANCIAS.values() if pid]
    try:
        if not pids:
            arquivo.unlink(missing_ok=True)
            return
        dono = info_processo(os.getpid())
        excel = []
        for pid in pids:
            info = info_processo(pid)
            if info:
                excel.append({"pid": pid, "criacao": info[1]})
        pasta.mkdir(parents=True, exist_ok=True)
        arquivo.write_text(
            json.dumps({"dono": {"pid": os.getpid(), "criacao": dono[1] if dono else 0}, "excel": excel}),
            encoding="utf-8",
        )
    except OSError:
        pass  # controle de órfãos é bônus: nunca derruba o robô


def limpar_orfaos() -> int:
    """Mata Excel que ficou de uma execução anterior MORTA — e só ele. Retorna quantos matou."""
    pasta = _pasta_pids()
    if not pasta.exists():
        return 0
    mortos = 0
    for arq in pasta.glob("*.json"):
        try:
            dados = json.loads(arq.read_text(encoding="utf-8"))
            dono = dados.get("dono", {})
            info_dono = info_processo(int(dono.get("pid", 0)))
            if info_dono and info_dono[1] == dono.get("criacao"):
                continue  # dono ainda vivo (outra execução em andamento): não mexe
            for item in dados.get("excel", []):
                info = info_processo(int(item["pid"]))
                if info and info[0] == "excel.exe" and info[1] == item.get("criacao"):
                    _matar_pid(int(item["pid"]))
                    mortos += 1
            arq.unlink(missing_ok=True)
        except (OSError, ValueError, KeyError):
            continue
    return mortos


# ─── API ─────────────────────────────────────────
def abrir_excel(visible: bool = False, display_alerts: bool = False):
    """Sobe um Excel novo (processo próprio) e registra o PID dele."""
    import win32com.client

    global _orfaos_limpos
    if not _orfaos_limpos:
        _orfaos_limpos = True
        limpar_orfaos()

    xl = win32com.client.DispatchEx("Excel.Application")
    pid = _pid_da_janela(xl)
    _INSTANCIAS[id(xl)] = (xl, pid)
    try:
        xl.Visible = visible
        xl.DisplayAlerts = display_alerts
    except Exception:
        encerrar_excel(xl)
        raise
    _salvar_registro()
    return xl


def pedir_quit(xl) -> Optional[int]:
    """1ª metade do encerramento: manda o Quit e tira do registro. Devolve o PID.

    Depois disso o CHAMADOR precisa soltar a própria referência COM (`del xl`) e só então chamar
    `aguardar_encerramento(pid)`. O Excel só deixa o processo morrer quando as referências COM acabam;
    segurar `xl` durante a espera faz o processo "lingerar" até o timeout (e acabar morto à força).
    """
    entrada = _INSTANCIAS.pop(id(xl), (None, None))
    pid = entrada[1]
    try:
        xl.Quit()
    except Exception:
        pass
    return pid


def aguardar_encerramento(pid: Optional[int]) -> None:
    """2ª metade: espera o processo sair; se não sair, mata o PID dele (nunca outro)."""
    import gc
    gc.collect()  # derruba proxies COM órfãos que ainda seguram o Excel
    if pid:
        limite = time.time() + _ESPERA_ENCERRAR_S
        while time.time() < limite and info_processo(pid):
            time.sleep(0.3)
        info = info_processo(pid)
        if info and info[0] == "excel.exe":
            _matar_pid(pid)
    _salvar_registro()


def encerrar_excel(xl) -> None:
    """Quit + espera + kill do PID, numa chamada só. Atalho para quem não precisa soltar a referência antes
    (se o chamador seguir segurando `xl`, a espera vai até o timeout e o PID é morto à força)."""
    if xl is None:
        return
    pid = pedir_quit(xl)
    del xl
    aguardar_encerramento(pid)


@atexit.register
def _encerrar_todos() -> None:
    pids = []
    for chave in list(_INSTANCIAS):
        pids.append(pedir_quit(_INSTANCIAS[chave][0]))
    _INSTANCIAS.clear()  # solta as referências antes de esperar os processos
    for pid in pids:
        aguardar_encerramento(pid)
