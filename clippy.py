import base64
import os
import platform
import subprocess
import time
from dataclasses import dataclass
from typing import Any, Dict, List, Optional

import requests
import typer

app = typer.Typer(add_completion=False)

CONFIG_PATH_DEFAULT = os.path.expanduser("~/.config/clippy.toml")
IS_WINDOWS = platform.system().lower() == "windows"
IS_DARWIN = platform.system().lower() == "darwin"
IS_LINUX = platform.system().lower() == "linux"


@dataclass
class Config:
    server_url: str
    token: str
    device_id: str
    device_name: str
    priority: List[str]
    use_latest_per_device: bool = True


def shutil_which(cmd: str) -> Optional[str]:
    from shutil import which

    return which(cmd)


def _run(cmd: List[str], input_text: Optional[str] = None) -> str:
    p = subprocess.run(
        cmd,
        input=(input_text.encode("utf-8") if input_text is not None else None),
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        check=True,
    )
    return p.stdout.decode("utf-8", errors="replace")


def load_config(path: str) -> Config:
    try:
        import tomllib  # py3.11+
    except Exception:
        raise RuntimeError("Python 3.11+ required for tomllib. Or install tomli.")

    with open(path, "rb") as f:
        d = tomllib.load(f)

    return Config(
        server_url=d["server_url"].rstrip("/"),
        token=d["token"],
        device_id=d["device_id"],
        device_name=d.get("device_name", d["device_id"]),
        priority=list(d.get("priority", [])),
        use_latest_per_device=bool(d.get("use_latest_per_device", True)),
    )


def auth_headers(cfg: Config) -> Dict[str, str]:
    return {"Authorization": f"Bearer {cfg.token}"}


def choose_clip(cfg: Config, items: List[Dict[str, Any]]) -> Optional[Dict[str, Any]]:
    if not items:
        return None
    if not cfg.priority:
        return items[0]  # newest first

    newest_per_device: Dict[str, Dict[str, Any]] = {}
    for it in items:
        did = it.get("device_id", "")
        if did and did not in newest_per_device:
            newest_per_device[did] = it

    for did in cfg.priority:
        if did in newest_per_device:
            return newest_per_device[did]

    return items[0]


# ---------------- Windows helpers (only defined on Windows) ----------------
if IS_WINDOWS:
    import ctypes
    from ctypes import wintypes

    user32 = ctypes.WinDLL("user32", use_last_error=True)
    kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
    # Some Python builds lack wintypes.ULONG_PTR
    ULONG_PTR = getattr(wintypes, "ULONG_PTR", ctypes.c_size_t)

    CF_UNICODETEXT = 13
    GMEM_MOVEABLE = 0x0002

    INPUT_KEYBOARD = 1
    KEYEVENTF_KEYUP = 0x0002
    KEYEVENTF_UNICODE = 0x0004

    VK_CONTROL = 0x11
    VK_C = 0x43

    class KEYBDINPUT(ctypes.Structure):
        _fields_ = [
            ("wVk", wintypes.WORD),
            ("wScan", wintypes.WORD),
            ("dwFlags", wintypes.DWORD),
            ("time", wintypes.DWORD),
            ("dwExtraInfo", ULONG_PTR),
        ]

    class INPUT(ctypes.Structure):
        class _I(ctypes.Union):
            _fields_ = [("ki", KEYBDINPUT)]

        _anonymous_ = ("i",)
        _fields_ = [("type", wintypes.DWORD), ("i", _I)]

    def _send_key(vk: int, keyup: bool = False):
        flags = KEYEVENTF_KEYUP if keyup else 0
        inp = INPUT(
            type=INPUT_KEYBOARD,
            ki=KEYBDINPUT(wVk=vk, wScan=0, dwFlags=flags, time=0, dwExtraInfo=0),
        )
        n = user32.SendInput(1, ctypes.byref(inp), ctypes.sizeof(inp))
        if n != 1:
            raise ctypes.WinError(ctypes.get_last_error())

    def win_send_ctrl_c():
        _send_key(VK_CONTROL, False)
        _send_key(VK_C, False)
        _send_key(VK_C, True)
        _send_key(VK_CONTROL, True)

    def win_type_text(text: str):
        # Paste by typing Unicode keystrokes. Clipboard not used.
        for ch in text:
            code = ord(ch)
            down = INPUT(
                type=INPUT_KEYBOARD,
                ki=KEYBDINPUT(
                    wVk=0, wScan=code, dwFlags=KEYEVENTF_UNICODE, time=0, dwExtraInfo=0
                ),
            )
            up = INPUT(
                type=INPUT_KEYBOARD,
                ki=KEYBDINPUT(
                    wVk=0,
                    wScan=code,
                    dwFlags=KEYEVENTF_UNICODE | KEYEVENTF_KEYUP,
                    time=0,
                    dwExtraInfo=0,
                ),
            )
            user32.SendInput(1, ctypes.byref(down), ctypes.sizeof(down))
            user32.SendInput(1, ctypes.byref(up), ctypes.sizeof(up))

    def win_get_clipboard_text() -> str:
        if not user32.OpenClipboard(None):
            raise ctypes.WinError(ctypes.get_last_error())
        try:
            h = user32.GetClipboardData(CF_UNICODETEXT)
            if not h:
                return ""
            p = kernel32.GlobalLock(h)
            if not p:
                return ""
            try:
                return ctypes.wstring_at(p)
            finally:
                kernel32.GlobalUnlock(h)
        finally:
            user32.CloseClipboard()

    def win_set_clipboard_text(text: str) -> None:
        if not user32.OpenClipboard(None):
            raise ctypes.WinError(ctypes.get_last_error())
        try:
            user32.EmptyClipboard()
            data = text.encode("utf-16le") + b"\x00\x00"
            h_mem = kernel32.GlobalAlloc(GMEM_MOVEABLE, len(data))
            if not h_mem:
                raise ctypes.WinError(ctypes.get_last_error())
            p = kernel32.GlobalLock(h_mem)
            if not p:
                raise ctypes.WinError(ctypes.get_last_error())
            try:
                ctypes.memmove(p, data, len(data))
            finally:
                kernel32.GlobalUnlock(h_mem)
            if not user32.SetClipboardData(CF_UNICODETEXT, h_mem):
                raise ctypes.WinError(ctypes.get_last_error())
        finally:
            user32.CloseClipboard()

    def win_capture_selected_text_preserve_clipboard() -> str:
        def safe_get():
            try:
                return win_get_clipboard_text()
            except Exception:
                return ""

        def safe_set(s: str):
            try:
                win_set_clipboard_text(s)
            except Exception:
                pass

        original = safe_get()

        # Trigger Ctrl+C
        win_send_ctrl_c()

        # Wait up to ~0.5s for clipboard to change
        new_text = ""
        for _ in range(50):
            time.sleep(0.01)
            new_text = safe_get()
            if new_text != original and new_text != "":
                break

        # Restore clipboard no matter what
        safe_set(original)

        return new_text


# ---------------- X11 helpers ----------------
def x11_get_primary_text() -> str:
    try:
        return _run(["xclip", "-selection", "primary", "-o"])
    except Exception:
        return ""


def x11_get_clipboard_text() -> str:
    return _run(["xclip", "-selection", "clipboard", "-o"])


def x11_set_clipboard_text(text: str) -> None:
    p = subprocess.Popen(
        ["xclip", "-selection", "clipboard", "-in"],
        stdin=subprocess.PIPE,
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
        close_fds=True,
    )
    assert p.stdin is not None
    p.stdin.write(text.encode("utf-8"))
    p.stdin.close()
    try:
        p.wait(timeout=0.5)
    except subprocess.TimeoutExpired:
        p.kill()


def x11_ctrl_c():
    subprocess.run(["xdotool", "key", "--clearmodifiers", "ctrl+c"], check=True)


def x11_type_text(text: str):
    subprocess.run(
        ["xdotool", "type", "--clearmodifiers", "--delay", "0", text],
        check=True,
    )


# ---------------- Generic clipboard (fallback) ----------------
def clipboard_get_text() -> str:
    sysname = platform.system().lower()

    if sysname == "darwin":
        return _run(["pbpaste"])
    if sysname == "windows":
        return _run(["powershell", "-NoProfile", "-Command", "Get-Clipboard -Raw"])

    # Linux
    if shutil_which("wl-paste"):
        return _run(["wl-paste", "--no-newline"])
    if shutil_which("xclip"):
        return _run(["xclip", "-selection", "clipboard", "-o"])
    if shutil_which("xsel"):
        return _run(["xsel", "--clipboard", "--output"])
    raise RuntimeError("No clipboard tool found (need wl-paste or xclip/xsel).")


def clipboard_set_text(text: str) -> None:
    sysname = platform.system().lower()

    if sysname == "darwin":
        _run(["pbcopy"], input_text=text)
        return

    if sysname == "windows":
        _run(
            [
                "powershell",
                "-NoProfile",
                "-Command",
                "Set-Clipboard -Value ([Console]::In.ReadToEnd())",
            ],
            input_text=text,
        )
        return

    # Linux (prefer X11 tools if present)
    if shutil_which("xclip"):
        p = subprocess.Popen(
            ["xclip", "-selection", "clipboard", "-in"],
            stdin=subprocess.PIPE,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            close_fds=True,
        )
        assert p.stdin is not None
        p.stdin.write(text.encode("utf-8"))
        p.stdin.close()
        try:
            p.wait(timeout=0.5)
        except subprocess.TimeoutExpired:
            p.kill()
        return

    if shutil_which("xsel"):
        p = subprocess.Popen(
            ["xsel", "--clipboard", "--input"],
            stdin=subprocess.PIPE,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            close_fds=True,
        )
        assert p.stdin is not None
        p.stdin.write(text.encode("utf-8"))
        p.stdin.close()
        try:
            p.wait(timeout=0.5)
        except subprocess.TimeoutExpired:
            p.kill()
        return

    raise RuntimeError("No clipboard tool found (need xclip or xsel).")


def paste_active_window():
    sysname = platform.system().lower()

    if sysname == "darwin":
        subprocess.run(
            [
                "osascript",
                "-e",
                'tell application "System Events" to keystroke "v" using command down',
            ],
            check=True,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )
        return

    if sysname == "linux" and shutil_which("xdotool"):
        subprocess.run(["xdotool", "key", "--clearmodifiers", "ctrl+v"], check=True)
        return

    raise RuntimeError(
        "No paste method available (mac: osascript, linux x11: xdotool)."
    )


@app.command()
def send(config: str = typer.Option(CONFIG_PATH_DEFAULT, "--config")):
    cfg = load_config(config)

    try:
        if platform.system().lower() == "windows":
            text = win_capture_selected_text_preserve_clipboard()
        else:
            text = clipboard_get_text()

        if not text.strip():
            typer.echo("send: nothing captured (select text first)")
            raise typer.Exit(code=2)

        payload = {
            "device_id": cfg.device_id,
            "device_name": cfg.device_name,
            "content_type": "text/plain",
            "content_text": text,
        }

        r = requests.post(
            f"{cfg.server_url}/v1/clip",
            json=payload,
            headers=auth_headers(cfg),
            timeout=5,
        )
        if r.status_code != 200:
            typer.echo(f"send: http {r.status_code}", err=True)
            typer.echo(r.text[:300], err=True)
            raise typer.Exit(code=3)

        typer.echo("send: ok")
    except Exception as e:
        typer.echo(f"send: error {e!r}", err=True)
        raise


@app.command()
def fetch(config: str = typer.Option(CONFIG_PATH_DEFAULT, "--config")):
    cfg = load_config(config)

    typer.echo("fetch: requesting...")

    url = (
        f"{cfg.server_url}/v1/latest_per_device?include_content=1"
        if cfg.use_latest_per_device
        else f"{cfg.server_url}/v1/clips?limit=50&include_content=1"
    )

    r = requests.get(url, headers=auth_headers(cfg), timeout=(3, 10))
    typer.echo(f"fetch: http {r.status_code}")

    if r.status_code != 200:
        typer.echo(r.text[:500], err=True)
        raise typer.Exit(code=1)

    items = r.json().get("items", [])
    typer.echo(f"fetch: got {len(items)} items")

    chosen = choose_clip(cfg, items)
    if not chosen:
        typer.echo("fetch: no clips")
        raise typer.Exit(code=2)

    typer.echo(f"fetch: chosen {chosen.get('clip_id')} from {chosen.get('device_id')}")

    b64 = chosen.get("content_b64", "")
    if not b64:
        typer.echo("fetch: missing content_b64", err=True)
        raise typer.Exit(code=3)

    raw = base64.b64decode(b64.encode("ascii"))
    text = raw.decode("utf-8", errors="replace")

    # Windows: paste by typing (no clipboard)
    if platform.system().lower() == "windows":
        if not IS_WINDOWS:
            raise typer.Exit(code=4)
        win_type_text(text)
        return

    # Linux X11: paste by typing (no clipboard)
    if platform.system().lower() == "linux" and shutil_which("xdotool"):
        x11_type_text(text)
        typer.echo("fetch: done (typed, clipboard untouched)")
        return

    # macOS fallback: temporary clipboard swap + paste + restore
    typer.echo("fetch: pasting (clipboard preserved)...")
    original = clipboard_get_text()
    clipboard_set_text(text)
    time.sleep(0.03)
    paste_active_window()
    time.sleep(0.05)
    clipboard_set_text(original)

    typer.echo("fetch: done (clipboard preserved)")


if __name__ == "__main__":
    app()
