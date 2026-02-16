import base64
import ctypes
import os
import platform
import subprocess
import time
from ctypes import wintypes
from dataclasses import dataclass
from typing import Any, Dict, List, Optional

import requests
import typer

app = typer.Typer(add_completion=False)

CONFIG_PATH_DEFAULT = os.path.expanduser("~/.config/clippy.toml")

# ---------------- Windows low-level helpers ----------------

user32 = ctypes.WinDLL("user32", use_last_error=True)
kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)

CF_UNICODETEXT = 13
GMEM_MOVEABLE = 0x0002

INPUT_KEYBOARD = 1
KEYEVENTF_KEYUP = 0x0002
KEYEVENTF_UNICODE = 0x0004

VK_CONTROL = 0x11
VK_V = 0x56
VK_C = 0x43


class KEYBDINPUT(ctypes.Structure):
    _fields_ = [
        ("wVk", wintypes.WORD),
        ("wScan", wintypes.WORD),
        ("dwFlags", wintypes.DWORD),
        ("time", wintypes.DWORD),
        ("dwExtraInfo", wintypes.ULONG_PTR),
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


def win_send_ctrl_combo(vk_letter: int):
    _send_key(VK_CONTROL, False)
    _send_key(vk_letter, False)
    _send_key(vk_letter, True)
    _send_key(VK_CONTROL, True)


def win_type_text(text: str):
    # Paste by typing Unicode keystrokes, clipboard not used.
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
        # do not free h_mem after SetClipboardData, OS owns it
    finally:
        user32.CloseClipboard()


def win_capture_selected_text_preserve_clipboard() -> str:
    original = ""
    try:
        original = win_get_clipboard_text()
    except Exception:
        original = ""

    # Trigger copy in the focused app
    win_send_ctrl_combo(VK_C)
    time.sleep(0.03)

    new_text = ""
    try:
        new_text = win_get_clipboard_text()
    except Exception:
        new_text = ""

    # Restore clipboard
    try:
        win_set_clipboard_text(original)
    except Exception:
        pass

    return new_text


@dataclass
class Config:
    server_url: str
    token: str
    device_id: str
    device_name: str
    priority: List[str]
    use_latest_per_device: bool = True


def _run(cmd: List[str], input_text: Optional[str] = None) -> str:
    p = subprocess.run(
        cmd,
        input=(input_text.encode("utf-8") if input_text is not None else None),
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        check=True,
    )
    return p.stdout.decode("utf-8", errors="replace")


def clipboard_get_text() -> str:
    sys = platform.system().lower()

    if sys == "darwin":
        return _run(["pbpaste"])
    if sys == "windows":
        return _run(["powershell", "-NoProfile", "-Command", "Get-Clipboard -Raw"])
    # Linux
    # Wayland
    if shutil_which("wl-paste"):
        return _run(["wl-paste", "--no-newline"])
    # X11
    if shutil_which("xclip"):
        return _run(["xclip", "-selection", "clipboard", "-o"])
    if shutil_which("xsel"):
        return _run(["xsel", "--clipboard", "--output"])
    raise RuntimeError("No clipboard tool found (need wl-paste/wl-copy or xclip/xsel).")


def clipboard_set_text(text: str) -> None:
    sys = platform.system().lower()

    if sys == "darwin":
        _run(["pbcopy"], input_text=text)
        return

    if sys == "windows":
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

    # Linux (X11)
    if shutil_which("xclip"):
        # Explicit stdin mode. This should return immediately after input is consumed.
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
        # Wait a short time only; if it hangs, kill it.
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


def shutil_which(cmd: str) -> Optional[str]:
    from shutil import which

    return which(cmd)


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
        return items[0]  # server already sorted newest first

    # Build map: device_id -> first (newest) occurrence
    newest_per_device: Dict[str, Dict[str, Any]] = {}
    for it in items:
        did = it.get("device_id", "")
        if did and did not in newest_per_device:
            newest_per_device[did] = it

    # Pick by priority order
    for did in cfg.priority:
        if did in newest_per_device:
            return newest_per_device[did]

    # Fallback: newest overall
    return items[0]


@app.command()
def send(config: str = typer.Option(CONFIG_PATH_DEFAULT, "--config")):
    cfg = load_config(config)

    try:
        sysname = platform.system().lower()
        if sysname == "windows":
            text = win_capture_selected_text_preserve_clipboard()
        else:
            text = clipboard_get_text()

        if not text:
            return

        payload = {
            "device_id": cfg.device_id,
            "device_name": cfg.device_name,
            "content_type": "text/plain",
            "content_text": text,
        }

        requests.post(
            f"{cfg.server_url}/v1/clip",
            json=payload,
            headers=auth_headers(cfg),
            timeout=2,
        )
    except Exception:
        pass


def x11_get_primary_text() -> str:
    # Reads highlighted selection without Ctrl+C
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
    # Types text into the focused window. No clipboard involved.
    subprocess.run(
        ["xdotool", "type", "--clearmodifiers", "--delay", "0", text],
        check=True,
    )


def paste_active_window():
    sysname = platform.system().lower()

    # Windows: Ctrl+V (needs no extra tools if you later implement SendInput)
    # For now, simplest is to rely on external automation, but we’ll keep Windows stub.
    if sysname == "windows":
        # TODO: implement SendInput with ctypes
        raise RuntimeError("paste_active_window not implemented for Windows yet")

    # macOS: Cmd+V (requires Accessibility permission if you use osascript)
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

    # Linux X11: xdotool Ctrl+V
    if shutil_which("xdotool"):
        subprocess.run(["xdotool", "key", "ctrl+v"], check=True)
        return

    raise RuntimeError("No paste method available. Install xdotool on Linux X11.")


@app.command()
def fetch(config: str = typer.Option(CONFIG_PATH_DEFAULT, "--config")):
    cfg = load_config(config)

    # Debug prints (safe to keep, can be removed later)
    typer.echo("fetch: requesting...")

    # Use explicit connect/read timeouts
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

    data = r.json()
    items = data.get("items", [])
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

    typer.echo("fetch: writing clipboard...")

    if platform.system().lower() == "windows":
        # paste without touching clipboard
        win_type_text(text)
        return

    # IMPORTANT: Wayland wl-copy can block if you wait on it.
    # Use communicate() to feed stdin and return.
    typer.echo("fetch: pasting without overwriting clipboard...")

    original = clipboard_get_text()

    clipboard_set_text(text)
    # tiny delay to ensure clipboard owner updates before paste
    time.sleep(0.03)

    paste_active_window()

    # tiny delay so the target app reads clipboard before we restore
    time.sleep(0.05)
    clipboard_set_text(original)

    typer.echo("fetch: done (clipboard preserved)")

    typer.echo(
        f"fetched clip_id={chosen.get('clip_id')} from={chosen.get('device_id')}"
    )


if __name__ == "__main__":
    app()
