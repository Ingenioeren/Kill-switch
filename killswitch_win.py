# killswitch_win.py
# Windows-only network kill switch with:
# - Tiny Tk GUI + LED indicator
# - TRUE global hotkey (works while gaming) via RegisterHotKey (no external deps)
# - Hotkey warning color (green=registered, red=failed)
# - System tray icon (requires: pystray + pillow)
# - UAC auto-elevation on launch (prompts for admin)
#
# Build (recommended for fast start + easy sharing):
#   pip install pyinstaller pystray pillow
#   pyinstaller --noconsole --onedir --name "Kill switch" killswitch_win.py

import ctypes
from ctypes import wintypes
import json
import os
import queue
import subprocess
import sys
import threading
import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
import webbrowser
from datetime import datetime
from time import perf_counter

APP_NAME = "Kill switch"
APP_VERSION = "1.0.1"
CONFIG_PATH = os.path.join(os.path.expanduser("~"), f".{APP_NAME.lower()}.json")
def resource_path(rel: str) -> str:
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, rel)


APP_ICON_PATH = resource_path("app.ico")

WM_HOTKEY = 0x0312
MUTEX_NAME = "Global\\KillSwitchSingleInstance"
SW_RESTORE = 9

MOD_ALT = 0x0001
MOD_CONTROL = 0x0002
MOD_SHIFT = 0x0004
MOD_WIN = 0x0008
MOD_NOREPEAT = 0x4000

user32 = ctypes.windll.user32
shell32 = ctypes.windll.shell32
kernel32 = ctypes.windll.kernel32

# Ensure correct pointer/int sizes for 64-bit Windows when dealing with WndProc hooks
LONG_PTR = ctypes.c_longlong if ctypes.sizeof(ctypes.c_void_p) == 8 else ctypes.c_long
LRESULT = LONG_PTR
WNDPROC = ctypes.WINFUNCTYPE(
    LRESULT,
    wintypes.HWND,
    wintypes.UINT,
    wintypes.WPARAM,
    wintypes.LPARAM,
)

user32.SetWindowLongPtrW.restype = LONG_PTR
user32.SetWindowLongPtrW.argtypes = [wintypes.HWND, ctypes.c_int, WNDPROC]
user32.GetWindowLongPtrW.restype = LONG_PTR
user32.GetWindowLongPtrW.argtypes = [wintypes.HWND, ctypes.c_int]
user32.CallWindowProcW.restype = LRESULT
user32.CallWindowProcW.argtypes = [WNDPROC, wintypes.HWND, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM]
user32.FindWindowW.argtypes = [wintypes.LPCWSTR, wintypes.LPCWSTR]
user32.FindWindowW.restype = wintypes.HWND
user32.ShowWindow.argtypes = [wintypes.HWND, ctypes.c_int]
user32.ShowWindow.restype = wintypes.BOOL
user32.SetForegroundWindow.argtypes = [wintypes.HWND]
user32.SetForegroundWindow.restype = wintypes.BOOL

kernel32.CreateMutexW.argtypes = [ctypes.c_void_p, wintypes.BOOL, wintypes.LPCWSTR]
kernel32.CreateMutexW.restype = wintypes.HANDLE
kernel32.ReleaseMutex.argtypes = [wintypes.HANDLE]
kernel32.ReleaseMutex.restype = wintypes.BOOL
kernel32.CloseHandle.argtypes = [wintypes.HANDLE]
kernel32.CloseHandle.restype = wintypes.BOOL

ERROR_ALREADY_EXISTS = 183
_single_instance_mutex = None

# Virtual-key codes (minimal set; extend as needed)
VK = {
    "A": 0x41, "B": 0x42, "C": 0x43, "D": 0x44, "E": 0x45, "F": 0x46, "G": 0x47,
    "H": 0x48, "I": 0x49, "J": 0x4A, "K": 0x4B, "L": 0x4C, "M": 0x4D, "N": 0x4E,
    "O": 0x4F, "P": 0x50, "Q": 0x51, "R": 0x52, "S": 0x53, "T": 0x54, "U": 0x55,
    "V": 0x56, "W": 0x57, "X": 0x58, "Y": 0x59, "Z": 0x5A,
    "F1": 0x70, "F2": 0x71, "F3": 0x72, "F4": 0x73, "F5": 0x74, "F6": 0x75,
    "F7": 0x76, "F8": 0x77, "F9": 0x78, "F10": 0x79, "F11": 0x7A, "F12": 0x7B,
    "PAUSE": 0x13, "INSERT": 0x2D, "DELETE": 0x2E, "HOME": 0x24, "END": 0x23,
    "PGUP": 0x21, "PGDN": 0x22,
}

# Default blocked keys (main keys) to avoid common in-game conflicts
BLOCKED_KEYS_DEFAULT = [
    "W", "A", "S", "D", "SPACE", "V", "Q", "R", "F", "M",
    "TAB", "CAPSLOCK", "INSERT", "DELETE",
    "F1", "F2", "F3",
]


def is_admin() -> bool:
    try:
        return bool(shell32.IsUserAnAdmin())
    except Exception:
        return False


def relaunch_as_admin_or_exit():
    """
    Relaunches the current script/exe with UAC prompt using 'runas'.
    If user cancels elevation, the current process exits.
    """
    # Prevent recursion if already elevated or explicitly marked
    if is_admin():
        return
    if "--no-elevate" in sys.argv:
        return

    # Build command line: keep args except internal marker
    args = [a for a in sys.argv[1:] if a != "--no-elevate"]
    params = " ".join([f'"{a}"' if " " in a else a for a in args])

    exe = sys.executable
    # If running as a frozen PyInstaller exe, sys.executable is the exe path already.
    # If running as python script, sys.executable is python.exe and we must pass script path.
    if getattr(sys, "frozen", False):
        # Running packaged exe
        file_to_run = exe
        params_to_run = params
    else:
        # Running from python: run python.exe "script.py" <args>
        file_to_run = exe
        script_path = os.path.abspath(sys.argv[0])
        params_to_run = f'"{script_path}" {params}'.strip()

    # ShellExecuteW returns > 32 on success, otherwise error code.
    rc = shell32.ShellExecuteW(None, "runas", file_to_run, params_to_run, None, 1)
    if rc <= 32:
        # User likely canceled or system policy denied
        sys.exit(1)
    sys.exit(0)


def show_existing_instance() -> bool:
    hwnd = user32.FindWindowW(None, APP_NAME)
    if not hwnd:
        return False
    try:
        user32.ShowWindow(hwnd, SW_RESTORE)
        user32.SetForegroundWindow(hwnd)
    except Exception:
        return False
    return True


def acquire_single_instance() -> bool:
    global _single_instance_mutex
    _single_instance_mutex = kernel32.CreateMutexW(None, False, MUTEX_NAME)
    if not _single_instance_mutex:
        return True
    if ctypes.GetLastError() == ERROR_ALREADY_EXISTS:
        show_existing_instance()
        try:
            kernel32.CloseHandle(_single_instance_mutex)
        except Exception:
            pass
        _single_instance_mutex = None
        return False
    return True


def release_single_instance():
    global _single_instance_mutex
    if _single_instance_mutex:
        try:
            kernel32.ReleaseMutex(_single_instance_mutex)
        except Exception:
            pass
        try:
            kernel32.CloseHandle(_single_instance_mutex)
        except Exception:
            pass
        _single_instance_mutex = None


def run(cmd: list[str]) -> subprocess.CompletedProcess:
    creationflags = 0
    startupinfo = None
    if os.name == "nt":
        creationflags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
        startupinfo = subprocess.STARTUPINFO()
        startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
    return subprocess.run(
        cmd,
        capture_output=True,
        text=True,
        shell=False,
        creationflags=creationflags,
        startupinfo=startupinfo,
    )


def run_powershell_json(ps_cmd: str):
    p = run(["powershell", "-NoProfile", "-Command", ps_cmd])
    if p.returncode != 0:
        raise RuntimeError((p.stdout + "\n" + p.stderr).strip())
    try:
        data = json.loads(p.stdout.strip() or "null")
    except Exception as e:
        raise RuntimeError(f"Failed to parse PowerShell JSON output: {e}")
    return data


def run_powershell(ps_cmd: str) -> subprocess.CompletedProcess:
    return run(["powershell", "-NoProfile", "-Command", ps_cmd])


def get_internet_connected_aliases() -> set[str]:
    """
    Returns interface aliases that Windows reports as having Internet connectivity
    (via Network Location Awareness).
    """
    ps = (
        "Get-NetConnectionProfile | "
        "Select-Object InterfaceAlias,IPv4Connectivity,IPv6Connectivity | "
        "ConvertTo-Json -Compress"
    )
    data = run_powershell_json(ps)
    if data is None:
        return set()
    if isinstance(data, dict):
        data = [data]
    out: set[str] = set()
    for row in data:
        if not isinstance(row, dict):
            continue
        alias = row.get("InterfaceAlias")
        v4 = str(row.get("IPv4Connectivity", "")).lower()
        v6 = str(row.get("IPv6Connectivity", "")).lower()
        if alias and ("internet" in v4 or "internet" in v6):
            out.add(alias)
    return out


def get_wired_wireless_aliases() -> set[str]:
    """
    Returns physical adapter aliases for Ethernet (802.3) and Wi-Fi (802.11).
    """
    ps = (
        "Get-NetAdapter | "
        "Select-Object InterfaceAlias,MediaType,NdisPhysicalMedium,Physical | "
        "ConvertTo-Json -Compress"
    )
    data = run_powershell_json(ps)
    if data is None:
        return set()
    if isinstance(data, dict):
        data = [data]
    out: set[str] = set()
    for row in data:
        if not isinstance(row, dict):
            continue
        alias = row.get("InterfaceAlias")
        if not alias:
            continue
        media = str(row.get("MediaType", "")).lower()
        ndis = str(row.get("NdisPhysicalMedium", "")).lower()
        physical = str(row.get("Physical", "")).lower() in ("true", "yes", "1")
        is_wired = "802.3" in media or "ethernet" in ndis
        is_wifi = "802.11" in media or "wireless" in ndis or "wlan" in ndis
        # Some adapters report Physical as empty; accept 802.3/802.11 regardless.
        if (is_wired or is_wifi) and (physical or str(row.get("Physical", "")).strip() == ""):
            out.add(alias)
        # Fallback: include common Wi-Fi naming even if Physical flags are weird
        if any(tok in str(alias).lower() for tok in ("wi-fi", "wifi", "wlan")):
            out.add(alias)
    return out


def get_adapter_kind_map() -> dict[str, str]:
    """
    Returns map of InterfaceAlias -> kind ("wifi"|"lan"|"other").
    Uses Get-NetAdapter metadata when available, falls back to name heuristics.
    """
    ps = (
        "Get-NetAdapter | "
        "Select-Object InterfaceAlias,MediaType,NdisPhysicalMedium,Physical | "
        "ConvertTo-Json -Compress"
    )
    kind_map: dict[str, str] = {}
    try:
        data = run_powershell_json(ps)
    except Exception:
        data = None
    if data is None:
        data = []
    if isinstance(data, dict):
        data = [data]
    for row in data:
        if not isinstance(row, dict):
            continue
        alias = row.get("InterfaceAlias")
        if not alias:
            continue
        media = str(row.get("MediaType", "")).lower()
        ndis = str(row.get("NdisPhysicalMedium", "")).lower()
        name_l = str(alias).lower()
        is_wifi = "802.11" in media or "wireless" in ndis or "wlan" in ndis or "wi-fi" in name_l or "wifi" in name_l
        is_lan = "802.3" in media or "ethernet" in ndis or "lan" in name_l
        if is_wifi:
            kind_map[alias] = "wifi"
        elif is_lan:
            kind_map[alias] = "lan"
        else:
            kind_map[alias] = "other"
    return kind_map


def list_interfaces() -> list[dict]:
    p = run(["netsh", "interface", "show", "interface"])
    if p.returncode != 0:
        raise RuntimeError((p.stdout + "\n" + p.stderr).strip())

    lines = [ln.rstrip() for ln in p.stdout.splitlines() if ln.strip()]
    sep_idx = None
    for i, ln in enumerate(lines):
        if set(ln.strip()) <= {"-"} and len(ln.strip()) >= 3:
            sep_idx = i
            break
    if sep_idx is None or sep_idx + 1 >= len(lines):
        return []

    rows = lines[sep_idx + 1 :]
    out = []
    for r in rows:
        parts = r.split()
        if len(parts) < 4:
            continue
        out.append(
            {
                "admin_state": parts[0],  # Enabled/Disabled
                "state": parts[1],        # Connected/Disconnected
                "type": parts[2],         # Dedicated/Loopback/...
                "name": " ".join(parts[3:]),
            }
        )
    return out


def list_running_processes() -> list[str]:
    ps = (
        "Get-Process | "
        "Where-Object { $_.MainWindowTitle -and $_.MainWindowTitle.Trim().Length -gt 0 } | "
        "Select-Object -ExpandProperty ProcessName | "
        "Sort-Object -Unique | "
        "ConvertTo-Json -Compress"
    )
    data = run_powershell_json(ps)
    if data is None:
        return []
    if isinstance(data, list):
        return [str(x) for x in data if x]
    if isinstance(data, str):
        return [data]
    return []


def show_splash(root: tk.Tk):
    splash = tk.Toplevel(root)
    splash.overrideredirect(True)
    splash.resizable(False, False)
    try:
        if os.path.exists(APP_ICON_PATH):
            splash.iconbitmap(APP_ICON_PATH)
    except Exception:
        pass

    width, height = 360, 180
    x = (splash.winfo_screenwidth() - width) // 2
    y = (splash.winfo_screenheight() - height) // 2
    splash.geometry(f"{width}x{height}+{x}+{y}")

    frm = tk.Frame(splash, padx=16, pady=16)
    frm.pack(fill="both", expand=True)
    tk.Label(frm, text=APP_NAME, font=("Segoe UI", 16, "bold")).pack(pady=(8, 6))
    tk.Label(frm, text="Loading...", font=("Segoe UI", 10)).pack(pady=(0, 10))

    pb = ttk.Progressbar(frm, orient="horizontal", length=260, mode="indeterminate")
    pb.pack(pady=(0, 6))
    pb.start(10)

    splash.update_idletasks()
    splash.update()
    return splash, pb


def set_interface_admin(name: str, enabled: bool) -> None:
    admin = "enabled" if enabled else "disabled"
    if not is_admin():
        raise RuntimeError("Administrator rights are required to toggle network adapters.\n\nRight-click and 'Run as administrator', or start the EXE from an elevated prompt.")
    # Quote interface name to survive spaces, dashes, and locale-specific characters.
    p = run(["netsh", "interface", "set", "interface", f'name="{name}"', f"admin={admin}"])
    if p.returncode != 0:
        # Some locales/netsh builds return "No more data is available" even though it succeeded.
        # Re-check the interface state; if already in the desired state, treat as success.
        try:
            ifaces = list_interfaces()
            for i in ifaces:
                if i["name"].lower() == name.lower() and i["admin_state"].lower() == admin:
                    return
        except Exception:
            pass

        msg = (p.stdout + "\n" + p.stderr).strip()
        raise RuntimeError(f'Failed to set "{name}" admin={admin}.\n{msg}')


def set_adapters_admin_ps(names: list[str], enabled: bool) -> bool:
    if not names:
        return True
    if not is_admin():
        raise RuntimeError(
            "Administrator rights are required to toggle network adapters.\n\n"
            "Right-click and 'Run as administrator', or start the EXE from an elevated prompt."
        )
    action = "Enable" if enabled else "Disable"
    quoted = ",".join([f'"{n}"' for n in names])
    ps = f"{action}-NetAdapter -Name {quoted} -Confirm:$false"
    p = run_powershell(ps)
    return p.returncode == 0


def load_config() -> dict:
    if not os.path.exists(CONFIG_PATH):
        return {}
    try:
        with open(CONFIG_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def save_config(cfg: dict) -> None:
    try:
        with open(CONFIG_PATH, "w", encoding="utf-8") as f:
            json.dump(cfg, f, indent=2)
    except Exception:
        pass


def parse_hotkey(s: str) -> tuple[int, int, str]:
    parts = [p.strip().upper() for p in (s or "").split("+") if p.strip()]
    if len(parts) < 2:
        raise ValueError("Hotkey must include at least one modifier and one key, e.g. CTRL+ALT+K.")

    mods = 0
    key_name = None

    for p in parts:
        if p in ("CTRL", "CONTROL"):
            mods |= MOD_CONTROL
        elif p == "ALT":
            mods |= MOD_ALT
        elif p == "SHIFT":
            mods |= MOD_SHIFT
        elif p in ("WIN", "WINDOWS"):
            mods |= MOD_WIN
        else:
            key_name = p

    if mods == 0:
        raise ValueError("Hotkey must include a modifier (CTRL/ALT/SHIFT/WIN).")
    if not key_name:
        raise ValueError("Hotkey missing main key (e.g. K, F12).")

    if key_name in VK:
        vk = VK[key_name]
    elif len(key_name) == 1 and "A" <= key_name <= "Z":
        vk = ord(key_name)
    else:
        raise ValueError(f"Unsupported key: {key_name}. Use A-Z, F1-F12, or e.g. PAUSE/INSERT/DELETE.")

    mods |= MOD_NOREPEAT
    return mods, vk, key_name


class KillSwitchApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title(APP_NAME)
        self.root.resizable(False, False)
        try:
            if os.path.exists(APP_ICON_PATH):
                self.root.iconbitmap(APP_ICON_PATH)
        except Exception:
            pass

        self.cfg = load_config()
        self.killed = False
        self.enabled_before_kill: list[str] = self.cfg.get("enabled_before_kill", [])
        # Hotkeys: separate disable/enable
        if "hotkey" in self.cfg and "hotkey_disable" not in self.cfg:
            # Migrate legacy single hotkey to disable hotkey
            self.cfg["hotkey_disable"] = self.cfg.get("hotkey")
            save_config(self.cfg)
        self.hotkey_disable_str = self.cfg.get("hotkey_disable", "CTRL+ALT+K")
        self.hotkey_enable_str = self.cfg.get("hotkey_enable", "CTRL+ALT+E")
        self.log_path = os.path.join(os.path.expanduser("~"), f".{APP_NAME.lower()}.log")
        self.selected_adapters = set(self.cfg.get("selected_adapters", []))
        self.selected_adapters_from_config = "selected_adapters" in self.cfg
        self.adapter_vars: dict[str, tk.BooleanVar] = {}
        self.adapter_checks: dict[str, tk.Checkbutton] = {}
        self.adapter_kind_map: dict[str, str] = {}

        self.kill_program_name = self.cfg.get("kill_program_name", "")
        self.kill_on_disable = bool(self.cfg.get("kill_on_disable", False))
        self.kill_mode_str = self.cfg.get("kill_mode", "both")
        self.process_names: list[str] = []
        if "blocked_keys" not in self.cfg:
            self.cfg["blocked_keys"] = BLOCKED_KEYS_DEFAULT[:]
            save_config(self.cfg)
        self.blocked_keys = [str(k).upper() for k in self.cfg.get("blocked_keys", [])]

        self.hotkey_disable_id = 1
        self.hotkey_enable_id = 2
        self.hotkey_disable_registered = False
        self.hotkey_enable_registered = False
        self.hotkey_disable_failed = False
        self.hotkey_enable_failed = False

        self.status_var = tk.StringVar(value="Kill switch: ?")
        self.hk_disable_var = tk.StringVar(value=self.hotkey_disable_str)
        self.hk_disable_state_var = tk.StringVar(value="...")
        self.hk_enable_var = tk.StringVar(value=self.hotkey_enable_str)
        self.hk_enable_state_var = tk.StringVar(value="...")
        self.ready_var = tk.StringVar(value="Ready: ?")
        self.log_lines: list[str] = []
        self.kill_name_var = tk.StringVar(value=self.kill_program_name)
        self.kill_on_disable_var = tk.BooleanVar(value=self.kill_on_disable)
        self.kill_mode_var = tk.StringVar(value=self.kill_mode_str)

        # Tray related
        self.tray_icon = None
        self.tray_available = False
        self._tray_thread = None
        self._ui_queue: queue.Queue = queue.Queue()
        self._ui_queue_poll_ms = 50

        notebook = ttk.Notebook(root)
        notebook.pack(fill="both", expand=True)

        program_tab = tk.Frame(notebook)
        about_tab = tk.Frame(notebook)
        notebook.add(program_tab, text="  Kill Switch  ")
        notebook.add(about_tab, text="  About  ")

        # == About Tab ==
        about_frame = tk.Frame(about_tab, padx=24, pady=20)
        about_frame.pack(fill="both", expand=True, anchor="nw")

        # Header
        tk.Label(about_frame, text=APP_NAME, font=("Segoe UI", 18, "bold")).pack(anchor="w", pady=(0, 2))
        tk.Label(about_frame, text=f"Version {APP_VERSION}  —  Windows Network Kill Switch",
            font=("Segoe UI", 9), fg="#666").pack(anchor="w", pady=(0, 2))

        tk.Frame(about_frame, height=1, bg="#dddddd").pack(fill="x", pady=(8, 12))

        # Author row
        author_row = tk.Frame(about_frame)
        author_row.pack(anchor="w", pady=(0, 2))
        tk.Label(author_row, text="Made by ", font=("Segoe UI", 10)).pack(side="left")
        tk.Label(author_row, text="Ingenioeren", font=("Segoe UI", 10, "bold")).pack(side="left")

        link = tk.Label(about_frame, text="github.com/Ingenioeren",
            font=("Segoe UI", 9), fg="#1565c0", cursor="hand2")
        link.pack(anchor="w", pady=(2, 14))
        link.bind("<Button-1>", lambda _e: webbrowser.open("https://github.com/Ingenioeren"))

        tk.Frame(about_frame, height=1, bg="#dddddd").pack(fill="x", pady=(0, 12))

        # Quick-start section
        tk.Label(about_frame, text="Quick Start", font=("Segoe UI", 10, "bold")).pack(anchor="w", pady=(0, 4))
        steps = [
            "1.  Tick the network adapters you want to disable on the right panel.",
            "2.  Choose a Kill Mode — Internet, Program, or Both.",
            "3.  Press the big button (or the Disable hotkey) to activate.",
            "4.  Press again (or the Enable hotkey) to restore your connection.",
        ]
        for s in steps:
            tk.Label(about_frame, text=s, font=("Segoe UI", 9), justify="left").pack(anchor="w", pady=1)

        tk.Frame(about_frame, height=1, bg="#dddddd").pack(fill="x", pady=(10, 10))

        # Notes section
        tk.Label(about_frame, text="Notes", font=("Segoe UI", 10, "bold")).pack(anchor="w", pady=(0, 4))
        notes = [
            "•  Requires administrator rights to enable/disable adapters.",
            "•  Re-enabling the kill switch will NOT relaunch killed programs.",
            "•  The last selected program is remembered between sessions.",
            "•  Avoid hotkey combos that conflict with your game or other apps.",
            "•  Use Ctrl+Alt+K to disable and Ctrl+Alt+E to re-enable (default).",
        ]
        for n in notes:
            tk.Label(about_frame, text=n, font=("Segoe UI", 9), justify="left", fg="#333").pack(anchor="w", pady=1)

        tk.Frame(about_frame, height=1, bg="#dddddd").pack(fill="x", pady=(10, 10))

        tk.Label(about_frame, text="Use at your own risk.",
            font=("Segoe UI", 8), fg="#999").pack(anchor="w")

        # == Kill Switch Tab ==
        # Button style shared across all secondary buttons
        _BTN = dict(relief="raised", bd=2, cursor="hand2",
                    bg="#37474f", fg="white",
                    activebackground="#263238", activeforeground="white")

        # Log anchored to bottom first so it never gets squeezed out
        self.log_frame = tk.LabelFrame(program_tab, text=" Log ", font=("Segoe UI", 9, "bold"),
            padx=8, pady=6)
        self.log_frame.pack(side="bottom", fill="x", padx=14, pady=(0, 14))
        self.log_text = tk.Text(self.log_frame, height=5, state="disabled",
            wrap="none", font=("Consolas", 8))
        self.log_scroll = tk.Scrollbar(self.log_frame, orient="vertical", command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=self.log_scroll.set)
        self.log_text.pack(side="left", fill="both", expand=True)
        self.log_scroll.pack(side="right", fill="y")

        # Main area: left controls | right adapters
        tab_frame = tk.Frame(program_tab)
        tab_frame.pack(side="top", fill="both", expand=True)

        left = tk.Frame(tab_frame, padx=14, pady=14)
        left.pack(side="left", fill="both", expand=True)

        tk.Frame(tab_frame, width=1, bg="#cccccc").pack(side="left", fill="y", pady=8)

        right = tk.Frame(tab_frame, padx=10, pady=14)
        right.pack(side="right", fill="y")

        # -- Status Row --
        status_row = tk.Frame(left)
        status_row.pack(fill="x", pady=(0, 10))
        self.status_box = tk.Label(status_row, text="Inactive", width=10,
            font=("Segoe UI", 9, "bold"), fg="white", bg="#1565c0", padx=10, pady=5)
        self.status_box.pack(side="left")
        self.ready_box = tk.Label(status_row, text="Not Ready", width=10,
            font=("Segoe UI", 9), fg="white", bg="#9e9e9e", padx=10, pady=5)
        self.ready_box.pack(side="left", padx=(6, 0))

        # -- Big Toggle Button --
        self.toggle_btn = tk.Button(left, text="ACTIVATE KILL SWITCH",
            command=self.on_toggle_clicked,
            font=("Segoe UI", 12, "bold"),
            bg="#c62828", fg="white",
            activebackground="#b71c1c", activeforeground="white",
            relief="raised", bd=3, cursor="hand2", pady=12)
        self.toggle_btn.pack(fill="x", pady=(0, 14))

        # -- Kill Mode --
        mode_frame = tk.LabelFrame(left, text=" Kill Mode ", font=("Segoe UI", 9, "bold"),
            padx=12, pady=8)
        mode_frame.pack(fill="x", pady=(0, 10))
        for val, label in [
            ("internet", "Internet only  —  disable network adapters"),
            ("program",  "Program only   —  terminate selected process"),
            ("both",     "Both           —  disable adapters + kill process"),
        ]:
            tk.Radiobutton(mode_frame, text=label, variable=self.kill_mode_var,
                value=val, font=("Segoe UI", 9),
                command=self._on_kill_mode_change).pack(anchor="w", pady=2)

        # -- Program --
        self.prog_frame = tk.LabelFrame(left, text=" Program ", font=("Segoe UI", 9, "bold"),
            padx=12, pady=8)
        self.prog_frame.pack(fill="x", pady=(0, 10))
        prog_row = tk.Frame(self.prog_frame)
        prog_row.pack(fill="x")
        tk.Label(prog_row, text="Process:", font=("Segoe UI", 9), width=8, anchor="w").pack(side="left")
        self.kill_combo = ttk.Combobox(prog_row, textvariable=self.kill_name_var, width=22, state="readonly")
        self.kill_combo.pack(side="left", padx=(4, 6))
        self.kill_combo.bind("<<ComboboxSelected>>", self.on_program_selection)
        self.kill_btn = tk.Button(prog_row, text="Kill Now", font=("Segoe UI", 9),
            padx=8, pady=3, command=self.on_kill_now, **_BTN)
        self.kill_btn.pack(side="left")
        self.kill_on_disable_chk = tk.Checkbutton(
            self.prog_frame,
            text="Auto-kill when kill switch activates  (Both mode)",
            variable=self.kill_on_disable_var,
            font=("Segoe UI", 9),
            command=self.on_kill_on_disable_toggle,
        )
        self.kill_on_disable_chk.pack(anchor="w", pady=(6, 0))

        # -- Hotkeys --
        hk_frame = tk.LabelFrame(left, text=" Hotkeys ", font=("Segoe UI", 9, "bold"),
            padx=12, pady=8)
        hk_frame.pack(fill="x", pady=(0, 10))
        hk_inner = tk.Frame(hk_frame)
        hk_inner.pack(fill="x")
        hk_inner.columnconfigure(1, weight=1)

        tk.Label(hk_inner, text="Disable:", font=("Segoe UI", 9), width=8, anchor="w").grid(row=0, column=0, sticky="w")
        tk.Label(hk_inner, textvariable=self.hk_disable_var, font=("Segoe UI", 9, "bold"), anchor="w").grid(row=0, column=1, sticky="w", padx=(4, 4))
        self.hk_disable_state_label = tk.Label(hk_inner, textvariable=self.hk_disable_state_var,
            font=("Segoe UI", 9), width=7, anchor="w")
        self.hk_disable_state_label.grid(row=0, column=2, sticky="w")
        tk.Button(hk_inner, text="Set…", font=("Segoe UI", 8), padx=6, pady=2,
            command=self.on_set_disable_hotkey, **_BTN).grid(row=0, column=3, padx=(6, 0))

        tk.Label(hk_inner, text="Enable:", font=("Segoe UI", 9), width=8, anchor="w").grid(row=1, column=0, sticky="w", pady=(4, 0))
        tk.Label(hk_inner, textvariable=self.hk_enable_var, font=("Segoe UI", 9, "bold"), anchor="w").grid(row=1, column=1, sticky="w", padx=(4, 4), pady=(4, 0))
        self.hk_enable_state_label = tk.Label(hk_inner, textvariable=self.hk_enable_state_var,
            font=("Segoe UI", 9), width=7, anchor="w")
        self.hk_enable_state_label.grid(row=1, column=2, sticky="w", pady=(4, 0))
        tk.Button(hk_inner, text="Set…", font=("Segoe UI", 8), padx=6, pady=2,
            command=self.on_set_enable_hotkey, **_BTN).grid(row=1, column=3, padx=(6, 0), pady=(4, 0))

        # -- Bottom Buttons --
        btn_row = tk.Frame(left)
        btn_row.pack(fill="x", pady=(0, 4))
        self.hide_btn = tk.Button(btn_row, text="Hide to Tray", font=("Segoe UI", 9),
            padx=10, pady=4, command=self.hide_to_tray, **_BTN)
        self.hide_btn.pack(side="left")
        self.refresh_adapters_btn = tk.Button(btn_row, text="Refresh Adapters",
            font=("Segoe UI", 9), padx=10, pady=4,
            command=self.refresh_adapters_only, **_BTN)
        self.refresh_adapters_btn.pack(side="left", padx=(6, 0))
        self.refresh_programs_btn = tk.Button(btn_row, text="Refresh Programs",
            font=("Segoe UI", 9), padx=10, pady=4,
            command=self.refresh_programs_only, **_BTN)
        self.refresh_programs_btn.pack(side="left", padx=(6, 0))

        # -- Adapter Frame (right panel, clearly separated) --
        self.adapter_frame = tk.LabelFrame(right, text=" Network Adapters ",
            font=("Segoe UI", 9, "bold"), padx=8, pady=6)
        self.adapter_frame.pack(fill="both", expand=True)

        self.refresh_status()
        self.root.after(100, self._on_kill_mode_change)  # apply initial mode state after widgets exist
        self.register_hotkey_from_string(self.hotkey_disable_str, self.hotkey_disable_id, silent_fail=False)
        self.register_hotkey_from_string(self.hotkey_enable_str, self.hotkey_enable_id, silent_fail=False)
        self.hook_wm_hotkey()
        self.init_tray_icon()
        self.root.after(0, self._process_ui_queue)

        # Close button => minimize to tray (if available), else exit
        self.root.protocol("WM_DELETE_WINDOW", self.on_close_clicked)

    def set_status_box(self, active: bool):
        if active:
            self.status_box.configure(text="Active", fg="white", bg="#c62828")
        else:
            self.status_box.configure(text="Inactive", fg="white", bg="#1565c0")
        self._update_toggle_btn()
        self.update_tray_icon_image()

    def set_hotkey_ui_state(self, which: str, ok: bool, hk: str):
        if which == "disable":
            self.hk_disable_var.set(hk)
            label = self.hk_disable_state_label
            var = self.hk_disable_state_var
            self.hotkey_disable_failed = not ok
        else:
            self.hk_enable_var.set(hk)
            label = self.hk_enable_state_label
            var = self.hk_enable_state_var
            self.hotkey_enable_failed = not ok

        if ok:
            var.set("OK")
            label.configure(fg="green")
        else:
            var.set("FAILED")
            label.configure(fg="red")

        self.set_ready_state(self.hotkey_disable_registered and self.hotkey_enable_registered)

    def set_ready_state(self, ready: bool):
        if ready:
            self.ready_box.configure(text="Ready", fg="white", bg="#2e7d32")
        else:
            self.ready_box.configure(text="Not Ready", fg="white", bg="#9e9e9e")

    def get_selected_adapters(self) -> set[str]:
        return {name for name, var in self.adapter_vars.items() if var.get()}

    def sort_adapter_names(self, names: list[str]) -> list[str]:
        def rank(name: str) -> int:
            kind = self.adapter_kind_map.get(name, "other")
            if kind == "wifi":
                return 0
            if kind == "lan":
                return 1
            return 2

        return sorted(names, key=lambda n: (rank(n), n.lower()))

    def on_adapter_toggle(self):
        selected = sorted(self.get_selected_adapters())
        self.cfg["selected_adapters"] = selected
        self.selected_adapters = set(selected)
        self.selected_adapters_from_config = True
        save_config(self.cfg)

    def on_kill_on_disable_toggle(self):
        self.kill_on_disable = bool(self.kill_on_disable_var.get())
        self.cfg["kill_on_disable"] = self.kill_on_disable
        save_config(self.cfg)

    def _on_kill_mode_change(self):
        mode = self.kill_mode_var.get()
        self.cfg["kill_mode"] = mode
        save_config(self.cfg)
        prog_on = mode in ("program", "both")
        net_on = mode in ("internet", "both")
        self.kill_combo.configure(state="readonly" if prog_on else "disabled")
        self.kill_btn.configure(state="normal" if prog_on else "disabled")
        # kill_on_disable checkbox only meaningful in "both" mode
        self.kill_on_disable_chk.configure(state="normal" if mode == "both" else "disabled")
        for cb in self.adapter_checks.values():
            cb.configure(state="normal" if net_on else "disabled")
        self.refresh_adapters_btn.configure(state="normal" if net_on else "disabled")
        self._update_toggle_btn()

    def _update_toggle_btn(self):
        mode = self.kill_mode_var.get() if hasattr(self, "kill_mode_var") else "both"
        if self.killed and mode != "program":
            self.toggle_btn.configure(
                text="RESTORE NETWORK",
                bg="#2e7d32", activebackground="#1b5e20",
            )
        else:
            labels = {
                "internet": "DISABLE INTERNET",
                "program":  "KILL PROGRAM",
                "both":     "ACTIVATE KILL SWITCH",
            }
            self.toggle_btn.configure(
                text=labels.get(mode, "ACTIVATE KILL SWITCH"),
                bg="#c62828", activebackground="#b71c1c",
            )

    def kill_program(self, name: str) -> bool:
        name = (name or "").strip()
        if not name:
            return False
        # Accept either exe name or process name without .exe
        if not name.lower().endswith(".exe"):
            name = f"{name}.exe"
        p = run(["taskkill", "/IM", name, "/F"])
        ok = p.returncode == 0
        if ok:
            self.log(f"Killed: {name}")
        else:
            msg = (p.stdout + "\n" + p.stderr).strip()
            self.log(f"Kill failed: {name} ({msg})")
        return ok

    def on_kill_now(self):
        name = self.kill_name_var.get().strip()
        self.kill_program_name = name
        self.cfg["kill_program_name"] = name
        save_config(self.cfg)
        self.kill_program(name)

    def refresh_process_list(self):
        try:
            names = list_running_processes()
        except Exception:
            names = []
        # Always keep the saved program in the list even if it's not currently running
        saved = self.kill_program_name
        if saved and saved not in names:
            names = [saved] + names
        names = [""] + names
        self.process_names = names
        def apply():
            self.kill_combo["values"] = self.process_names
            current = self.kill_name_var.get()
            if current in self.process_names:
                self.kill_combo.set(current)
            elif saved in self.process_names:
                self.kill_combo.set(saved)
            else:
                self.kill_combo.set("")
            self.on_program_selection()
        self.ui_call(apply)

    def on_program_selection(self, _event=None):
        name = self.kill_name_var.get().strip()
        if name:
            self.kill_program_name = name
            self.cfg["kill_program_name"] = name
            save_config(self.cfg)
        else:
            if self.kill_on_disable_var.get():
                self.kill_on_disable_var.set(False)
                self.kill_on_disable = False
                self.cfg["kill_on_disable"] = False
                save_config(self.cfg)

    def update_adapter_list(self, names: list[str], default_set: set[str], connected_set: set[str]):
        names_set = set(names)
        for name in list(self.adapter_vars.keys()):
            if name not in names_set:
                try:
                    self.adapter_checks[name].destroy()
                except Exception:
                    pass
                self.adapter_checks.pop(name, None)
                self.adapter_vars.pop(name, None)

        for name in self.sort_adapter_names(names):
            if name in self.adapter_vars:
                continue
            if self.selected_adapters_from_config:
                initial = name in self.selected_adapters
            else:
                initial = name in default_set
            var = tk.BooleanVar(value=initial)
            cb = tk.Checkbutton(self.adapter_frame, text=name, variable=var, command=self.on_adapter_toggle)
            cb.pack(anchor="w")
            self.adapter_vars[name] = var
            self.adapter_checks[name] = cb

    def ui_call(self, fn):
        self._ui_queue.put(fn)

    def log(self, msg: str):
        ts = datetime.now().strftime("%H:%M:%S")
        line = f"[{ts}] {msg}"
        try:
            with open(self.log_path, "a", encoding="utf-8") as f:
                f.write(line + "\n")
        except Exception:
            pass
        try:
            print(line, flush=True)
        except Exception:
            pass
        def append():
            try:
                self.log_text.configure(state="normal")
                self.log_text.insert("end", line + "\n")
                self.log_text.see("end")
                self.log_text.configure(state="disabled")
            except Exception:
                pass
        self.ui_call(append)

    def _process_ui_queue(self):
        try:
            while True:
                fn = self._ui_queue.get_nowait()
                try:
                    fn()
                except Exception:
                    pass
        except queue.Empty:
            pass

        try:
            self.root.after(self._ui_queue_poll_ms, self._process_ui_queue)
        except Exception:
            pass

    def refresh_status(self, refresh_programs: bool = True, log_summary: bool = True):
        def work():
            try:
                if log_summary:
                    self.log("Checking connections...")
                ifaces = list_interfaces()
                try:
                    total = len([i for i in ifaces if i["type"].lower() != "loopback"])
                except Exception:
                    total = len(ifaces)
                if log_summary:
                    self.log(f"Found {total} network adapter(s).")
                enabled = [
                    i for i in ifaces
                    if i["admin_state"].lower() == "enabled"
                    and i["type"].lower() != "loopback"
                ]
                try:
                    self.adapter_kind_map = get_adapter_kind_map()
                except Exception:
                    self.adapter_kind_map = {}
                disabled = [
                    i for i in ifaces
                    if i["admin_state"].lower() == "disabled"
                    and i["type"].lower() != "loopback"
                ]
                # Only treat as "killed" if this app previously disabled adapters.
                if self.killed:
                    enabled_now = {i["name"] for i in enabled}
                    killed_now = len(self.enabled_before_kill) > 0 and all(
                        name not in enabled_now for name in self.enabled_before_kill
                    )
                    self.killed = killed_now

                default_set = {i["name"] for i in enabled}
                names = [i["name"] for i in ifaces if i["type"].lower() != "loopback"]
                connected_set = {i["name"] for i in ifaces if i["state"].lower() == "connected"}
                self.ui_call(lambda: self.update_adapter_list(names, default_set, connected_set))

                self.ui_call(lambda: self.status_var.set("Kill switch: ON" if self.killed else "Kill switch: OFF"))
                self.ui_call(lambda: self.set_status_box(self.killed))
                if refresh_programs:
                    self.refresh_process_list()
            except Exception as e:
                self.log(f"Status check failed: {e}")
                self.ui_call(lambda: self.status_var.set("Kill switch: ERR"))
                self.ui_call(lambda: self.set_status_box(False))
                self.ui_call(lambda: messagebox.showerror("Error", str(e)))

        threading.Thread(target=work, daemon=True).start()

    def refresh_adapters_only(self):
        self.log("Updating network adapters...")
        self.refresh_status(refresh_programs=False, log_summary=True)

    def refresh_programs_only(self):
        self.log("Updating programs...")
        self.refresh_process_list()

    def disable_network(self):
        mode = self.kill_mode_var.get()
        do_net = mode in ("internet", "both")
        do_prog = mode == "program" or (mode == "both" and self.kill_on_disable_var.get())

        if do_net:
            self.log("Disabling selected network adapters...")
            ifaces = list_interfaces()
            selected = self.get_selected_adapters()
            if not selected:
                self.log("No network adapters selected.")
                try:
                    self.ui_call(lambda: messagebox.showwarning(
                        "No network adapters selected",
                        "Select at least one network adapter in the adapter list first."
                    ))
                except Exception:
                    pass
                if not do_prog:
                    self.killed = False
                    return
            else:
                to_disable = [
                    i["name"]
                    for i in ifaces
                    if i["type"].lower() != "loopback"
                    and i["admin_state"].lower() == "enabled"
                    and (i["name"] in selected)
                ]
                self.log(f"Disabling {len(to_disable)} network adapter(s)...")
                if not to_disable:
                    self.log("No enabled network adapters matched.")
                    try:
                        self.ui_call(lambda: messagebox.showwarning(
                            "No network adapters matched",
                            "No enabled network adapters were found to disable.\n"
                            "Check adapter names and status, then try again."
                        ))
                    except Exception:
                        pass
                    if not do_prog:
                        self.killed = False
                        return
                else:
                    self.enabled_before_kill = to_disable[:]
                    self.cfg["enabled_before_kill"] = self.enabled_before_kill
                    save_config(self.cfg)
                    start = perf_counter()
                    ok = set_adapters_admin_ps(to_disable, enabled=False)
                    if not ok:
                        for name in to_disable:
                            self.log(f"Disabling: {name}")
                            set_interface_admin(name, enabled=False)
                    elapsed_ms = (perf_counter() - start) * 1000.0
                    self.log(f"Disable time: {elapsed_ms:.0f} ms")
                    self.killed = True

        if do_prog:
            name = self.kill_name_var.get().strip()
            if name:
                try:
                    self.kill_program(name)
                    self.killed = True
                except Exception:
                    pass
            elif mode == "program":
                self.log("No program selected.")
                try:
                    self.ui_call(lambda: messagebox.showwarning(
                        "No program selected",
                        "Select a program from the Process list first."
                    ))
                except Exception:
                    pass
                self.killed = False
                return

    def enable_network(self):
        self.log("Enabling network adapters...")
        if not self.enabled_before_kill:
            ifaces = list_interfaces()
            self.enabled_before_kill = [i["name"] for i in ifaces if i["type"].lower() != "loopback"]

        start = perf_counter()
        ok = set_adapters_admin_ps(self.enabled_before_kill, enabled=True)
        if not ok:
            for name in self.enabled_before_kill:
                try:
                    self.log(f"Enabling {name}...")
                    set_interface_admin(name, enabled=True)
                except Exception:
                    pass
        elapsed_ms = (perf_counter() - start) * 1000.0
        self.log(f"Enable time: {elapsed_ms:.0f} ms")
        self.killed = False

    def on_toggle_clicked(self):
        def work():
            try:
                mode = self.kill_mode_var.get()
                self.log("Toggle clicked.")
                if self.killed and mode != "program":
                    self.enable_network()
                else:
                    self.disable_network()
                self.refresh_status(log_summary=False)
            except Exception as e:
                self.log(f"Toggle error: {e}")
                self.ui_call(lambda: messagebox.showerror("Error", str(e)))
                self.refresh_status(log_summary=False)

        threading.Thread(target=work, daemon=True).start()

    def run_disable(self):
        def work():
            try:
                self.log("Disable hotkey pressed.")
                self.disable_network()
                self.refresh_status(log_summary=False)
            except Exception as e:
                self.log(f"Disable error: {e}")
                self.ui_call(lambda: messagebox.showerror("Error", str(e)))
                self.refresh_status(log_summary=False)

        threading.Thread(target=work, daemon=True).start()

    def run_enable(self):
        def work():
            try:
                self.log("Enable hotkey pressed.")
                if self.kill_mode_var.get() != "program":
                    self.enable_network()
                self.refresh_status(log_summary=False)
            except Exception as e:
                self.log(f"Enable error: {e}")
                self.ui_call(lambda: messagebox.showerror("Error", str(e)))
                self.refresh_status(log_summary=False)

        threading.Thread(target=work, daemon=True).start()

    # ---- Global hotkey ----
    def unregister_hotkey(self, hotkey_id: int):
        if hotkey_id == self.hotkey_disable_id and self.hotkey_disable_registered:
            try:
                hwnd = self.root.winfo_id()
                user32.UnregisterHotKey(hwnd, hotkey_id)
            except Exception:
                pass
            self.hotkey_disable_registered = False
            self.hotkey_disable_failed = False
        elif hotkey_id == self.hotkey_enable_id and self.hotkey_enable_registered:
            try:
                hwnd = self.root.winfo_id()
                user32.UnregisterHotKey(hwnd, hotkey_id)
            except Exception:
                pass
            self.hotkey_enable_registered = False
            self.hotkey_enable_failed = False
        self.set_ready_state(self.hotkey_disable_registered and self.hotkey_enable_registered)

    def register_hotkey_from_string(self, hk: str, hotkey_id: int, silent_fail: bool):
        try:
            mods, vk, key_name = parse_hotkey(hk)
        except Exception as e:
            which = "disable" if hotkey_id == self.hotkey_disable_id else "enable"
            self.ui_call(lambda: self.set_hotkey_ui_state(which, False, hk))
            if not silent_fail:
                messagebox.showerror("Hotkey", str(e))
            return False

        blocked = {k.upper() for k in self.blocked_keys}
        if key_name in blocked:
            msg = (
                f'"{hk}" uses a blocked key "{key_name}".\n'
                f"Choose a different key or edit blocked keys in:\n{CONFIG_PATH}"
            )
            which = "disable" if hotkey_id == self.hotkey_disable_id else "enable"
            self.ui_call(lambda: self.set_hotkey_ui_state(which, False, hk))
            if not silent_fail:
                messagebox.showerror("Hotkey", msg)
            return False

        self.unregister_hotkey(hotkey_id)

        hwnd = self.root.winfo_id()
        ok = bool(user32.RegisterHotKey(hwnd, hotkey_id, mods, vk))
        which = "disable" if hotkey_id == self.hotkey_disable_id else "enable"
        self.ui_call(lambda: self.set_hotkey_ui_state(which, ok, hk))

        if not ok and not silent_fail:
            messagebox.showerror(
                "Hotkey",
                f"Failed to register hotkey: {hk}\n\n"
                "It may already be used by another application.\n"
                "Try a different combo (e.g. CTRL+SHIFT+F12)."
            )
            return False

        if ok:
            if hotkey_id == self.hotkey_disable_id:
                self.hotkey_disable_registered = True
                self.hotkey_disable_str = hk
                self.cfg["hotkey_disable"] = hk
            else:
                self.hotkey_enable_registered = True
                self.hotkey_enable_str = hk
                self.cfg["hotkey_enable"] = hk
            save_config(self.cfg)

        return ok

    def hook_wm_hotkey(self):
        if getattr(self, "_wndproc_installed", False):
            return

        hwnd = self.root.winfo_id()
        GWL_WNDPROC = -4
        old_ptr = user32.GetWindowLongPtrW(hwnd, GWL_WNDPROC)
        self._old_wndproc = ctypes.cast(old_ptr, WNDPROC)

        @WNDPROC
        def new_wndproc(hWnd, msg, wParam, lParam):
            if msg == WM_HOTKEY:
                if int(wParam) == self.hotkey_disable_id:
                    self.ui_call(self.run_disable)
                    return 0
                if int(wParam) == self.hotkey_enable_id:
                    self.ui_call(self.run_enable)
                    return 0
            return user32.CallWindowProcW(self._old_wndproc, hWnd, msg, wParam, lParam)

        self._new_wndproc = new_wndproc
        user32.SetWindowLongPtrW(hwnd, GWL_WNDPROC, self._new_wndproc)
        self._wndproc_installed = True

    def unhook_wm_hotkey(self):
        if getattr(self, "_wndproc_installed", False):
            try:
                hwnd = self.root.winfo_id()
                GWL_WNDPROC = -4
                user32.SetWindowLongPtrW(hwnd, GWL_WNDPROC, self._old_wndproc)
            except Exception:
                pass
            self._wndproc_installed = False

    def on_set_hotkey(self, which: str):
        win = tk.Toplevel(self.root)
        title = "Set Disable Hotkey" if which == "disable" else "Set Enable Hotkey"
        win.title(title)
        win.resizable(False, False)
        win.grab_set()

        tk.Label(
            win,
            text="Press the desired hotkey combination\n(hold modifiers, press final key).",
            justify="left",
            padx=12,
            pady=10,
        ).pack(anchor="w")

        current = self.hotkey_disable_str if which == "disable" else self.hotkey_enable_str
        status = tk.StringVar(value=f"Current: {current}")
        status_lbl = tk.Label(win, textvariable=status, padx=12, pady=4, fg="blue")
        status_lbl.pack(anchor="w")

        tk.Label(win, text="ESC to cancel.", padx=12).pack(anchor="w")

        pressed_mods: set[str] = set()

        def add_mod(keysym: str):
            if keysym in ("Control_L", "Control_R"):
                pressed_mods.add("CTRL")
            elif keysym in ("Shift_L", "Shift_R"):
                pressed_mods.add("SHIFT")
            elif keysym in ("Alt_L", "Alt_R"):
                pressed_mods.add("ALT")
            elif keysym in ("Super_L", "Super_R", "Meta_L", "Meta_R"):
                pressed_mods.add("WIN")

        def remove_mod(keysym: str):
            if keysym in ("Control_L", "Control_R"):
                pressed_mods.discard("CTRL")
            elif keysym in ("Shift_L", "Shift_R"):
                pressed_mods.discard("SHIFT")
            elif keysym in ("Alt_L", "Alt_R"):
                pressed_mods.discard("ALT")
            elif keysym in ("Super_L", "Super_R", "Meta_L", "Meta_R"):
                pressed_mods.discard("WIN")

        def on_key_press(event):
            if event.keysym == "Escape":
                win.destroy()
                return

            if event.keysym in ("Control_L", "Control_R", "Shift_L", "Shift_R", "Alt_L", "Alt_R", "Super_L", "Super_R", "Meta_L", "Meta_R"):
                add_mod(event.keysym)
                return

            # Build from currently pressed modifiers
            mods = list(pressed_mods)
            key = event.keysym.upper()
            special_map = {
                "PAUSE": "PAUSE",
                "INSERT": "INSERT",
                "DELETE": "DELETE",
                "HOME": "HOME",
                "END": "END",
                "PRIOR": "PGUP",      # Page Up
                "NEXT": "PGDN",       # Page Down
            }
            if key.startswith("F") and key[1:].isdigit():
                norm_key = key
            elif key in special_map:
                norm_key = special_map[key]
            elif len(key) == 1 and "A" <= key <= "Z":
                norm_key = key
            else:
                return

            if not mods:
                return  # enforce at least one modifier

            hk = "+".join(mods + [norm_key])
            status.set(f"Selected: {hk}")
            hotkey_id = self.hotkey_disable_id if which == "disable" else self.hotkey_enable_id
            ok = self.register_hotkey_from_string(hk, hotkey_id, silent_fail=False)
            if ok:
                win.destroy()

        def on_key_release(event):
            remove_mod(event.keysym)

        # Focus and capture all key presses
        win.bind("<KeyPress>", on_key_press)
        win.bind("<KeyRelease>", on_key_release)
        win.focus_force()

    def on_set_disable_hotkey(self):
        self.on_set_hotkey("disable")

    def on_set_enable_hotkey(self):
        self.on_set_hotkey("enable")

    # ---- Tray icon ----
    def init_tray_icon(self):
        try:
            import pystray  # type: ignore
            from PIL import Image, ImageDraw  # type: ignore
        except Exception:
            self.tray_available = False
            return

        self.tray_available = True
        self._pystray = pystray
        self._PIL_Image = Image
        self._PIL_Draw = ImageDraw

        self.tray_icon = self._pystray.Icon(APP_NAME, self.make_tray_image(), APP_NAME, menu=self.make_tray_menu())

        def run_tray():
            try:
                self.tray_icon.run()
            except Exception:
                pass

        self._tray_thread = threading.Thread(target=run_tray, daemon=True)
        self._tray_thread.start()

    def make_tray_menu(self):
        pystray = self._pystray
        return pystray.Menu(
            pystray.MenuItem("Show Window", lambda: self.ui_call(self.show_window), default=True),
            pystray.MenuItem("Hide Window", lambda: self.ui_call(self.hide_to_tray)),
            pystray.MenuItem("Toggle Kill switch", lambda: self.ui_call(self.on_toggle_clicked)),
            pystray.MenuItem("Exit", lambda: self.ui_call(self.exit_app)),
        )

    def make_tray_image(self):
        Image = self._PIL_Image
        ImageDraw = self._PIL_Draw

        base = None
        if os.path.exists(APP_ICON_PATH):
            try:
                base = Image.open(APP_ICON_PATH).convert("RGBA")
            except Exception:
                base = None

        if base is None:
            base = Image.new("RGBA", (64, 64), (0, 0, 0, 0))
            d = ImageDraw.Draw(base)
            d.ellipse((10, 10, 54, 54), outline=(30, 30, 30, 255), width=4)
            d.ellipse((16, 16, 48, 48), fill=(80, 80, 80, 255))

        img = base.resize((64, 64))
        d = ImageDraw.Draw(img)
        dot = (200, 0, 0, 255) if self.killed else (30, 120, 220, 255)
        d.ellipse((46, 46, 62, 62), fill=dot, outline=(30, 30, 30, 255))
        return img

    def update_tray_icon_image(self):
        if self.tray_available and self.tray_icon is not None:
            try:
                self.tray_icon.icon = self.make_tray_image()
            except Exception:
                pass

    def hide_to_tray(self):
        if not self.tray_available:
            messagebox.showwarning(
                "Tray not available",
                "System tray support requires:\n  pip install pystray pillow\n\n"
                "Rebuild the EXE after installing them."
            )
            return
        self.root.withdraw()

    def show_window(self):
        self.root.deiconify()
        try:
            self.root.state("normal")
        except Exception:
            pass
        self.root.lift()
        try:
            self.root.focus_force()
        except Exception:
            pass

    def toggle_window_visibility(self):
        if self.root.state() == "withdrawn":
            self.show_window()
        else:
            self.root.withdraw()

    def exit_app(self):
        if self.killed:
            try:
                restore = messagebox.askyesno(
                    "Restore connections?",
                    "Kill switch is active. Restore previously disabled connections before exit?",
                )
            except Exception:
                restore = False
            if restore:
                try:
                    self.enable_network()
                except Exception:
                    pass
        self.unregister_hotkey(self.hotkey_disable_id)
        self.unregister_hotkey(self.hotkey_enable_id)
        self.unhook_wm_hotkey()
        if self.tray_available and self.tray_icon is not None:
            try:
                self.tray_icon.stop()
            except Exception:
                pass
        try:
            self.root.quit()
        except Exception:
            pass
        self.root.destroy()
        release_single_instance()
        try:
            sys.exit(0)
        except Exception:
            pass

    def on_close_clicked(self):
        if self.tray_available:
            self.hide_to_tray()
        else:
            self.exit_app()


def main():
    # Hide the console window immediately (when running as .py, not PyInstaller)
    if os.name == "nt" and not getattr(sys, "frozen", False):
        try:
            hwnd = ctypes.windll.kernel32.GetConsoleWindow()
            if hwnd:
                ctypes.windll.user32.ShowWindow(hwnd, 0)  # SW_HIDE
        except Exception:
            pass

    # UAC prompt on launch if not admin
    relaunch_as_admin_or_exit()

    if not acquire_single_instance():
        return

    root = tk.Tk()
    root.withdraw()
    splash, pb = show_splash(root)
    KillSwitchApp(root)
    try:
        pb.stop()
    except Exception:
        pass
    try:
        splash.destroy()
    except Exception:
        pass
    root.deiconify()
    root.mainloop()


if __name__ == "__main__":
    main()
