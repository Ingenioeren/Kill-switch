# Network Kill Switch

Windows desktop app for instantly disabling and restoring selected network adapters with global hotkeys.

It is built for situations where you want a fast offline toggle while gaming, testing, or isolating a machine without opening Windows network settings every time.

## Features

- Disable selected network adapters with one click or a global hotkey
- Restore previously disabled adapters with a separate global hotkey
- Works while other apps or games are focused
- Auto-prompts for administrator rights on launch
- Optional system tray support
- Optional "kill process on disable" behavior
- Remembers selected adapters and hotkey settings
- Writes a simple local log file for troubleshooting

## Requirements

- Windows
- Python 3.10 or newer
- Administrator rights to enable or disable adapters

Optional dependencies for tray support:

- `pystray`
- `pillow`

Optional dependency for packaging:

- `pyinstaller`

## Run From Source

Install optional tray dependencies if you want full functionality:

```powershell
pip install pystray pillow
```

Start the app:

```powershell
python .\killswitch_win.py
```

The app will request administrator privileges if needed.

## Build Executable

Install build dependencies:

```powershell
pip install pyinstaller pystray pillow
```

Build with:

```powershell
pyinstaller --noconsole --onedir --name "Kill switch" .\killswitch_win.py
```

This project also includes a PyInstaller spec file at `Kill switch.spec`.

## How It Works

The app uses Windows network adapter controls through PowerShell and standard Windows APIs:

- `Disable-NetAdapter` / `Enable-NetAdapter` for adapter control
- `RegisterHotKey` for true system-wide hotkeys
- Tkinter for the desktop UI
- `pystray` and `pillow` for tray support

When you disable the connection, the app stores which adapters it turned off so it can restore the same adapters later.

## Usage

1. Launch the app as administrator.
2. Select the network adapters you want the kill switch to control.
3. Set or keep the default hotkeys.
4. Press `Disable Network Adapters` or use the disable hotkey.
5. Press `Enable Network Adapters` or use the enable hotkey to restore connectivity.

Optional:

- Choose a running program and enable `Kill program on network adapter disable`
- Hide the app to the system tray

## Default Hotkeys

- Disable: `CTRL+ALT+K`
- Enable: `CTRL+ALT+E`

Some common gaming keys are blocked by default to reduce conflicts.

## Config And Logs

The app stores its config and log in your Windows user profile:

- Config: `%USERPROFILE%\\.kill switch.json`
- Log: `%USERPROFILE%\\.kill switch.log`

Saved settings include:

- selected adapters
- disable hotkey
- enable hotkey
- kill-on-disable setting
- selected process name

## Project Files

- `killswitch_win.py` - main application
- `Kill switch.spec` - PyInstaller build spec
- `app.ico` - application icon

## Notes

- This is a Windows-only project.
- Administrator access is required for adapter changes.
- If a hotkey fails to register, another application is usually already using it.
- Tray support is unavailable unless `pystray` and `pillow` are installed.

## License

This project is licensed under the MIT License.
