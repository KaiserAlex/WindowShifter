# WindowShifter

A lightweight Windows 11 system-tray utility that lets you quickly manage virtual desktops and move windows between them using a keyboard shortcut.

**No .NET SDK or runtime install needed** -- runs on any Windows 11 machine using PowerShell. The only dependency ([VirtualDesktop module](https://github.com/MScholtes/PSVirtualDesktop)) is auto-installed on first launch.

## How it works

1. **First launch:** A setup dialog lets you choose your preferred hotkey (default: `Ctrl + Shift + W`). WindowShifter also offers to register itself for Windows autostart.
2. **After setup:** WindowShifter runs silently in the system tray.
3. The hotkey supports three modes depending on how many times you press it:

| Tap | Mode | What it does |
|-----|------|-------------|
| **Single** | **Switch Desktop** | Opens a popup to switch to another desktop (no window moved) |
| **Double** | **Move Window** | Opens a popup to move the active window to another desktop and switch to it |
| **Triple** | **Close Desktop** | Opens a popup to close/remove a desktop |

### Switch Desktop (single tap)
- Press your hotkey once. A popup with a blue **"Switch Desktop"** header appears.
- Press a **number key (1–9)** to switch to that desktop.
- Press **`Esc`** or click away to cancel.

### Move Window (double tap)
- Press your hotkey twice quickly. A popup with **"Move Window To Desktop"** header appears showing the active window title.
- Press a **number key (1–9)** to move the window there and switch to it.
- Press **`N`** to **create a new desktop** -- you'll be prompted for a name, and the window is automatically moved there.
- Press **`Esc`** or click away to cancel.

### Close Desktop (triple tap)
- Press your hotkey three times quickly. A popup with a red **"Close Desktop"** header appears.
- Press a **number key (1–9)** to close that desktop. Windows on it are moved to an adjacent desktop.
- Press **`Esc`** or click away to cancel.

## Quick Start

### Option A: Double-click the launcher

Just double-click **`WindowShifter.bat`** -- it automatically uses PowerShell 7 (`pwsh`) if available, otherwise falls back to Windows PowerShell 5.1.

### Option B: Run directly

```powershell
pwsh -WindowStyle Hidden -ExecutionPolicy Bypass -File WindowShifter.ps1
```

### Clone from source

```powershell
git clone https://github.com/KaiserAlex/WindowShifter.git
cd WindowShifter
.\WindowShifter.bat
```

## First Run

On the first launch, WindowShifter will:

1. **Auto-install the VirtualDesktop module** if not already present (one-time, with notification).
2. **Ask you to choose a hotkey** -- press any key combination with at least one modifier (Ctrl, Shift, Alt). Default is `Ctrl + Shift + W`.
3. **Ask to register for autostart** -- adds itself to `HKCU\Software\Microsoft\Windows\CurrentVersion\Run` so it launches on every Windows login. You can decline.

Settings are saved to `windowshifter-settings.json` next to the script.

## Autostart

WindowShifter checks on every launch whether it's registered in Windows autostart. If not, it asks for your permission. No admin rights required.

To **skip the autostart prompt**, run with:

```powershell
.\WindowShifter.ps1 -NoAutostart
```

To **manually remove** autostart:

```powershell
Remove-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Run" -Name "WindowShifter"
```

## System Tray

WindowShifter uses a custom icon (window frame with green arrow) in the system tray.

- **Double-click** the tray icon to manually open the picker popup.
- **Right-click** for context menu:
  - **Hotkey** -- shows the current hotkey; click to **change** it (opens setup dialog)
  - **Remove from Autostart** -- removes WindowShifter from Windows autostart
  - **GitHub Repository** -- opens the project page in your browser
  - **Exit** -- quit the application

## Configuration

All settings are stored in `windowshifter-settings.json` next to the script:

```json
{
  "ModifierKeys": 6,
  "VirtualKey": 87,
  "HotkeyDisplay": "Ctrl + Shift + W",
  "FirstRunComplete": true
}
```

Delete this file to trigger the first-run setup again.

## Technical Details

- **Pure PowerShell** -- no compiled code, minimal dependencies
- **Win32 Interop** -- `RegisterHotKey`, `GetForegroundWindow`, `GetWindowText` via P/Invoke
- **[PSVirtualDesktop](https://github.com/MScholtes/PSVirtualDesktop)** -- PowerShell module using internal COM APIs (can move any window, not just own-process)
- **WinForms** -- used for system tray icon, popup UI, and hotkey setup dialog (built into Windows)

## Requirements

- **Windows 11** (virtual desktop names require Windows 11)
- **PowerShell 7 recommended** ([install](https://learn.microsoft.com/en-us/powershell/scripting/install/installing-powershell-on-windows)) -- Windows PowerShell 5.1 also works
- **VirtualDesktop module** -- auto-installed on first launch, or install manually:
  ```powershell
  Install-Module -Name VirtualDesktop -Scope CurrentUser
  ```

## Limitations

- Supports up to **9 virtual desktops** (keys 1–9)
- The configured hotkey must not conflict with another application
- UI is WinForms-based (functional but not as polished as WPF)

## License

MIT
