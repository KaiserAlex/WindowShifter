<#
.SYNOPSIS
    WindowShifter - Move windows between virtual desktops with a hotkey.
.DESCRIPTION
    System-tray utility for Windows 11. Press a configurable hotkey (default: Ctrl+Shift+W)
    to show a popup listing all virtual desktops. Press a number to move the active window.
#>

param([switch]$NoAutostart)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# -- Win32 Interop -----------------------------------------------------------------

Add-Type @'
using System;
using System.Runtime.InteropServices;
using System.Text;

public class Win32 {
    [DllImport("user32.dll")] public static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);
    [DllImport("user32.dll")] public static extern bool UnregisterHotKey(IntPtr hWnd, int id);
    [DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow();
    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    public static extern int GetWindowText(IntPtr hWnd, StringBuilder text, int count);
    [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);
}
'@

# -- Virtual Desktop (PSVirtualDesktop module) -------------------------------------

# Auto-install the VirtualDesktop module if not available
if (-not (Get-Module -ListAvailable -Name VirtualDesktop)) {
    try {
        [System.Windows.Forms.MessageBox]::Show(
            "The VirtualDesktop PowerShell module is required but not installed.`n`nIt will be installed now (one-time setup).",
            "WindowShifter - First Time Setup",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        ) | Out-Null
        Install-Module -Name VirtualDesktop -Scope CurrentUser -Force -AllowClobber
    } catch {
        [System.Windows.Forms.MessageBox]::Show(
            "Failed to install the VirtualDesktop module:`n" + $_.Exception.Message + "`n`nPlease run manually:`nInstall-Module -Name VirtualDesktop -Scope CurrentUser",
            "WindowShifter - Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
        exit 1
    }
}

Import-Module VirtualDesktop -Force -ErrorAction Stop -WarningAction SilentlyContinue

# -- Settings ----------------------------------------------------------------------

$scriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent (Convert-Path .) }
$settingsPath = Join-Path $scriptDir "windowshifter-settings.json"

$defaultSettings = @{
    ModifierKeys  = 6
    VirtualKey    = 0x57
    HotkeyDisplay = "Ctrl + Shift + W"
    FirstRunComplete = $false
}

function Load-Settings {
    if (Test-Path $settingsPath) {
        try {
            $json = Get-Content $settingsPath -Raw | ConvertFrom-Json
            return @{
                ModifierKeys     = [uint32]$json.ModifierKeys
                VirtualKey       = [uint32]$json.VirtualKey
                HotkeyDisplay    = [string]$json.HotkeyDisplay
                FirstRunComplete = [bool]$json.FirstRunComplete
            }
        } catch { }
    }
    return $defaultSettings.Clone()
}

function Save-Settings($s) {
    $s | ConvertTo-Json | Set-Content $settingsPath -Force
}

# -- First-Run Hotkey Setup --------------------------------------------------------

function Show-HotkeySetup($settings) {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "WindowShifter - Hotkey Setup"
    $form.Size = New-Object System.Drawing.Size(420, 280)
    $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = "FixedDialog"
    $form.MaximizeBox = $false
    $form.MinimizeBox = $false
    $form.TopMost = $true

    $lblTitle = New-Object System.Windows.Forms.Label
    $lblTitle.Text = "Welcome to WindowShifter!"
    $lblTitle.Font = New-Object System.Drawing.Font("Segoe UI", 14, [System.Drawing.FontStyle]::Bold)
    $lblTitle.Location = New-Object System.Drawing.Point(20, 15)
    $lblTitle.AutoSize = $true
    $form.Controls.Add($lblTitle)

    $lblDesc = New-Object System.Windows.Forms.Label
    $descText = "Press the hotkey combination you want to use."
    $descText += "`nClick the box below then press your desired keys."
    $descText += "`nDefault: Ctrl + Shift + W"
    $lblDesc.Text = $descText
    $lblDesc.Location = New-Object System.Drawing.Point(20, 50)
    $lblDesc.Size = New-Object System.Drawing.Size(370, 50)
    $form.Controls.Add($lblDesc)

    $txtHotkey = New-Object System.Windows.Forms.TextBox
    $txtHotkey.Text = $settings.HotkeyDisplay
    $txtHotkey.Font = New-Object System.Drawing.Font("Segoe UI", 14, [System.Drawing.FontStyle]::Bold)
    $txtHotkey.TextAlign = "Center"
    $txtHotkey.ReadOnly = $true
    $txtHotkey.Location = New-Object System.Drawing.Point(20, 110)
    $txtHotkey.Size = New-Object System.Drawing.Size(360, 36)
    $txtHotkey.BackColor = [System.Drawing.Color]::FromArgb(240, 240, 240)

    $txtHotkey.Tag = @{
        Modifiers = $settings.ModifierKeys
        VKey      = $settings.VirtualKey
        Display   = $settings.HotkeyDisplay
    }

    $script:setupHintLabel = $null

    $txtHotkey.Add_KeyDown({
        param($sender, $e)
        $e.SuppressKeyPress = $true
        $e.Handled = $true

        $key = $e.KeyCode
        $modifierKeys = @(
            [System.Windows.Forms.Keys]::ControlKey,
            [System.Windows.Forms.Keys]::ShiftKey,
            [System.Windows.Forms.Keys]::Menu,
            [System.Windows.Forms.Keys]::LWin,
            [System.Windows.Forms.Keys]::RWin
        )
        if ($modifierKeys -contains $key) { return }

        $mod = [uint32]0
        $parts = @()
        if ($e.Control) { $mod = $mod -bor 0x0002; $parts += "Ctrl" }
        if ($e.Alt)     { $mod = $mod -bor 0x0001; $parts += "Alt"  }
        if ($e.Shift)   { $mod = $mod -bor 0x0004; $parts += "Shift" }

        if ($mod -eq 0) {
            if ($script:setupHintLabel) {
                $script:setupHintLabel.Text = "Please include at least one modifier key."
            }
            return
        }

        $parts += $key.ToString()
        $display = $parts -join " + "

        $sender.Text = $display
        $sender.Tag = @{ Modifiers = $mod; VKey = [uint32]$key; Display = $display }
        if ($script:setupHintLabel) {
            $script:setupHintLabel.Text = "Hotkey set to: $display"
        }
    })
    $form.Controls.Add($txtHotkey)

    $lblHint = New-Object System.Windows.Forms.Label
    $lblHint.Text = "Click the box above then press your desired hotkey..."
    $lblHint.ForeColor = [System.Drawing.Color]::Gray
    $lblHint.Location = New-Object System.Drawing.Point(20, 155)
    $lblHint.Size = New-Object System.Drawing.Size(360, 20)
    $form.Controls.Add($lblHint)
    $script:setupHintLabel = $lblHint

    $btnReset = New-Object System.Windows.Forms.Button
    $btnReset.Text = "Reset to Default"
    $btnReset.Location = New-Object System.Drawing.Point(140, 195)
    $btnReset.Size = New-Object System.Drawing.Size(120, 32)
    $btnReset.Add_Click({
        $txtHotkey.Text = "Ctrl + Shift + W"
        $txtHotkey.Tag = @{ Modifiers = [uint32]6; VKey = [uint32]0x57; Display = "Ctrl + Shift + W" }
        if ($script:setupHintLabel) { $script:setupHintLabel.Text = "Reset to default." }
    })
    $form.Controls.Add($btnReset)

    $btnSave = New-Object System.Windows.Forms.Button
    $btnSave.Text = "Save and Start"
    $btnSave.Location = New-Object System.Drawing.Point(270, 195)
    $btnSave.Size = New-Object System.Drawing.Size(110, 32)
    $btnSave.Add_Click({
        $form.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $form.Close()
    })
    $form.Controls.Add($btnSave)
    $form.AcceptButton = $btnSave

    $result = $form.ShowDialog()
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        $tag = $txtHotkey.Tag
        $settings.ModifierKeys  = $tag.Modifiers
        $settings.VirtualKey    = $tag.VKey
        $settings.HotkeyDisplay = $tag.Display
    }
    $form.Dispose()
    return $settings
}

# -- Autostart ---------------------------------------------------------------------

function Ensure-Autostart {
    if ($NoAutostart) { return }
    try {
        $regPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Run"
        $current = Get-ItemProperty -Path $regPath -Name "WindowShifter" -ErrorAction SilentlyContinue

        $exePath = $null
        if ($PSScriptRoot) {
            $scriptFile = Join-Path $PSScriptRoot "WindowShifter.ps1"
            # Use pwsh if available, fallback to powershell.exe
            $psExe = if (Get-Command pwsh -ErrorAction SilentlyContinue) { "pwsh.exe" } else { "powershell.exe" }
            $exePath = "$psExe -WindowStyle Hidden -ExecutionPolicy Bypass -File `"$scriptFile`""
        }

        if (-not $exePath) { return }

        if ($current -and ($current.WindowShifter -match [regex]::Escape($scriptFile))) { return }

        $answer = [System.Windows.Forms.MessageBox]::Show(
            "WindowShifter is not set to start automatically with Windows.`n`nWould you like to add it to autostart?",
            "WindowShifter - Autostart",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question
        )

        if ($answer -eq [System.Windows.Forms.DialogResult]::Yes) {
            New-ItemProperty -Path $regPath -Name "WindowShifter" -Value $exePath -PropertyType String -Force | Out-Null
        }
    } catch {
        $errMsg = "Could not configure autostart:`n" + $_.Exception.Message + "`n`nYou can add it manually via the Startup folder."
        [System.Windows.Forms.MessageBox]::Show(
            $errMsg,
            "WindowShifter",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        ) | Out-Null
    }
}

# -- Desktop Picker Popup ---------------------------------------------------------

function Show-DesktopPicker {
    param([IntPtr]$targetHwnd, [string]$windowTitle)

    $desktopList = @(Get-DesktopList)
    $currentDesktop = Get-CurrentDesktop
    $currentIndex = Get-DesktopIndex $currentDesktop

    if ($desktopList.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("No virtual desktops found.", "WindowShifter") | Out-Null
        return
    }

    $form = New-Object System.Windows.Forms.Form
    $form.Text = "WindowShifter"
    $form.FormBorderStyle = "None"
    $form.StartPosition = "CenterScreen"
    $form.TopMost = $true
    $form.ShowInTaskbar = $false
    $form.BackColor = [System.Drawing.Color]::FromArgb(32, 32, 32)
    $form.ForeColor = [System.Drawing.Color]::White
    $form.Padding = New-Object System.Windows.Forms.Padding(20)

    $y = 20

    $lblHeader = New-Object System.Windows.Forms.Label
    $lblHeader.Text = "Move Window To Desktop"
    $lblHeader.ForeColor = [System.Drawing.Color]::FromArgb(170, 170, 170)
    $lblHeader.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $lblHeader.Location = New-Object System.Drawing.Point(20, $y)
    $lblHeader.AutoSize = $true
    $form.Controls.Add($lblHeader)
    $y += 28

    $lblTitle = New-Object System.Windows.Forms.Label
    $truncatedTitle = $windowTitle
    if ($truncatedTitle.Length -gt 100) {
        $truncatedTitle = $truncatedTitle.Substring(0, 97) + "..."
    }
    $lblTitle.Text = $truncatedTitle
    $lblTitle.Font = New-Object System.Drawing.Font("Segoe UI", 13, [System.Drawing.FontStyle]::Bold)
    $lblTitle.ForeColor = [System.Drawing.Color]::White
    $lblTitle.Location = New-Object System.Drawing.Point(20, $y)
    $lblTitle.Size = New-Object System.Drawing.Size(360, 25)
    $lblTitle.AutoSize = $false
    $form.Controls.Add($lblTitle)
    $y += 40

    $sep = New-Object System.Windows.Forms.Label
    $sep.BorderStyle = "Fixed3D"
    $sep.Location = New-Object System.Drawing.Point(20, $y)
    $sep.Size = New-Object System.Drawing.Size(340, 2)
    $form.Controls.Add($sep)
    $y += 12

    $maxDesktops = [Math]::Min($desktopList.Count, 9)
    for ($i = 0; $i -lt $maxDesktops; $i++) {
        $d = $desktopList[$i]
        $isCurrent = ($i -eq $currentIndex)
        $desktopName = $d.Name
        if ([string]::IsNullOrWhiteSpace($desktopName)) {
            $desktopName = "Desktop " + ($i + 1)
        }
        $marker = ""
        if ($isCurrent) { $marker = "  * current" }

        $lbl = New-Object System.Windows.Forms.Label
        $num = $i + 1
        $lbl.Text = "  $num   $desktopName$marker"
        $lbl.Font = New-Object System.Drawing.Font("Segoe UI", 12)
        if ($isCurrent) {
            $lbl.ForeColor = [System.Drawing.Color]::FromArgb(102, 187, 106)
        } else {
            $lbl.ForeColor = [System.Drawing.Color]::FromArgb(220, 220, 220)
        }
        $lbl.Location = New-Object System.Drawing.Point(20, $y)
        $lbl.AutoSize = $true
        $form.Controls.Add($lbl)
        $y += 30
    }

    $y += 8

    $lblHint = New-Object System.Windows.Forms.Label
    $lblHint.Text = "Press number to move  |  Esc to cancel"
    $lblHint.ForeColor = [System.Drawing.Color]::FromArgb(119, 119, 119)
    $lblHint.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $lblHint.Location = New-Object System.Drawing.Point(20, $y)
    $lblHint.AutoSize = $true
    $form.Controls.Add($lblHint)
    $y += 30

    $form.Size = New-Object System.Drawing.Size(400, $y)

    $form.Add_Paint({
        param($s, $e)
        $pen = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(85, 85, 85), 1)
        $rect = New-Object System.Drawing.Rectangle(0, 0, ($s.Width - 1), ($s.Height - 1))
        $e.Graphics.DrawRectangle($pen, $rect)
        $pen.Dispose()
    })

    # Ensure keyboard focus on show
    $form.Add_Shown({
        $form.Activate()
        $form.Focus()
    })

    $form.Add_KeyDown({
        param($sender, $e)
        if ($e.KeyCode -eq [System.Windows.Forms.Keys]::Escape) {
            $form.Close()
            return
        }

        $num = -1
        switch ($e.KeyCode) {
            ([System.Windows.Forms.Keys]::D1)      { $num = 1 }
            ([System.Windows.Forms.Keys]::NumPad1)  { $num = 1 }
            ([System.Windows.Forms.Keys]::D2)      { $num = 2 }
            ([System.Windows.Forms.Keys]::NumPad2)  { $num = 2 }
            ([System.Windows.Forms.Keys]::D3)      { $num = 3 }
            ([System.Windows.Forms.Keys]::NumPad3)  { $num = 3 }
            ([System.Windows.Forms.Keys]::D4)      { $num = 4 }
            ([System.Windows.Forms.Keys]::NumPad4)  { $num = 4 }
            ([System.Windows.Forms.Keys]::D5)      { $num = 5 }
            ([System.Windows.Forms.Keys]::NumPad5)  { $num = 5 }
            ([System.Windows.Forms.Keys]::D6)      { $num = 6 }
            ([System.Windows.Forms.Keys]::NumPad6)  { $num = 6 }
            ([System.Windows.Forms.Keys]::D7)      { $num = 7 }
            ([System.Windows.Forms.Keys]::NumPad7)  { $num = 7 }
            ([System.Windows.Forms.Keys]::D8)      { $num = 8 }
            ([System.Windows.Forms.Keys]::NumPad8)  { $num = 8 }
            ([System.Windows.Forms.Keys]::D9)      { $num = 9 }
            ([System.Windows.Forms.Keys]::NumPad9)  { $num = 9 }
        }

        if ($num -ge 1 -and $num -le $maxDesktops) {
            try {
                $targetDesktop = Get-Desktop $($num - 1)
                Move-Window -Desktop $targetDesktop -Hwnd $targetHwnd
            } catch {
                $errMsg = "Failed to move window:`n" + $_.Exception.Message
                [System.Windows.Forms.MessageBox]::Show(
                    $errMsg,
                    "WindowShifter",
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Error
                ) | Out-Null
            }
            $form.Close()
        }
    })

    $form.Add_Deactivate({ $form.Close() })
    $form.KeyPreview = $true
    $form.ShowDialog() | Out-Null
    $form.Dispose()
}

# -- Main --------------------------------------------------------------------------

# Hide the console window
Add-Type @'
using System;
using System.Runtime.InteropServices;
public class ConsoleHelper {
    [DllImport("kernel32.dll")] public static extern IntPtr GetConsoleWindow();
    [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
}
'@
$consoleHwnd = [ConsoleHelper]::GetConsoleWindow()
if ($consoleHwnd -ne [IntPtr]::Zero) {
    [ConsoleHelper]::ShowWindow($consoleHwnd, 0) | Out-Null
}

# Load settings
$settings = Load-Settings

# First-run setup
if (-not $settings.FirstRunComplete) {
    $settings = Show-HotkeySetup $settings
    $settings.FirstRunComplete = $true
    Save-Settings $settings
}

# Autostart check
Ensure-Autostart

# -- Hotkey listener ---------------------------------------------------------------

# Build referenced assemblies list dynamically (handles .NET 8/9 where Message moved)
$referencedAssemblies = @('System.Windows.Forms')
$primitivesPath = Join-Path ([System.Runtime.InteropServices.RuntimeEnvironment]::GetRuntimeDirectory()) "System.Windows.Forms.Primitives.dll"
if (-not (Test-Path $primitivesPath)) {
    # Try next to the PS executable
    $primitivesPath = Join-Path (Split-Path ([System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName)) "System.Windows.Forms.Primitives.dll"
}
if (Test-Path $primitivesPath) {
    $referencedAssemblies += $primitivesPath
}

Add-Type -ReferencedAssemblies $referencedAssemblies -TypeDefinition @'
using System;
using System.Windows.Forms;

public class HotkeyWindow : NativeWindow {
    private const int WM_HOTKEY = 0x0312;
    public event EventHandler HotkeyPressed;

    public HotkeyWindow() {
        CreateHandle(new CreateParams() {
            Caption = "WindowShifterHotkey",
            Style = 0
        });
    }

    protected override void WndProc(ref Message m) {
        if (m.Msg == WM_HOTKEY) {
            if (HotkeyPressed != null) HotkeyPressed(this, EventArgs.Empty);
        }
        base.WndProc(ref m);
    }
}
'@

$hotkeyWindow = New-Object HotkeyWindow
$HOTKEY_ID = 9000
$registered = [Win32]::RegisterHotKey($hotkeyWindow.Handle, $HOTKEY_ID, $settings.ModifierKeys, $settings.VirtualKey)

if (-not $registered) {
    $failMsg = "Failed to register hotkey " + $settings.HotkeyDisplay + ".`nAnother application may be using it."
    [System.Windows.Forms.MessageBox]::Show(
        $failMsg,
        "WindowShifter",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Warning
    ) | Out-Null
}

$hotkeyWindow.Add_HotkeyPressed({
    $hwnd = [Win32]::GetForegroundWindow()
    $sb = New-Object System.Text.StringBuilder(256)
    [Win32]::GetWindowText($hwnd, $sb, 256) | Out-Null
    $title = $sb.ToString()
    if ([string]::IsNullOrWhiteSpace($title)) { $title = "(Untitled Window)" }
    Show-DesktopPicker -targetHwnd $hwnd -windowTitle $title
})

# -- Custom tray icon --------------------------------------------------------------

$trayIcon = New-Object System.Windows.Forms.NotifyIcon

# Load custom icon if available, otherwise use default
$iconPath = Join-Path $scriptDir "windowshifter.ico"
if (Test-Path $iconPath) {
    try {
        $trayIcon.Icon = New-Object System.Drawing.Icon($iconPath)
    } catch {
        $trayIcon.Icon = [System.Drawing.SystemIcons]::Application
    }
} else {
    $trayIcon.Icon = [System.Drawing.SystemIcons]::Application
}

$trayIcon.Text = "WindowShifter - " + $settings.HotkeyDisplay
$trayIcon.Visible = $true

$contextMenu = New-Object System.Windows.Forms.ContextMenuStrip
$hotkeyLabel = "Hotkey: " + $settings.HotkeyDisplay
$hotkeyItem = $contextMenu.Items.Add($hotkeyLabel)
$hotkeyItem.Enabled = $false
$contextMenu.Items.Add("-") | Out-Null
$contextMenu.Items.Add("Exit", $null, {
    [Win32]::UnregisterHotKey($hotkeyWindow.Handle, $HOTKEY_ID) | Out-Null
    $trayIcon.Visible = $false
    $trayIcon.Dispose()
    $hotkeyWindow.DestroyHandle()
    [System.Windows.Forms.Application]::Exit()
}) | Out-Null
$trayIcon.ContextMenuStrip = $contextMenu

$trayIcon.Add_DoubleClick({
    $hwnd = [Win32]::GetForegroundWindow()
    $sb = New-Object System.Text.StringBuilder(256)
    [Win32]::GetWindowText($hwnd, $sb, 256) | Out-Null
    $title = $sb.ToString()
    if ([string]::IsNullOrWhiteSpace($title)) { $title = "(Untitled Window)" }
    Show-DesktopPicker -targetHwnd $hwnd -windowTitle $title
})

# Run the message loop (ApplicationContext keeps it alive without a visible form)
$appContext = New-Object System.Windows.Forms.ApplicationContext
[System.Windows.Forms.Application]::Run($appContext)

# Cleanup
[Win32]::UnregisterHotKey($hotkeyWindow.Handle, $HOTKEY_ID) | Out-Null
$hotkeyWindow.DestroyHandle()
$trayIcon.Dispose()
