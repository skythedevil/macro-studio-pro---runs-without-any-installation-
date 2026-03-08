try {
    # Relative log location
    $Global:LogPath = Join-Path $PSScriptRoot "MacroStudio_Error.log"

    # Robust Assembly Loading
    $assemblies = @("System.Windows.Forms", "System.Drawing", "Microsoft.VisualBasic")
    foreach ($asm in $assemblies) {
        if (-not ([System.AppDomain]::CurrentDomain.GetAssemblies() | Where-Object { $_.FullName -like "$asm*" })) {
            try { Add-Type -AssemblyName $asm } catch { }
        }
    }
<#
.SYNOPSIS
    Testing Macro Studio Pro - Advanced Productivity & Automation Suite
.DESCRIPTION
    A professional-grade automation tool for simulating user input and visual matching.
    Designed for high-performance productivity without requiring administrative privileges.
    
    CORE CAPABILITIES:
    - Professional 3-Column Studio Layout
    - Native Performance with Multi-Theme Support (Dark/Light)
    - Smart Window Context Matching & Verification
    - Visual AI Anchor Alignment (60x60 Pixel Matching)
    - Seamless Non-Admin Automation (Mouse/Keyboard/Scroll)
    - In-GUI Macro Management (Add/Remove/Edit Steps)
    - Zero-Dependency Portable PowerShell Architecture
    - Global Hotkey Integration (F8 Start/Stop, F9 Pause)
.NOTES
    Author: Saurabh Yadav
    Version: 1.5.0
    Release Date: 2026-03-08
    Copyright: (c) 2026 Macro Studio Labs. All rights reserved.
    Security: This script contains standard UI automation methods. 
#>


# --- Native Methods (Robust Definition) ---
if (-not ("TestingMacroStudioProWin32" -as [type])) {
    $NativeMethodsCode = @'
using System;
using System.Runtime.InteropServices;
using System.Text;

public class TestingMacroStudioProWin32 {
    [DllImport("user32.dll")]
    public static extern bool SetCursorPos(int X, int Y);

    [DllImport("user32.dll")]
    public static extern void mouse_event(uint dwFlags, uint dx, uint dy, uint dwData, uint dwExtraInfo);

    [DllImport("user32.dll")]
    public static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, uint dwExtraInfo);
    
    [DllImport("user32.dll")]
    public static extern short GetAsyncKeyState(int vKey);

    [DllImport("user32.dll")]
    public static extern IntPtr GetForegroundWindow();

    [DllImport("user32.dll")]
    public static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

    [DllImport("user32.dll")]
    public static extern bool SetForegroundWindow(IntPtr hWnd);

    [DllImport("user32.dll")]
    public static extern int SetWindowLong(IntPtr hWnd, int nIndex, int dwNewLong);
    
    [DllImport("user32.dll")]
    public static extern int GetWindowLong(IntPtr hWnd, int nIndex);

    public const int MOUSEEVENTF_LEFTDOWN = 0x02;
    public const int MOUSEEVENTF_LEFTUP = 0x04;
    public const int MOUSEEVENTF_RIGHTDOWN = 0x08;
    public const int MOUSEEVENTF_RIGHTUP = 0x10;
    public const int KEYEVENTF_KEYUP = 0x0002;
    public const int MOUSEEVENTF_WHEEL = 0x0800;
    public const int GWL_EXSTYLE = -20;
    public const int WS_EX_LAYERED = 0x80000;
    public const int WS_EX_TRANSPARENT = 0x20;
}
'@
    Add-Type -TypeDefinition $NativeMethodsCode
}

# --- GUI Construction ---
$Form = New-Object Windows.Forms.Form
$Form.Text = "Testing Macro Studio"
$Form.Size = New-Object Drawing.Size(750, 680)
$Form.StartPosition = "CenterScreen"
$Form.FormBorderStyle = "FixedDialog"
$Form.MaximizeBox = $false

$Font = New-Object Drawing.Font("Segoe UI", 10)
$BoldFont = New-Object Drawing.Font("Segoe UI", 10, [Drawing.FontStyle]::Bold)
$Form.Font = $Font

# --- Global Application State ---
$Global:IsDarkMode = $true
$Global:MacroSteps = New-Object System.Collections.Generic.List[Object]
$Global:TypeBuffer = ""
$Global:BufferTime = 0
$Global:BufferWin = ""

# --- Floating HUD for Playback ---
$FloatingHUD = New-Object Windows.Forms.Form
$FloatingHUD.FormBorderStyle = "None"
$FloatingHUD.TopMost = $true
$FloatingHUD.ShowInTaskbar = $false
$FloatingHUD.BackColor = [Drawing.Color]::FromArgb(40, 40, 40)
$FloatingHUD.Opacity = 0.8
$FloatingHUD.Size = New-Object Drawing.Size(450, 45)
$FloatingHUD.StartPosition = "Manual"
$HUDLabel = New-Object Windows.Forms.Label
$HUDLabel.ForeColor = [Drawing.Color]::White
$HUDLabel.Dock = "Fill"
$HUDLabel.TextAlign = "MiddleCenter"
$HUDLabel.Font = New-Object Drawing.Font("Segoe UI", 11, [Drawing.FontStyle]::Bold)
$FloatingHUD.Controls.Add($HUDLabel)

# Position HUD at bottom (Robust detection)
$primaryScreen = [System.Windows.Forms.Screen]::PrimaryScreen
if ($null -eq $primaryScreen) { $primaryScreen = [System.Windows.Forms.Screen]::AllScreens[0] }
$screenW = [int]$primaryScreen.Bounds.Width
$screenH = [int]$primaryScreen.Bounds.Height
$HUDPosX = [int]($screenW / 2 - 225)
$HUDPosY = [int]($screenH - 120)
$FloatingHUD.Location = New-Object Drawing.Point($HUDPosX, $HUDPosY)

# Make HUD Click-Through
$FloatingHUD.Add_Load({
    $hWnd = $FloatingHUD.Handle
    $exStyle = [TestingMacroStudioProWin32]::GetWindowLong($hWnd, -20)
    [TestingMacroStudioProWin32]::SetWindowLong($hWnd, -20, $exStyle -bor 0x80000 -bor 0x20) | Out-Null
})

# Pre-initialize colors to avoid null during control creation
$Global:colorBG = [Drawing.Color]::FromArgb(30,30,30)
$Global:colorPanel = [Drawing.Color]::FromArgb(45,45,45)
$Global:colorFG = [Drawing.Color]::White
$Global:colorAccent = [Drawing.Color]::FromArgb(81, 121, 255)
$Global:colorDanger = [Drawing.Color]::FromArgb(231, 76, 60)
$Global:colorSuccess = [Drawing.Color]::FromArgb(46, 204, 113)
$Global:colorBtnGray = [Drawing.Color]::FromArgb(70, 70, 70)

function Apply-Theme {
    param($isDark)
    if ($isDark) {
        $Global:colorBG = [Drawing.Color]::FromArgb(30,30,30)
        $Global:colorPanel = [Drawing.Color]::FromArgb(45,45,45)
        $Global:colorFG = [Drawing.Color]::White
        $Global:colorAccent = [Drawing.Color]::FromArgb(81, 121, 255)
        $Global:colorDanger = [Drawing.Color]::FromArgb(231, 76, 60)
        $Global:colorSuccess = [Drawing.Color]::FromArgb(46, 204, 113)
        $Global:colorBtnGray = [Drawing.Color]::FromArgb(70, 70, 70)
    } else {
        $Global:colorBG = [Drawing.Color]::FromArgb(240, 242, 245) # Facebook Gray
        $Global:colorPanel = [Drawing.Color]::White
        $Global:colorFG = [Drawing.Color]::Black
        $Global:colorAccent = [Drawing.Color]::FromArgb(24, 119, 242) # Facebook Blue
        $Global:colorDanger = [Drawing.Color]::FromArgb(210, 45, 45)
        $Global:colorSuccess = [Drawing.Color]::FromArgb(35, 160, 85)
        $Global:colorBtnGray = [Drawing.Color]::FromArgb(220, 220, 220)
    }

    $Form.BackColor = $Global:colorBG
    $Form.ForeColor = $Global:colorFG
    
    # Update All Controls
    $LabelSteps.ForeColor = $Global:colorAccent
    $ListSteps.BackColor = $Global:colorPanel; $ListSteps.ForeColor = $Global:colorFG
    
    $BtnAddType.BackColor = $Global:colorAccent; $BtnAddType.ForeColor = [Drawing.Color]::White
    $BtnAddScroll.BackColor = $Global:colorAccent; $BtnAddScroll.ForeColor = [Drawing.Color]::White
    $BtnClear.BackColor = $Global:colorDanger; $BtnClear.ForeColor = [Drawing.Color]::White
    
    $LabelDelay.ForeColor = $Global:colorAccent
    $TxtExtraDelay.BackColor = $Global:colorPanel; $TxtExtraDelay.ForeColor = $Global:colorFG
    $TxtRepeat.BackColor = $Global:colorPanel; $TxtRepeat.ForeColor = $Global:colorFG
    $SliderSpeed.BackColor = $Global:colorBG
    
    $LabelSettings.ForeColor = $Global:colorAccent
    $LabelPreview.ForeColor = $Global:colorAccent
    $PicPreview.BackColor = $Global:colorPanel
    
    $BtnSave.BackColor = $Global:colorBtnGray; $BtnSave.ForeColor = $Global:colorFG
    $BtnLoad.BackColor = $Global:colorBtnGray; $BtnLoad.ForeColor = $Global:colorFG
    $BtnInfo.BackColor = $Global:colorBtnGray; $BtnInfo.ForeColor = $Global:colorFG
    $BtnTheme.BackColor = $Global:colorBtnGray; $BtnTheme.ForeColor = $Global:colorFG
    $BtnEdit.BackColor = $Global:colorBtnGray; $BtnEdit.ForeColor = $Global:colorFG
    $BtnRemove.BackColor = $Global:colorBtnGray; $BtnRemove.ForeColor = $Global:colorFG
    
    $LabelOps.ForeColor = $Global:colorAccent
    $LabelProj.ForeColor = $Global:colorAccent
    $LabelShortcuts.ForeColor = $Global:colorAccent
    $LabelKeys.ForeColor = $Global:colorFG

    # Slider Contrast Fix
    $SliderSpeed.BackColor = if ($isDark) { [Drawing.Color]::FromArgb(60,60,60) } else { $Global:colorBG }

    if ($BtnAction.Text -eq "PLAY MACRO") {
        $BtnAction.BackColor = $Global:colorSuccess
    } else {
        $BtnAction.BackColor = $Global:colorDanger
    }
    $BtnAction.ForeColor = [Drawing.Color]::White
    $StatusStrip.BackColor = $Global:colorBG; $StatusLabel.ForeColor = $Global:colorAccent
    $lblHUD.ForeColor = $Global:colorAccent
}

function Add-HoverEffect($btn, $type) {
    if ($null -eq $btn) { return }
    $btn.Add_MouseEnter({
        $hover = switch($type) {
            "Accent"  { [Drawing.Color]::FromArgb(100, 140, 255) }
            "Danger"  { [Drawing.Color]::FromArgb(255, 100, 100) }
            "Gray"    { if($Global:IsDarkMode){[Drawing.Color]::FromArgb(100,100,100)}else{[Drawing.Color]::FromArgb(200,200,200)} }
            Default   { $Global:colorAccent }
        }
        $this.BackColor = $hover
    }.GetNewClosure())

    $btn.Add_MouseLeave({
        $base = switch($type) {
            "Accent"  { $Global:colorAccent }
            "Danger"  { $Global:colorDanger }
            "Gray"    { $Global:colorBtnGray }
            Default   { $Global:colorAccent }
        }
        $this.BackColor = $base
    }.GetNewClosure())
}

# --- UI Elements ---

# COLUMN 1: MACRO ENGINE (Left)
$LabelSteps = New-Object Windows.Forms.Label
$LabelSteps.Text = "MACRO ENGINE"
$LabelSteps.Location = New-Object Drawing.Point(10, 10)
$LabelSteps.Font = $BoldFont
$LabelSteps.AutoSize = $true
$Form.Controls.Add($LabelSteps)

$ListSteps = New-Object Windows.Forms.ListBox
$ListSteps.Location = New-Object Drawing.Point(10, 35)
$ListSteps.Size = New-Object Drawing.Size(300, 480)
$ListSteps.BorderStyle = "None"
$Form.Controls.Add($ListSteps)

$BtnAction = New-Object Windows.Forms.Button
$BtnAction.Text = "RECORD SESSION (F8)"
$BtnAction.Location = New-Object Drawing.Point(10, 525)
$BtnAction.Size = New-Object Drawing.Size(300, 90)
$BtnAction.Font = New-Object Drawing.Font("Segoe UI", 16, [Drawing.FontStyle]::Bold)
$BtnAction.FlatStyle = "Flat"; $BtnAction.FlatAppearance.BorderSize = 0
$Form.Controls.Add($BtnAction)

# COLUMN 2: OPERATIONS & CONFIG (Middle)
$LabelOps = New-Object Windows.Forms.Label
$LabelOps.Text = "STEP OPERATIONS"
$LabelOps.Location = New-Object Drawing.Point(330, 10)
$LabelOps.Font = $BoldFont
$LabelOps.AutoSize = $true
$Form.Controls.Add($LabelOps)

$BtnAddType = New-Object Windows.Forms.Button
$BtnAddType.Text = "ADD TYPING"
$BtnAddType.Location = New-Object Drawing.Point(330, 35)
$BtnAddType.Size = New-Object Drawing.Size(180, 40)
$BtnAddType.FlatStyle = "Flat"; $BtnAddType.FlatAppearance.BorderSize = 0
Add-HoverEffect $BtnAddType "Accent"
$Form.Controls.Add($BtnAddType)

$BtnAddScroll = New-Object Windows.Forms.Button
$BtnAddScroll.Text = "ADD SCROLL"
$BtnAddScroll.Location = New-Object Drawing.Point(330, 85)
$BtnAddScroll.Size = New-Object Drawing.Size(180, 40)
$BtnAddScroll.FlatStyle = "Flat"; $BtnAddScroll.FlatAppearance.BorderSize = 0
Add-HoverEffect $BtnAddScroll "Accent"
$Form.Controls.Add($BtnAddScroll)

$BtnEdit = New-Object Windows.Forms.Button
$BtnEdit.Text = "EDIT STEP"
$BtnEdit.Location = New-Object Drawing.Point(330, 135)
$BtnEdit.Size = New-Object Drawing.Size(85, 35)
$BtnEdit.FlatStyle = "Flat"; $BtnEdit.FlatAppearance.BorderSize = 0
Add-HoverEffect $BtnEdit "Gray"
$Form.Controls.Add($BtnEdit)

$BtnRemove = New-Object Windows.Forms.Button
$BtnRemove.Text = "REMOVE"
$BtnRemove.Location = New-Object Drawing.Point(425, 135)
$BtnRemove.Size = New-Object Drawing.Size(85, 35)
$BtnRemove.FlatStyle = "Flat"; $BtnRemove.FlatAppearance.BorderSize = 0
Add-HoverEffect $BtnRemove "Gray"
$Form.Controls.Add($BtnRemove)

$BtnClear = New-Object Windows.Forms.Button
$BtnClear.Text = "CLEAR ALL"
$BtnClear.Location = New-Object Drawing.Point(330, 180)
$BtnClear.Size = New-Object Drawing.Size(180, 35)
$BtnClear.FlatStyle = "Flat"; $BtnClear.FlatAppearance.BorderSize = 0
Add-HoverEffect $BtnClear "Danger"
$Form.Controls.Add($BtnClear)

$LabelSettings = New-Object Windows.Forms.Label
$LabelSettings.Text = "SYSTEM CONFIG"
$LabelSettings.Location = New-Object Drawing.Point(330, 195)
$LabelSettings.Font = $BoldFont
$LabelSettings.AutoSize = $true
$Form.Controls.Add($LabelSettings)

$ChkShowNotify = New-Object Windows.Forms.CheckBox
$ChkShowNotify.Text = "On-Screen Messages"
$ChkShowNotify.Location = New-Object Drawing.Point(330, 220)
$ChkShowNotify.AutoSize = $true; $ChkShowNotify.Checked = $true
$Form.Controls.Add($ChkShowNotify)

$ChkUseVisual = New-Object Windows.Forms.CheckBox
$ChkUseVisual.Text = "Visual Matching AI"
$ChkUseVisual.Location = New-Object Drawing.Point(330, 245)
$ChkUseVisual.AutoSize = $true; $ChkUseVisual.Checked = $true
$Form.Controls.Add($ChkUseVisual)

$ChkMinimizeAll = New-Object Windows.Forms.CheckBox
$ChkMinimizeAll.Text = "Auto-Minimize Start"
$ChkMinimizeAll.Location = New-Object Drawing.Point(330, 270)
$ChkMinimizeAll.AutoSize = $true
$Form.Controls.Add($ChkMinimizeAll)

$LabelProj = New-Object Windows.Forms.Label
$LabelProj.Text = "PROJECT MANAGEMENT"
$LabelProj.Location = New-Object Drawing.Point(330, 310)
$LabelProj.Font = $BoldFont
$LabelProj.AutoSize = $true
$Form.Controls.Add($LabelProj)

$BtnSave = New-Object Windows.Forms.Button
$BtnSave.Text = "SAVE PROJECT"
$BtnSave.Location = New-Object Drawing.Point(330, 335)
$BtnSave.Size = New-Object Drawing.Size(180, 35)
$BtnSave.Visible = $false
$BtnSave.FlatStyle = "Flat"; $BtnSave.FlatAppearance.BorderSize = 0
Add-HoverEffect $BtnSave "Gray"
$Form.Controls.Add($BtnSave)

$BtnLoad = New-Object Windows.Forms.Button
$BtnLoad.Text = "LOAD PROJECT"
$BtnLoad.Location = New-Object Drawing.Point(330, 380)
$BtnLoad.Size = New-Object Drawing.Size(180, 35)
$BtnLoad.FlatStyle = "Flat"; $BtnLoad.FlatAppearance.BorderSize = 0
Add-HoverEffect $BtnLoad "Gray"
$Form.Controls.Add($BtnLoad)

$BtnTheme = New-Object Windows.Forms.Button
$BtnTheme.Text = "THEME"
$BtnTheme.Location = New-Object Drawing.Point(330, 583)
$BtnTheme.Size = New-Object Drawing.Size(85, 32)
$BtnTheme.FlatStyle = "Flat"; $BtnTheme.FlatAppearance.BorderSize = 0
Add-HoverEffect $BtnTheme "Gray"
$Form.Controls.Add($BtnTheme)

$BtnInfo = New-Object Windows.Forms.Button
$BtnInfo.Text = "CREDITS"
$BtnInfo.Location = New-Object Drawing.Point(425, 583)
$BtnInfo.Size = New-Object Drawing.Size(85, 32)
$BtnInfo.FlatStyle = "Flat"; $BtnInfo.FlatAppearance.BorderSize = 0
Add-HoverEffect $BtnInfo "Gray"
$Form.Controls.Add($BtnInfo)

# COLUMN 3: PREVIEW & PARAMETERS (Right)
$LabelPreview = New-Object Windows.Forms.Label
$LabelPreview.Text = "VISUAL ANCHOR"
$LabelPreview.Location = New-Object Drawing.Point(530, 10)
$LabelPreview.Font = $BoldFont
$LabelPreview.AutoSize = $true
$Form.Controls.Add($LabelPreview)

$PicPreview = New-Object Windows.Forms.PictureBox
$PicPreview.Location = New-Object Drawing.Point(530, 35)
$PicPreview.Size = New-Object Drawing.Size(200, 200)
$PicPreview.BorderStyle = "FixedSingle"
$PicPreview.SizeMode = "Zoom"
$Form.Controls.Add($PicPreview)

$LabelDelay = New-Object Windows.Forms.Label
$LabelDelay.Text = "EXECUTION PARAMETERS"
$LabelDelay.Location = New-Object Drawing.Point(530, 265)
$LabelDelay.Font = $BoldFont
$LabelDelay.AutoSize = $true
$Form.Controls.Add($LabelDelay)

$LabelWait = New-Object Windows.Forms.Label
$LabelWait.Text = "Extra Delay (ms):"
$LabelWait.Location = New-Object Drawing.Point(530, 300)
$LabelWait.AutoSize = $true
$Form.Controls.Add($LabelWait)

$TxtExtraDelay = New-Object Windows.Forms.TextBox
$TxtExtraDelay.Text = "0"
$TxtExtraDelay.Location = New-Object Drawing.Point(645, 297)
$TxtExtraDelay.Multiline = $true
$TxtExtraDelay.Size = New-Object Drawing.Size(70, 30)
$TxtExtraDelay.BorderStyle = "FixedSingle"
$TxtExtraDelay.Add_KeyDown({ if ($_.KeyCode -eq "Enter") { $_.SuppressKeyPress = $true } })
$Form.Controls.Add($TxtExtraDelay)

$LabelSpeed = New-Object Windows.Forms.Label
$LabelSpeed.Text = "Playback Speed: 1.0x"
$LabelSpeed.Location = New-Object Drawing.Point(530, 335)
$LabelSpeed.AutoSize = $true
$Form.Controls.Add($LabelSpeed)

$SliderSpeed = New-Object Windows.Forms.TrackBar
$SliderSpeed.Location = New-Object Drawing.Point(530, 360)
$SliderSpeed.Size = New-Object Drawing.Size(200, 45)
$SliderSpeed.Minimum = 1; $SliderSpeed.Maximum = 50; $SliderSpeed.Value = 10
$SliderSpeed.Add_Scroll({
    $val = $SliderSpeed.Value / 10.0
    $LabelSpeed.Text = "Playback Speed: {0:N1}x" -f $val
})
$Form.Controls.Add($SliderSpeed)

$LabelRepeat = New-Object Windows.Forms.Label
$LabelRepeat.Text = "Repeat Count:"
$LabelRepeat.Location = New-Object Drawing.Point(530, 415)
$LabelRepeat.AutoSize = $true
$Form.Controls.Add($LabelRepeat)

$TxtRepeat = New-Object Windows.Forms.TextBox
$TxtRepeat.Text = "1"
$TxtRepeat.Location = New-Object Drawing.Point(645, 412)
$TxtRepeat.Multiline = $true
$TxtRepeat.Size = New-Object Drawing.Size(70, 30)
$TxtRepeat.BorderStyle = "FixedSingle"
$TxtRepeat.Add_KeyDown({ if ($_.KeyCode -eq "Enter") { $_.SuppressKeyPress = $true } })
$Form.Controls.Add($TxtRepeat)

$ChkInfinite = New-Object Windows.Forms.CheckBox
$ChkInfinite.Text = "Infinite Loop (Stop with F8)"
$ChkInfinite.Location = New-Object Drawing.Point(530, 445)
$ChkInfinite.AutoSize = $true
$Form.Controls.Add($ChkInfinite)

# SHORTCUTS SECTION
$LabelShortcuts = New-Object Windows.Forms.Label
$LabelShortcuts.Text = "GLOBAL HOTKEYS"
$LabelShortcuts.Location = New-Object Drawing.Point(530, 485)
$LabelShortcuts.Font = $BoldFont
$LabelShortcuts.AutoSize = $true
$Form.Controls.Add($LabelShortcuts)

$LabelKeys = New-Object Windows.Forms.Label
$LabelKeys.Text = "F8:  START / STOP MACRO`nF9:  PAUSE / RESUME"
$LabelKeys.Location = New-Object Drawing.Point(530, 510)
$LabelKeys.Size = New-Object Drawing.Size(200, 50)
$Form.Controls.Add($LabelKeys)

$StatusStrip = New-Object Windows.Forms.StatusStrip
$StatusLabel = New-Object Windows.Forms.ToolStripStatusLabel
$StatusLabel.Text = "Ready"
$StatusStrip.Items.Add($StatusLabel) | Out-Null
$Form.Controls.Add($StatusStrip)

# --- Playback HUD Label ---
$lblHUD = New-Object Windows.Forms.Label
$lblHUD.Text = "No Macro Active"
$lblHUD.Location = New-Object Drawing.Point(30, 620)
$lblHUD.Size = New-Object Drawing.Size(690, 25)
$lblHUD.ForeColor = [Drawing.Color]::Gray
$lblHUD.Font = New-Object Drawing.Font("Segoe UI", 9, [Drawing.FontStyle]::Italic)
$lblHUD.TextAlign = "MiddleCenter"
$Form.Controls.Add($lblHUD)

# Final Theme Application
Apply-Theme $Global:IsDarkMode

function Show-Notification($msg) {
    if (-not $ChkShowNotify.Checked) { return }
    if ($null -eq $msg -or $msg -eq "") { return }
    $Toast = New-Object Windows.Forms.Form
    $Toast.Text = "Macro Notification"
    $Toast.Size = New-Object Drawing.Size(400, 50)
    $Toast.StartPosition = "Manual"
    $Toast.FormBorderStyle = "None"
    $Toast.TopMost = $true
    $Toast.BackColor = [Drawing.Color]::FromArgb(45, 45, 45)
    $Toast.Opacity = 0.85
    
    # Position bottom center
    $screen = [System.Windows.Forms.Screen]::PrimaryScreen.Bounds
    $x = ($screen.Width / 2) - 200
    $y = $screen.Height - 100
    $Toast.Location = New-Object Drawing.Point($x, $y)
    
    $Label = New-Object Windows.Forms.Label
    $Label.Text = $msg
    $Label.ForeColor = [Drawing.Color]::White
    $Label.Dock = "Fill"
    $Label.TextAlign = "MiddleCenter"
    $Label.Font = New-Object Drawing.Font("Segoe UI", 11, [Drawing.FontStyle]::Bold)
    $Toast.Controls.Add($Label)
    
    $Toast.Show()
    $Toast.Refresh()
    Start-Sleep -Milliseconds 800
    $Toast.Close()
}

function Get-ActiveWindowTitle() {
    $hWnd = [TestingMacroStudioProWin32]::GetForegroundWindow()
    if ($hWnd -eq [IntPtr]::Zero) { return "" }
    $sb = New-Object System.Text.StringBuilder(512)
    [TestingMacroStudioProWin32]::GetWindowText($hWnd, $sb, $sb.Capacity) | Out-Null
    return $sb.ToString().Trim()
}

# --- Visual Alignment Engine ---
function Get-ScreenSnippet($x, $y, $w = 50, $h = 50) {
    try {
        $bmp = New-Object Drawing.Bitmap($w, $h)
        $graphics = [Drawing.Graphics]::FromImage($bmp)
        # Handle screen edges to prevent out-of-bounds exceptions
        $srcX = [Math]::Max(0, $x - [int]($w/2))
        $srcY = [Math]::Max(0, $y - [int]($h/2))
        $graphics.CopyFromScreen($srcX, $srcY, 0, 0, $bmp.Size)
        $graphics.Dispose()
        return $bmp
    } catch { return $null }
}

function Image-ToBase64($bmp) {
    if ($null -eq $bmp) { return "" }
    $ms = New-Object System.IO.MemoryStream
    $bmp.Save($ms, [Drawing.Imaging.ImageFormat]::Png)
    $bytes = $ms.ToArray()
    $ms.Dispose()
    return [Convert]::ToBase64String($bytes)
}

function Base64-ToImage($base64) {
    if ([string]::IsNullOrWhiteSpace($base64)) { return $null }
    try {
        $bytes = [Convert]::FromBase64String($base64)
        $ms = New-Object System.IO.MemoryStream($bytes)
        $img = [Drawing.Image]::FromStream($ms)
        # Force a clone into a new bitmap to guarantee stream independence
        $bmp = New-Object Drawing.Bitmap($img.Width, $img.Height)
        $g = [Drawing.Graphics]::FromImage($bmp)
        $g.SmoothingMode = "HighQuality"
        $g.InterpolationMode = "HighQualityBicubic"
        $g.DrawImage($img, 0, 0, $img.Width, $img.Height)
        $g.Dispose()
        $img.Dispose()
        $ms.Dispose()
        return $bmp
    } catch { return $null }
}

function Find-ImageSnippet($largeBmp, $smallBmp) {
    if ($null -eq $largeBmp -or $null -eq $smallBmp) { return $null }
    for ($x = 0; $x -le ($largeBmp.Width - $smallBmp.Width); $x++) {
        for ($y = 0; $y -le ($largeBmp.Height - $smallBmp.Height); $y++) {
            $match = $true
            $checkPoints = @(
                @($x, $y), 
                @($x + $smallBmp.Width - 1, $y), 
                @($x, $y + $smallBmp.Height - 1), 
                @($x + $smallBmp.Width - 1, $y + $smallBmp.Height - 1),
                @($x + [int]($smallBmp.Width/2), $y + [int]($smallBmp.Height/2))
            )
            foreach ($p in $checkPoints) {
                if ($largeBmp.GetPixel($p[0], $p[1]).ToArgb() -ne $smallBmp.GetPixel($p[0]-$x, $p[1]-$y).ToArgb()) {
                    $match = $false; break
                }
            }
            if ($match) { return New-Object Drawing.Point($x, $y) }
        }
    }
    return $null
}

function Refresh-StepList {
    # Prevent GUI flicker and hang during large list updates
    $ListSteps.BeginUpdate()
    $ListSteps.Items.Clear()
    $count = 1
    foreach ($step in $Global:MacroSteps) {
        if ($step.ActionType -eq "Move") { continue } # Skip background moves for clean UI
        $appTag = if ($step.WindowTitle) { "[$($step.WindowTitle)] " } else { "" }
        $displayText = switch($step.ActionType) {
            "Click"       { "Step ${count}: ${appTag}Click at ($($step.ScreenX), $($step.ScreenY))" }
            "RightClick"  { "Step ${count}: ${appTag}Right-Click at ($($step.ScreenX), $($step.ScreenY))" }
            "DoubleClick" { "Step ${count}: ${appTag}Double-Click at ($($step.ScreenX), $($step.ScreenY))" }
            "Scroll"      { "Step ${count}: ${appTag}Scroll $($step.TextToType)" }
            "Type"        { "Step ${count}: ${appTag}Type '$($step.TextToType)'" }
            Default       { "Step ${count}: ${appTag}Action: $($step.ActionType)" }
        }
        $displayText += " - Wait $($step.WaitTimeMS) ms"
        $ListSteps.Items.Add($displayText) | Out-Null
        $count++
    }
    $ListSteps.EndUpdate()
    
    # Update visibility of buttons
    if ($Global:MacroSteps.Count -gt 0) {
        $BtnAction.Text = "PLAY MACRO"
        $BtnAction.BackColor = $Global:colorSuccess
        $BtnSave.Visible = $true
    } else {
        $BtnAction.Text = "RECORD SESSION (F8)"
        $BtnAction.BackColor = $Global:colorDanger
        $BtnSave.Visible = $false
    }
}

function Add-MacroStep($action, $x, $y, $text, $delay, $image = "", $winTitle = "") {
    $step = [PSCustomObject]@{
        "Instructions" = switch($action) {
            "Click"       { "Move mouse to $x, $y and Left Click" }
            "RightClick"  { "Move mouse to $x, $y and Right Click" }
            "DoubleClick" { "Move mouse to $x, $y and Double Click" }
            "Move"        { "Move mouse to $x, $y" }
            "Type"        { "Type the message: '$text'" }
            "Scroll"      { "Scroll mouse wheel by $text units" }
            Default       { "Action: $action" }
        }
        "ActionType" = $action
        "ScreenX" = $x
        "ScreenY" = $y
        "TextToType" = $text
        "WaitTimeMS" = [int]$delay
        "AnchorImage" = $image
        "WindowTitle" = $winTitle
    }
    $Global:MacroSteps.Add($step)
    Refresh-StepList
}

function Start-Recording {
    if ($ChkMinimizeAll.Checked) {
        (New-Object -ComObject Shell.Application).MinimizeAll()
        Start-Sleep -Milliseconds 500
    }
    $Form.WindowState = "Minimized"
    $StatusLabel.Text = "RECORDING... (Press F8 to Stop)"
    Show-Notification "STARTING SEAMLESS RECORDING"
    
    $Global:MacroSteps.Clear()
    $ListSteps.Items.Clear()
    $BtnSave.Visible = $false
    
    $lastTime = Get-Date
    $leftDown = $false
    $rightDown = $false
    $lastClickTime = 0
    $lastClickPos = [Drawing.Point]::Empty
    $lastMousePos = [Windows.Forms.Cursor]::Position
    $keyDown = @{} 
    
    $keyMap = @{
        0x01 = "LeftClick"; 0x11 = "Control"; 0x12 = "Menu"
        0x10 = "Shift"; 0x08 = "Back"; 0x0D = "Enter"
        0x20 = "Space"; 0x09 = "Tab"; 0x1B = "Escape"
        0x21 = "PageUp"; 0x22 = "PageDown"; 0x25 = "Left"
        0x26 = "Up"; 0x27 = "Right"; 0x28 = "Down"
    }
    for($vk=0x30;$vk -le 0x5A;$vk++) { $keyMap[$vk] = [char]$vk }

    while ([TestingMacroStudioProWin32]::GetAsyncKeyState(0x77) -band 0x8000) { Start-Sleep -Milliseconds 10 }

    $isPaused = $false
    while ($true) {
        if ([TestingMacroStudioProWin32]::GetAsyncKeyState(0x77) -band 0x8000) { break }
        if ([TestingMacroStudioProWin32]::GetAsyncKeyState(0x78) -band 0x8000) {
            $isPaused = -not $isPaused
            if ($isPaused) {
                Show-Notification "RECORDING PAUSED"
                $StatusLabel.Text = "RECORDING PAUSED"
            } else {
                Show-Notification "RECORDING RESUMED"
                $StatusLabel.Text = "RECORDING..."
                $lastTime = Get-Date
            }
            while ([TestingMacroStudioProWin32]::GetAsyncKeyState(0x78) -band 0x8000) { Start-Sleep -Milliseconds 10 }
        }
        if ($isPaused) { Start-Sleep -Milliseconds 100; continue }

        $currentTime = Get-Date
        $delta = ($currentTime - $lastTime).TotalMilliseconds
        
        # --- High-Speed Typing Buffer ---
        $isAnyKeyDown = $false
        foreach($vk in $keyMap.Keys) {
            if ($vk -eq 0x01) { continue } # Handle Mouse Separately
            $isDown = [TestingMacroStudioProWin32]::GetAsyncKeyState($vk) -band 0x8000
            if ($isDown -and -not $keyDown[$vk]) {
                $isAnyKeyDown = $true
                $keyName = $keyMap[$vk]
                
                # Special Keys trigger immediate flush
                # Arrows (0x25-0x28) are now special to avoid being typed as strings
                $isSpecial = ($vk -eq 0x08 -or $vk -eq 0x0D -or $vk -eq 0x09 -or $vk -eq 0x1B -or $vk -eq 0x11 -or $vk -eq 0x12 -or ($vk -ge 0x25 -and $vk -le 0x28) -or $isPaused)
                
                if ($isSpecial) {
                    # Flush buffer before special key
                    if ($Global:TypeBuffer -ne "") {
                        # Capture image where mouse is when typing finishes (or starts)
                        $pos = [Windows.Forms.Cursor]::Position
                        $bmp = Get-ScreenSnippet $pos.X $pos.Y
                        $img = Image-ToBase64 $bmp; if($bmp){$bmp.Dispose()}
                        Add-MacroStep "Type" 0 0 $Global:TypeBuffer $Global:BufferTime $img $Global:BufferWin
                        $Global:TypeBuffer = ""
                    }
                    
                    $sendKey = switch($keyName) {
                        "Back"     { "{BACKSPACE}" }; "Enter"    { "{ENTER}" }
                        "Space"    { " " }; "Tab"      { "{TAB}" }
                        "Escape"   { "{ESC}" }; "PageUp"   { "{PGUP}" }
                        "PageDown" { "{PGDN}" }; "Up"       { "{UP}" }
                        "Down"     { "{DOWN}" }; "Left"     { "{LEFT}" }
                        "Right"    { "{RIGHT}" }
                        Default    { $keyName.ToString().ToLower() }
                    }
                    
                    $modifierPrefix = ""
                    if ([TestingMacroStudioProWin32]::GetAsyncKeyState(0x11) -band 0x8000) { $modifierPrefix += "^" }
                    if ([TestingMacroStudioProWin32]::GetAsyncKeyState(0x12) -band 0x8000) { $modifierPrefix += "%" }
                    if ([TestingMacroStudioProWin32]::GetAsyncKeyState(0x10) -band 0x8000) { $modifierPrefix += "+" }
                    
                    $finalKey = $modifierPrefix + $sendKey
                    Add-MacroStep "Type" 0 0 $finalKey $delta "" (Get-ActiveWindowTitle)
                } else {
                    # Append to buffer
                    if ($Global:TypeBuffer -eq "") { 
                        $Global:BufferTime = $delta 
                        $Global:BufferWin = Get-ActiveWindowTitle
                        # We could capture image here, but let's do it at flush for simplicity
                    }
                    $char = switch($keyName) {
                        "Space" { " " }
                        Default { $keyName.ToString().ToLower() }
                    }
                    $Global:TypeBuffer += $char
                }
                $lastTime = $currentTime
            }
            $keyDown[$vk] = $isDown
        }

        # --- Mouse Capture ---
        $pos = [Windows.Forms.Cursor]::Position
        
        # 1. Capture Movement (Throttled Move)
        if ($pos.X -ne $lastMousePos.X -or $pos.Y -ne $lastMousePos.Y) {
            Add-MacroStep "Move" $pos.X $pos.Y "" $delta "" ""
            $lastMousePos = $pos
            $lastTime = $currentTime
            $delta = 0 # Reset delta for immediately following action
        }

        # 2. Left Click (with Double Click Logic)
        $isLeftDown = [TestingMacroStudioProWin32]::GetAsyncKeyState(0x01) -band 0x8000
        if ($isLeftDown -and -not $leftDown) {
            if ($Global:TypeBuffer -ne "") {
                $bmp_m = Get-ScreenSnippet $pos.X $pos.Y; $img_m = Image-ToBase64 $bmp_m; if($bmp_m){$bmp_m.Dispose()}
                Add-MacroStep "Type" 0 0 $Global:TypeBuffer $Global:BufferTime $img_m $Global:BufferWin
                $Global:TypeBuffer = ""
            }
            
            $currentTimeMS = [DateTimeOffset]::Now.ToUnixTimeMilliseconds()
            $isDoubleClick = ($currentTimeMS - $lastClickTime -lt 400) -and ([Math]::Abs($pos.X - $lastClickPos.X) -lt 5) -and ([Math]::Abs($pos.Y - $lastClickPos.Y) -lt 5)
            
            if ($isDoubleClick -and $Global:MacroSteps.Count -gt 0) {
                # Convert previous single click to double click
                $prev = $Global:MacroSteps[$Global:MacroSteps.Count - 1]
                if ($prev.ActionType -eq "Click") {
                    $prev.ActionType = "DoubleClick"
                    $prev.Instructions = "Move mouse to $($pos.X), $($pos.Y) and Double Click"
                    Refresh-StepList
                } else {
                    Add-MacroStep "DoubleClick" $pos.X $pos.Y "" $delta "" (Get-ActiveWindowTitle)
                }
            } else {
                $bmp = Get-ScreenSnippet $pos.X $pos.Y; $img = Image-ToBase64 $bmp; if($bmp){$bmp.Dispose()}
                Add-MacroStep "Click" $pos.X $pos.Y "" $delta $img (Get-ActiveWindowTitle)
            }
            
            $lastClickTime = $currentTimeMS
            $lastClickPos = $pos
            $lastTime = Get-Date
            Show-Notification "CLICK DETECTED"
        }
        $leftDown = $isLeftDown

        # 3. Right Click
        $isRightDown = [TestingMacroStudioProWin32]::GetAsyncKeyState(0x02) -band 0x8000
        if ($isRightDown -and -not $rightDown) {
            $bmp = Get-ScreenSnippet $pos.X $pos.Y; $img = Image-ToBase64 $bmp; if($bmp){$bmp.Dispose()}
            Add-MacroStep "RightClick" $pos.X $pos.Y "" $delta $img (Get-ActiveWindowTitle)
            $lastTime = Get-Date
            Show-Notification "RIGHT CLICK"
        }
        $rightDown = $isRightDown

        Start-Sleep -Milliseconds 15
        [Windows.Forms.Application]::DoEvents()
    }
    
    # Final Flush
    if ($Global:TypeBuffer -ne "") {
        $pos_f = [Windows.Forms.Cursor]::Position
        $bmp_f = Get-ScreenSnippet $pos_f.X $pos_f.Y
        $img_f = Image-ToBase64 $bmp_f; if($bmp_f){$bmp_f.Dispose()}
        Add-MacroStep "Type" 0 0 $Global:TypeBuffer $Global:BufferTime $img_f $Global:BufferWin
        $Global:TypeBuffer = ""
    }
    
    $Form.WindowState = "Normal"
    $StatusLabel.Text = "Recording Finished. Steps: $($Global:MacroSteps.Count)"
    Show-Notification "RECORDING STOPPED"
}

function Start-Playback {
    $repeats = if ($ChkInfinite.Checked) { 999999 } else { [int]$TxtRepeat.Text }
    $extraDelay = [int]$TxtExtraDelay.Text
    $speed = $SliderSpeed.Value / 10.0
    if ($speed -le 0) { $speed = 1.0 }
    
    $HUDLabel.Text = "INITIALIZING MACRO..."
    $FloatingHUD.Show()
    
    $Form.WindowState = "Minimized"
    Start-Sleep -Seconds 2
    
    $stopMacro = $false
    for ($i = 0; $i -lt $repeats; $i++) {
        if ($stopMacro) { break }
        $StatusLabel.Text = "Playing Loop $($i + 1)..."
        
        $stepCount = 1
        foreach ($step in $Global:MacroSteps) {
            if ([TestingMacroStudioProWin32]::GetAsyncKeyState(0x77) -band 0x8000) { $stopMacro = $true; break }
            
            $hudText = "ACTION ($stepCount): $($step.ActionType)"
            if ($step.ActionType -eq "Type") { $hudText = "TYPE: '$($step.TextToType)'" }
            if ($step.WindowTitle) { $hudText += " in $($step.WindowTitle)" }
            $lblHUD.Text = $hudText; $lblHUD.Refresh()
            $HUDLabel.Text = $hudText; $FloatingHUD.Refresh()
            
            $StatusLabel.Text = "Step ${stepCount}: $($step.ActionType) in $($step.WindowTitle)"
            
            # --- Smart Window Wait ---
            if (-not [string]::IsNullOrEmpty($step.WindowTitle)) {
                $waitedForWin = 0
                $targetWin = $step.WindowTitle
                # Clean up title for matching (often browser titles contain extra parts)
                $matchPattern = if ($targetWin -like "* - *") { $targetWin.Split("-")[-1].Trim() } else { $targetWin }
                
                while ($waitedForWin -lt 6000) { 
                    $currentWin = Get-ActiveWindowTitle
                    # Robust check: Exact match OR partial match for common parts
                    if ($currentWin -eq $targetWin -or $currentWin -match [regex]::Escape($matchPattern)) { 
                        # Force foreground strictly as requested
                        $hWnd = [TestingMacroStudioProWin32]::GetForegroundWindow()
                        [TestingMacroStudioProWin32]::SetForegroundWindow($hWnd) | Out-Null
                        break 
                    }
                    
                    $StatusLabel.Text = "WAITING FOR: $targetWin ($waitedForWin ms)"
                    if ([TestingMacroStudioProWin32]::GetAsyncKeyState(0x77) -band 0x8000) { $stopMacro = $true; break }
                    Start-Sleep -Milliseconds 150
                    $waitedForWin += 150
                }
                if ($stopMacro) { break }
            }

            if ($step.ActionType -eq "Click") {
                $targetX = $step.ScreenX; $targetY = $step.ScreenY
                if ($ChkUseVisual.Checked -and -not [string]::IsNullOrEmpty($step.AnchorImage)) {
                    $anchor = Base64-ToImage $step.AnchorImage
                    if ($null -ne $anchor) {
                        $searchArea = Get-ScreenSnippet $targetX $targetY 100 100
                        $matchPos = Find-ImageSnippet $searchArea $anchor
                        if ($null -ne $matchPos) {
                            $targetX = ($targetX - 50) + $matchPos.X + 25
                            $targetY = ($targetY - 50) + $matchPos.Y + 25
                        }
                        $anchor.Dispose(); $searchArea.Dispose()
                    }
                }
                [TestingMacroStudioProWin32]::SetCursorPos($targetX, $targetY)
                Start-Sleep -Milliseconds 150
                [TestingMacroStudioProWin32]::mouse_event([TestingMacroStudioProWin32]::MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
                [TestingMacroStudioProWin32]::mouse_event([TestingMacroStudioProWin32]::MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
            }
            elseif ($step.ActionType -eq "RightClick") {
                [TestingMacroStudioProWin32]::SetCursorPos($step.ScreenX, $step.ScreenY)
                Start-Sleep -Milliseconds 100
                [TestingMacroStudioProWin32]::mouse_event([TestingMacroStudioProWin32]::MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, 0)
                [TestingMacroStudioProWin32]::mouse_event([TestingMacroStudioProWin32]::MOUSEEVENTF_RIGHTUP, 0, 0, 0, 0)
            }
            elseif ($step.ActionType -eq "DoubleClick") {
                [TestingMacroStudioProWin32]::SetCursorPos($step.ScreenX, $step.ScreenY)
                Start-Sleep -Milliseconds 100
                [TestingMacroStudioProWin32]::mouse_event([TestingMacroStudioProWin32]::MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
                [TestingMacroStudioProWin32]::mouse_event([TestingMacroStudioProWin32]::MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
                Start-Sleep -Milliseconds 50
                [TestingMacroStudioProWin32]::mouse_event([TestingMacroStudioProWin32]::MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
                [TestingMacroStudioProWin32]::mouse_event([TestingMacroStudioProWin32]::MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
            }
            elseif ($step.ActionType -eq "Move") {
                [TestingMacroStudioProWin32]::SetCursorPos($step.ScreenX, $step.ScreenY)
            }
            elseif ($step.ActionType -eq "Type") {
                # Excel Fix: Ensure fresh modifier state for navigation keys
                if ($step.TextToType -match "{(LEFT|RIGHT|UP|DOWN)}") {
                    [TestingMacroStudioProWin32]::keybd_event(0x11, 0, 0x0002, 0) # Ctrl Up
                    [TestingMacroStudioProWin32]::keybd_event(0x10, 0, 0x0002, 0) # Shift Up
                }
                [System.Windows.Forms.SendKeys]::SendWait($step.TextToType)
            }
            elseif ($step.ActionType -eq "Scroll") {
                [TestingMacroStudioProWin32]::mouse_event([TestingMacroStudioProWin32]::MOUSEEVENTF_WHEEL, 0, 0, [int]$step.TextToType, 0)
            }
            
            $finalWait = ([int]($step.WaitTimeMS / $speed)) + $extraDelay
            if ($finalWait -lt 50) { $finalWait = 50 }
            
            $waited = 0
            while ($waited -lt $finalWait) {
                if ([TestingMacroStudioProWin32]::GetAsyncKeyState(0x77) -band 0x8000) { $stopMacro = $true; break }
                if ([TestingMacroStudioProWin32]::GetAsyncKeyState(0x78) -band 0x8000) {
                    Show-Notification "MACRO PAUSED"
                    $StatusLabel.Text = "PLAYBACK PAUSED"
                    while ([TestingMacroStudioProWin32]::GetAsyncKeyState(0x78) -band 0x8000) { Start-Sleep -Milliseconds 10 }
                    while (-not ([TestingMacroStudioProWin32]::GetAsyncKeyState(0x78) -band 0x8000)) { 
                        if ([TestingMacroStudioProWin32]::GetAsyncKeyState(0x77) -band 0x8000) { $stopMacro = $true; break }
                        Start-Sleep -Milliseconds 100 
                    }
                    if ($stopMacro) { break }
                    Show-Notification "MACRO RESUMED"
                    $StatusLabel.Text = "Playing Loop $($i + 1)..."
                    while ([TestingMacroStudioProWin32]::GetAsyncKeyState(0x78) -band 0x8000) { Start-Sleep -Milliseconds 10 }
                }
                Start-Sleep -Milliseconds 100
                $waited += 100
            }
            $stepCount++
        }
        if (-not $stopMacro) { Show-Notification "Loop $($i+1) Completed" }
    }
    
    if (-not $stopMacro) {
        $HUDLabel.Text = "MACRO COMPLETED"
        $FloatingHUD.BackColor = $Global:colorSuccess
        $FloatingHUD.Refresh()
        Start-Sleep -Seconds 1.5
        $FloatingHUD.BackColor = [Drawing.Color]::FromArgb(40, 40, 40)
    }
    
    $FloatingHUD.Hide()
    $Form.WindowState = "Normal"
    $StatusLabel.Text = if ($stopMacro) { "Macro Stopped" } else { "Playback Complete" }
}

$BtnAction.Add_Click({
    if ($BtnAction.Text -eq "RECORD SESSION (F8)") {
        Start-Recording
    } else {
        Start-Playback
    }
})

$BtnAddType.Add_Click({
    $input = [Microsoft.VisualBasic.Interaction]::InputBox("Enter text to type:", "Typing Step", "")
    if ($input) { 
        $pos = [Windows.Forms.Cursor]::Position
        $bmp = Get-ScreenSnippet $pos.X $pos.Y 60 60
        $img = Image-ToBase64 $bmp; if($bmp){$bmp.Dispose()}
        $win = Get-ActiveWindowTitle
        Add-MacroStep "Type" 0 0 $input 0 $img $win
    }
})

$BtnAddScroll.Add_Click({
    $input = [Microsoft.VisualBasic.Interaction]::InputBox("Enter scroll amount (e.g. -120 to scroll down, 120 for up):", "Scroll Step", "-120")
    if ($input) { 
        $pos = [Windows.Forms.Cursor]::Position
        $bmp = Get-ScreenSnippet $pos.X $pos.Y 60 60
        $img = Image-ToBase64 $bmp; if($bmp){$bmp.Dispose()}
        $win = Get-ActiveWindowTitle
        Add-MacroStep "Scroll" 0 0 $input 0 $img $win
    }
})

$BtnClear.Add_Click({
    $MacroSteps.Clear()
    Refresh-StepList
    $StatusLabel.Text = "All Steps Cleared"
})

$BtnRemove.Add_Click({
    $idx = $ListSteps.SelectedIndex
    if ($idx -ge 0) {
        $MacroSteps.RemoveAt($idx)
        Refresh-StepList
        $StatusLabel.Text = "Step Removed"
    } else {
        [System.Windows.Forms.MessageBox]::Show("Please select a step to remove.", "Selection Required")
    }
})

function Show-StepEditor($idx) {
    if ($idx -lt 0 -or $idx -ge $Global:MacroSteps.Count) { return }
    $step = $Global:MacroSteps[$idx]

    $EditForm = New-Object Windows.Forms.Form
    $EditForm.Text = "Edit Macro Step"
    $EditForm.Size = New-Object Drawing.Size(450, 280)
    $EditForm.BackColor = $Global:colorBG
    $EditForm.ForeColor = $Global:colorFG
    $EditForm.StartPosition = "CenterParent"
    $EditForm.FormBorderStyle = "FixedDialog"
    $EditForm.MaximizeBox = $false

    $LabelTitle = New-Object Windows.Forms.Label
    $LabelTitle.Text = "EDIT STEP: $($step.ActionType)"
    $LabelTitle.Font = $BoldFont
    $LabelTitle.ForeColor = $Global:colorAccent
    $LabelTitle.Location = New-Object Drawing.Point(20, 20)
    $LabelTitle.AutoSize = $true
    $EditForm.Controls.Add($LabelTitle)

    # Primary Input
    $LabelVal = New-Object Windows.Forms.Label
    $LabelVal.Text = if ($step.ActionType -eq "Click") { "Click Metadata (Intent):" } else { "Text / Units to Send:" }
    $LabelVal.Location = New-Object Drawing.Point(20, 55)
    $LabelVal.AutoSize = $true
    $EditForm.Controls.Add($LabelVal)

    $TxtVal = New-Object Windows.Forms.TextBox
    $TxtVal.Text = if ($step.ActionType -eq "Click") { $step.Instructions } else { $step.TextToType }
    $TxtVal.Location = New-Object Drawing.Point(20, 80)
    $TxtVal.Size = New-Object Drawing.Size(390, 30)
    $TxtVal.BackColor = $Global:colorPanel
    $TxtVal.ForeColor = $Global:colorFG
    $TxtVal.BorderStyle = "FixedSingle"
    $EditForm.Controls.Add($TxtVal)

    # Delay Input
    $LabelDelay = New-Object Windows.Forms.Label
    $LabelDelay.Text = "Wait Time / Delay (ms):"
    $LabelDelay.Location = New-Object Drawing.Point(20, 125)
    $LabelDelay.AutoSize = $true
    $EditForm.Controls.Add($LabelDelay)

    $TxtWait = New-Object Windows.Forms.TextBox
    $TxtWait.Text = $step.WaitTimeMS.ToString()
    $TxtWait.Location = New-Object Drawing.Point(20, 150)
    $TxtWait.Size = New-Object Drawing.Size(150, 30)
    $TxtWait.BackColor = $Global:colorPanel
    $TxtWait.ForeColor = $Global:colorFG
    $TxtWait.BorderStyle = "FixedSingle"
    $EditForm.Controls.Add($TxtWait)

    # Buttons
    $BtnOk = New-Object Windows.Forms.Button
    $BtnOk.Text = "SAVE CHANGES"
    $BtnOk.Location = New-Object Drawing.Point(20, 195)
    $BtnOk.Size = New-Object Drawing.Size(180, 40)
    $BtnOk.BackColor = $Global:colorSuccess
    $BtnOk.ForeColor = [Drawing.Color]::White
    $BtnOk.FlatStyle = "Flat"; $BtnOk.FlatAppearance.BorderSize = 0
    $BtnOk.Add_Click({
        if ($step.ActionType -eq "Click") {
            $step.Instructions = $TxtVal.Text
        } else {
            $step.TextToType = $TxtVal.Text
        }
        if ($TxtWait.Text -as [int]) {
            $step.WaitTimeMS = [int]$TxtWait.Text
        }
        $EditForm.Tag = "OK"
        $EditForm.Close()
    })
    $EditForm.Controls.Add($BtnOk)

    $BtnCancel = New-Object Windows.Forms.Button
    $BtnCancel.Text = "CANCEL"
    $BtnCancel.Location = New-Object Drawing.Point(230, 195)
    $BtnCancel.Size = New-Object Drawing.Size(180, 40)
    $BtnCancel.BackColor = $Global:colorBtnGray
    $BtnCancel.ForeColor = $Global:colorFG
    $BtnCancel.FlatStyle = "Flat"; $BtnCancel.FlatAppearance.BorderSize = 0
    $BtnCancel.Add_Click({ $EditForm.Close() })
    $EditForm.Controls.Add($BtnCancel)

    $EditForm.ShowDialog()
    return $EditForm.Tag
}

$BtnEdit.Add_Click({
    $idx = $ListSteps.SelectedIndex
    if ($idx -ge 0) {
        $res = Show-StepEditor $idx
        if ($res -eq "OK") {
            Refresh-StepList
            $StatusLabel.Text = "Step Updated"
        }
    } else {
        [System.Windows.Forms.MessageBox]::Show("Please select a step to edit.", "Selection Required")
    }
})


$BtnTheme.Add_Click({
    $Global:IsDarkMode = -not $Global:IsDarkMode
    Apply-Theme $Global:IsDarkMode
})

$BtnInfo.Add_Click({
    $AboutForm = New-Object Windows.Forms.Form
    $AboutForm.Text = "About Macro Studio Pro"
    $AboutForm.Size = New-Object Drawing.Size(420, 360)
    $AboutForm.BackColor = $colorBG
    $AboutForm.ForeColor = $colorFG
    $AboutForm.StartPosition = "CenterParent"
    $AboutForm.FormBorderStyle = "FixedDialog"
    $AboutForm.MaximizeBox = $false

    $AboutTitle = New-Object Windows.Forms.Label
    $AboutTitle.Text = "MACRO STUDIO PRO"
    $AboutTitle.Font = New-Object Drawing.Font("Segoe UI", 18, [Drawing.FontStyle]::Bold)
    $AboutTitle.ForeColor = $colorAccent
    $AboutTitle.AutoSize = $true
    $AboutTitle.Location = New-Object Drawing.Point(20, 20)
    $AboutForm.Controls.Add($AboutTitle)

    $AboutAuthor = New-Object Windows.Forms.Label
    $AboutAuthor.Text = "Created by: Saurabh Yadav"
    $AboutAuthor.Font = $BoldFont
    $AboutAuthor.AutoSize = $true
    $AboutAuthor.Location = New-Object Drawing.Point(25, 65)
    $AboutForm.Controls.Add($AboutAuthor)

    # Email Link
    $LinkEmail = New-Object Windows.Forms.LinkLabel
    $LinkEmail.Text = "Email: skyhedevil@gmail.com"
    $LinkEmail.Location = New-Object Drawing.Point(25, 95)
    $LinkEmail.Size = New-Object Drawing.Size(300, 25)
    $LinkEmail.LinkColor = $colorAccent
    $LinkEmail.ActiveLinkColor = [Drawing.Color]::White
    $LinkEmail.Add_LinkClicked({ [Diagnostics.Process]::Start("mailto:skyhedevil@gmail.com") })
    $AboutForm.Controls.Add($LinkEmail)

    # GitHub Link
    $LinkGit = New-Object Windows.Forms.LinkLabel
    $LinkGit.Text = "GitHub: github.com/skyhedevil"
    $LinkGit.Location = New-Object Drawing.Point(25, 125)
    $LinkGit.Size = New-Object Drawing.Size(300, 25)
    $LinkGit.LinkColor = $colorAccent
    $LinkGit.ActiveLinkColor = [Drawing.Color]::White
    $LinkGit.Add_LinkClicked({ [Diagnostics.Process]::Start("https://github.com/skyhedevil") })
    $AboutForm.Controls.Add($LinkGit)

    # LinkedIn Link
    $LinkLinkedIn = New-Object Windows.Forms.LinkLabel
    $LinkLinkedIn.Text = "LinkedIn: Saurabh Yadav"
    $LinkLinkedIn.Location = New-Object Drawing.Point(25, 155)
    $LinkLinkedIn.Size = New-Object Drawing.Size(300, 25)
    $LinkLinkedIn.LinkColor = $colorAccent
    $LinkLinkedIn.ActiveLinkColor = [Drawing.Color]::White
    $LinkLinkedIn.Add_LinkClicked({ [Diagnostics.Process]::Start("https://www.linkedin.com/in/professionalsky/") })
    $AboutForm.Controls.Add($LinkLinkedIn)

    $AboutVer = New-Object Windows.Forms.Label
    $AboutVer.Text = "v1.5.0 - Professional Studio Edition`nBuilt for high-performance automation."
    $AboutVer.Location = New-Object Drawing.Point(25, 200)
    $AboutVer.Size = New-Object Drawing.Size(350, 45)
    $AboutVer.ForeColor = [Drawing.Color]::Gray
    $AboutForm.Controls.Add($AboutVer)

    $BtnClose = New-Object Windows.Forms.Button
    $BtnClose.Text = "CLOSE"
    $BtnClose.Location = New-Object Drawing.Point(145, 260)
    $BtnClose.Size = New-Object Drawing.Size(120, 40)
    $BtnClose.BackColor = $colorBtnGray
    $BtnClose.FlatStyle = "Flat"; $BtnClose.FlatAppearance.BorderSize = 0
    $BtnClose.Add_Click({ $AboutForm.Close() })
    $AboutForm.Controls.Add($BtnClose)

    $AboutForm.ShowDialog()
})

$BtnSave.Add_Click({
    $saveFile = New-Object Windows.Forms.SaveFileDialog
    $saveFile.Filter = "Macro files (*.json)|*.json"
    if ($saveFile.ShowDialog() -eq "OK") {
        $Global:MacroSteps | ConvertTo-Json -Depth 5 | Out-File $saveFile.FileName
        $StatusLabel.Text = "Macro Saved"
    }
})

$BtnLoad.Add_Click({
    $loadFile = New-Object Windows.Forms.OpenFileDialog
    $loadFile.Filter = "Macro files (*.json)|*.json"
    if ($loadFile.ShowDialog() -eq "OK") {
        $Global:MacroSteps.Clear()
        $ListSteps.Items.Clear()
        $data = Get-Content $loadFile.FileName | ConvertFrom-Json
        foreach ($item in $data) {
            Add-MacroStep $item.ActionType $item.ScreenX $item.ScreenY $item.TextToType $item.WaitTimeMS $item.AnchorImage $item.WindowTitle
        }
        $StatusLabel.Text = "Macro Loaded"
    }
})

# --- Event Binding ---
$ListSteps.Add_SelectedIndexChanged({
    $idx = $ListSteps.SelectedIndex
    if ($idx -ge 0 -and $idx -lt $Global:MacroSteps.Count) {
        $step = $Global:MacroSteps[$idx]
        if (-not [string]::IsNullOrEmpty($step.AnchorImage)) {
            $bmp = Base64-ToImage $step.AnchorImage
            if ($null -ne $bmp) {
                $oldImg = $PicPreview.Image
                $PicPreview.Image = $bmp
                if ($oldImg) { $oldImg.Dispose() }
            } else { $PicPreview.Image = $null }
        } else {
            $PicPreview.Image = $null
        }
        $PicPreview.Refresh()
    }
})

# Launch
Apply-Theme $Global:IsDarkMode
[System.Windows.Forms.Application]::EnableVisualStyles()
[Windows.Forms.Application]::Run($Form)
}
catch {
    $ErrorMsg = "`n[" + (Get-Date).ToString("yyyy-MM-dd HH:mm:ss") + "] SCRIPT ERROR:`n" + $_.Exception.Message + "`n" + $_.ScriptStackTrace + "`n" + ("=" * 50)
    
    # Try GUI error message
    try {
        if ([System.Type]::GetType("System.Windows.Forms.MessageBox")) {
            [System.Windows.Forms.MessageBox]::Show("CRITICAL ERROR DURING STARTUP:`n`n$($_.Exception.Message)`n`nDetails logged to: $Global:LogPath", "Macro Studio Pro Failure")
        }
    } catch { }

    # Write to local log file
    try {
        $ErrorMsg | Out-File $Global:LogPath -Append -Encoding UTF8
    } catch {
        # Fallback to temp directory if root folder is write-protected
        $fallback = Join-Path $env:TEMP "MacroStudio_Error.log"
        $ErrorMsg | Out-File $fallback -Append
    }
    
    Write-Error $_.Exception.Message
}
