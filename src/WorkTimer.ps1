#requires -Version 7.2

<#
.SYNOPSIS
    Presence-aware work timer with a Windows Forms GUI (PowerShell 7, Windows, no admin rights needed).

.DESCRIPTION
    Active work = (time since the workday started) minus the union of all inactive periods.
    Inactive periods are: automatic idle (no input for the configured threshold), manual
    entries, pauses, time the timer was not running, and time the computer was asleep.
    Every period can be added, edited, or deleted, and totals are recalculated exactly.

    Main window: End Workday, Pause Timer, Idle Periods, Work History.
    Everything else (edit start time, reopen/reset the day, start with Windows, idle
    threshold, data tools) is under Ctrl+Shift+A.

.NOTES
    Data folder: %APPDATA%\PowerShellPresenceWorkTimer
      Work History.xlsx   Excel work log, regenerated whenever history changes
      work-history.json   Internal copy of history (+ .bak); never locked by Excel
      state.json          Today's in-progress workday
      settings.json       User settings
      errors.log          Diagnostic log

    Launch without a console window (this is what the shortcuts created by the app use):
      conhost.exe --headless pwsh.exe -NoProfile -STA -WindowStyle Hidden -File "C:\Path\WorkTimer.ps1"
#>

[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

if (-not $IsWindows) { throw 'This script requires Windows.' }

# Windows Forms needs a single-threaded apartment. Relaunch hidden in STA if needed.
if ([Threading.Thread]::CurrentThread.GetApartmentState() -ne [Threading.ApartmentState]::STA) {
    if ([string]::IsNullOrWhiteSpace($PSCommandPath)) { throw 'Run this script with pwsh -STA.' }
    Start-Process -FilePath (Get-Process -Id $PID).Path -WindowStyle Hidden -ArgumentList @(
        '-NoProfile', '-STA', '-WindowStyle', 'Hidden', '-File', ('"{0}"' -f $PSCommandPath)
    )
    return
}

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Data
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName Microsoft.Win32.SystemEvents

if (-not ('PresenceWorkTimer.V4.Native' -as [type])) {
    Add-Type -ReferencedAssemblies @([Microsoft.Win32.SystemEvents].Assembly.Location) -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
using Microsoft.Win32;

namespace PresenceWorkTimer.V4
{
    public static class Native
    {
        [StructLayout(LayoutKind.Sequential)]
        private struct LASTINPUTINFO { public uint cbSize; public uint dwTime; }

        [DllImport("user32.dll")] private static extern bool GetLastInputInfo(ref LASTINPUTINFO info);
        [DllImport("kernel32.dll")] private static extern IntPtr GetConsoleWindow();
        [DllImport("kernel32.dll")] private static extern uint GetConsoleProcessList(uint[] list, uint count);
        [DllImport("user32.dll")] private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
        [DllImport("user32.dll")] private static extern bool AllowSetForegroundWindow(int processId);

        // Tick count (ms since boot, 32-bit) of the last keyboard/mouse input. 0 on failure.
        public static uint GetLastInputTick()
        {
            var info = new LASTINPUTINFO();
            info.cbSize = (uint)Marshal.SizeOf(typeof(LASTINPUTINFO));
            return GetLastInputInfo(ref info) ? info.dwTime : 0u;
        }

        // Milliseconds since the last input.
        public static long GetIdleMilliseconds()
        {
            var info = new LASTINPUTINFO();
            info.cbSize = (uint)Marshal.SizeOf(typeof(LASTINPUTINFO));
            if (!GetLastInputInfo(ref info)) return 0;
            return ComputeIdleMilliseconds(unchecked((uint)Environment.TickCount), info.dwTime);
        }

        // On some PCs Windows stamps the last input slightly AHEAD of the current tick count.
        // A plain unsigned subtraction turns that tiny negative into ~49.7 days of "idle" while the
        // user is typing. Reading the modular difference as signed handles both that case (-> 0)
        // and the normal 32-bit tick wrap-around.
        public static long ComputeIdleMilliseconds(uint nowTick, uint lastInputTick)
        {
            int difference = unchecked((int)(nowTick - lastInputTick));
            return difference < 0 ? 0 : difference;
        }

        public static string DescribeInputClock()
        {
            var info = new LASTINPUTINFO();
            info.cbSize = (uint)Marshal.SizeOf(typeof(LASTINPUTINFO));
            bool ok = GetLastInputInfo(ref info);
            return string.Format("tick={0} lastInput={1} ok={2}", unchecked((uint)Environment.TickCount), info.dwTime, ok);
        }

        // Hides the console only when it was created for this process alone
        // (never hides a terminal the user launched the script from).
        public static bool HideOwnConsole()
        {
            IntPtr handle = GetConsoleWindow();
            if (handle == IntPtr.Zero) return false;
            uint[] processes = new uint[4];
            if (GetConsoleProcessList(processes, (uint)processes.Length) > 1) return false;
            return ShowWindow(handle, 0);
        }

        public static void AllowAnyForeground() { AllowSetForegroundWindow(-1); }
    }

    // Taskbar identity. Without its own AppUserModelID the window is grouped under PowerShell
    // (pwsh.exe, a Store/MSIX app here) and shows PowerShell's icon. Giving the window and its
    // shortcuts the same ID makes the taskbar treat Work Timer as its own app with its own icon.
    [ComImport, Guid("886D8EEB-8CF2-4446-8D02-CDBA1DBDCF99"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IPropertyStore
    {
        [PreserveSig] int GetCount(out uint count);
        [PreserveSig] int GetAt(uint index, out PropertyKey key);
        [PreserveSig] int GetValue(ref PropertyKey key, IntPtr value);
        [PreserveSig] int SetValue(ref PropertyKey key, ref PropVariant value);
        [PreserveSig] int Commit();
    }

    [StructLayout(LayoutKind.Sequential, Pack = 4)]
    internal struct PropertyKey { public Guid FormatId; public uint PropertyId; }

    [StructLayout(LayoutKind.Explicit, Size = 24)]
    internal struct PropVariant { [FieldOffset(0)] public ushort VarType; [FieldOffset(8)] public IntPtr Pointer; }

    public static class AppIdentity
    {
        private static readonly Guid AppUserModelFormat = new Guid("9F4C2855-9F79-4B39-A8D0-E1D42DE1D5F3");
        private const uint RelaunchCommand = 2, RelaunchIconResource = 3, RelaunchDisplayName = 4, AppId = 5;

        [DllImport("shell32.dll", CharSet = CharSet.Unicode)]
        private static extern int SetCurrentProcessExplicitAppUserModelID(string appId);
        [DllImport("shell32.dll")]
        private static extern int SHGetPropertyStoreForWindow(IntPtr hwnd, ref Guid iid, [MarshalAs(UnmanagedType.Interface)] out IPropertyStore store);
        [DllImport("shell32.dll", CharSet = CharSet.Unicode)]
        private static extern int SHGetPropertyStoreFromParsingName(string path, IntPtr bindContext, int flags, ref Guid iid, [MarshalAs(UnmanagedType.Interface)] out IPropertyStore store);

        public static int SetProcessAppId(string appId) { return SetCurrentProcessExplicitAppUserModelID(appId); }

        // Relaunch* properties also make "Pin to taskbar" pin Work Timer (not PowerShell) with this icon.
        // Returns the names of optional relaunch properties Windows rejected ("" when all were accepted).
        public static string SetWindowAppId(IntPtr hwnd, string appId, string relaunchCommand, string iconPath, string displayName)
        {
            Guid iid = typeof(IPropertyStore).GUID;
            IPropertyStore store;
            Marshal.ThrowExceptionForHR(SHGetPropertyStoreForWindow(hwnd, ref iid, out store));
            string rejected = "";
            try
            {
                if (!string.IsNullOrEmpty(relaunchCommand) && !TrySetString(store, RelaunchCommand, relaunchCommand)) { rejected += "RelaunchCommand "; }
                if (!string.IsNullOrEmpty(iconPath) && !TrySetString(store, RelaunchIconResource, iconPath)) { rejected += "RelaunchIconResource "; }
                if (!TrySetString(store, RelaunchDisplayName, displayName)) { rejected += "RelaunchDisplayNameResource "; }
                SetString(store, AppId, appId);   // the essential one: gives the window its own taskbar group and icon
                Marshal.ThrowExceptionForHR(store.Commit());
            }
            finally { Marshal.ReleaseComObject(store); }
            return rejected.Trim();
        }

        private static bool TrySetString(IPropertyStore store, uint propertyId, string value)
        {
            try { SetString(store, propertyId, value); return true; } catch (COMException) { return false; }
        }

        public static void SetShortcutAppId(string shortcutPath, string appId)
        {
            Guid iid = typeof(IPropertyStore).GUID;
            IPropertyStore store;
            Marshal.ThrowExceptionForHR(SHGetPropertyStoreFromParsingName(shortcutPath, IntPtr.Zero, 2 /* GPS_READWRITE */, ref iid, out store));
            try
            {
                SetString(store, AppId, appId);
                Marshal.ThrowExceptionForHR(store.Commit());
            }
            finally { Marshal.ReleaseComObject(store); }
        }

        private static void SetString(IPropertyStore store, uint propertyId, string value)
        {
            var key = new PropertyKey { FormatId = AppUserModelFormat, PropertyId = propertyId };
            var variant = new PropVariant { VarType = 31 /* VT_LPWSTR */, Pointer = Marshal.StringToCoTaskMemUni(value) };
            try { Marshal.ThrowExceptionForHR(store.SetValue(ref key, ref variant)); }
            finally { Marshal.FreeCoTaskMem(variant.Pointer); }
        }
    }

    // Handled in C# because SystemEvents may raise events on a thread with no PowerShell runspace.
    public static class SessionWatcher
    {
        private static volatile bool pending;
        private static bool started;

        public static void Start()
        {
            if (started) return;
            SystemEvents.SessionSwitch += OnSessionSwitch;
            started = true;
        }

        public static void Stop()
        {
            if (!started) return;
            SystemEvents.SessionSwitch -= OnSessionSwitch;
            started = false;
        }

        public static bool Consume()
        {
            if (!pending) return false;
            pending = false;
            return true;
        }

        private static void OnSessionSwitch(object sender, SessionSwitchEventArgs e)
        {
            if (e.Reason == SessionSwitchReason.SessionUnlock ||
                e.Reason == SessionSwitchReason.SessionLogon ||
                e.Reason == SessionSwitchReason.RemoteConnect)
            {
                pending = true;
            }
        }
    }
}
'@
}

# ---------------------------------------------------------------------------
# MARK: Configuration and globals
# ---------------------------------------------------------------------------

$script:AppName = 'Work Timer'
$script:ScriptPath = if ([string]::IsNullOrWhiteSpace($PSCommandPath)) { $null } else { [IO.Path]::GetFullPath($PSCommandPath) }
$script:Culture = [Globalization.CultureInfo]::InvariantCulture
$script:Utf8 = [Text.UTF8Encoding]::new($false)
$script:AppDirectory = Join-Path ([Environment]::GetFolderPath('ApplicationData')) 'PowerShellPresenceWorkTimer'
$script:StatePath = Join-Path $script:AppDirectory 'state.json'
$script:LegacyStatePath = Join-Path $script:AppDirectory 'presence-work-timer-state.json'
$script:SettingsPath = Join-Path $script:AppDirectory 'settings.json'
$script:HistoryPath = Join-Path $script:AppDirectory 'work-history.json'
$script:WorkbookPath = Join-Path $script:AppDirectory 'Work History.xlsx'
$script:LegacyCsvPath = Join-Path $script:AppDirectory 'work-log.csv'
$script:ErrorLogPath = Join-Path $script:AppDirectory 'errors.log'
$script:ShortcutName = 'Work Timer.lnk'
$script:AppUserModelId = 'WorkTimer.PresenceAwareWorkTimer'   # Taskbar identity shared by the window and shortcuts.
$script:UserSid = [Security.Principal.WindowsIdentity]::GetCurrent().User.Value
$script:LegacyTaskName = "PowerShellPresenceWorkTimer-$($script:UserSid)"

$script:SleepGapSeconds = 90        # A gap between samples larger than this means the PC was asleep.
$script:RecoveryPromptSeconds = 120 # Downtime larger than this triggers the "count it?" prompt.
$script:SaveIntervalSeconds = 15
$script:ExportRetrySeconds = 30
$script:TenthHourSeconds = 360

# Inactive period types, in the order shown in the editor.
$script:PeriodTypes = [ordered]@{
    Manual  = 'Manual'
    Idle    = 'Idle (no input)'
    Paused  = 'Paused'
    Offline = 'Timer not running'
    Asleep  = 'Computer asleep'
    Stopped = 'After workday ended'
}

$script:State = $null
$script:Settings = $null
$script:History = $null
$script:MainForm = $null
$script:Ui = @{}
$script:Timer = $null
$script:Mutex = $null
$script:ActivateEvent = $null
$script:ShuttingDown = $false
$script:LastInputTick = [uint32]0
$script:IsPresent = $true
$script:IdleMilliseconds = [long]0
$script:LastSavedAt = [DateTime]::MinValue
$script:ExcelPending = $false
$script:LastExportAttempt = [DateTime]::MinValue
$script:LastTickError = $null
$script:LastTickErrorLoggedAt = [DateTime]::MinValue
$script:StartupWarnings = [Collections.Generic.List[string]]::new()
$script:ReviewPromptDone = $false   # The "fix the previous day" prompt was already shown this session.
$script:ReviewPromptDue = $false
$script:ReviewDialogOpen = $false
$script:HistoryNeedsRewrite = $false
$script:HistoryFileVersion = 3
$script:Scale = 1.0
$script:Colors = @{}
$script:Fonts = @{}
$script:AppIcon = $null

# ---------------------------------------------------------------------------
# MARK: General helpers
# ---------------------------------------------------------------------------

function Initialize-Storage {
    [void][IO.Directory]::CreateDirectory($script:AppDirectory)
}

function Write-AppLog {
    param(
        [Parameter(Mandatory)][string]$Message,
        [AllowNull()][Management.Automation.ErrorRecord]$ErrorRecord
    )
    try {
        Initialize-Storage
        if ((Test-Path -LiteralPath $script:ErrorLogPath -PathType Leaf) -and ([IO.FileInfo]::new($script:ErrorLogPath).Length -gt 1MB)) {
            Move-Item -LiteralPath $script:ErrorLogPath -Destination "$($script:ErrorLogPath).old" -Force
        }
        $line = '{0}  {1}' -f (Get-Date).ToString('yyyy-MM-dd HH:mm:ss', $script:Culture), $Message
        if ($null -ne $ErrorRecord) {
            $line += ' | {0}: {1}' -f $ErrorRecord.Exception.GetType().FullName, $ErrorRecord.Exception.Message
            if ($null -ne $ErrorRecord.InvocationInfo) {
                $line += ' | ' + ($ErrorRecord.InvocationInfo.PositionMessage -replace '\s+', ' ')
            }
        }
        [IO.File]::AppendAllText($script:ErrorLogPath, $line + [Environment]::NewLine, $script:Utf8)
    }
    catch { }
}

function Write-FileAtomic {
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][AllowEmptyString()][string]$Content,
        [switch]$KeepBackup
    )
    Initialize-Storage
    $temp = Join-Path (Split-Path -Parent $Path) ('.{0}.{1}.tmp' -f [IO.Path]::GetFileName($Path), [Guid]::NewGuid().ToString('N'))
    try {
        [IO.File]::WriteAllText($temp, $Content, $script:Utf8)
        if ($KeepBackup -and (Test-Path -LiteralPath $Path -PathType Leaf)) {
            [IO.File]::Replace($temp, $Path, "$Path.bak", $true)
        }
        else {
            [IO.File]::Move($temp, $Path, $true)
        }
    }
    finally {
        if (Test-Path -LiteralPath $temp) { Remove-Item -LiteralPath $temp -Force -ErrorAction SilentlyContinue }
    }
}

function Get-PropertyValue {
    param([AllowNull()][object]$Object, [Parameter(Mandatory)][string]$Name, [AllowNull()][object]$Default = $null)
    if ($null -ne $Object -and $null -ne $Object.PSObject.Properties[$Name]) { return $Object.$Name }
    return $Default
}

function New-List {
    Write-Output -NoEnumerate ([Collections.ArrayList]::new())
}

function New-Id {
    [Guid]::NewGuid().ToString('N')
}

function Get-DateKey {
    param([DateTime]$Date = (Get-Date))
    $Date.ToString('yyyy-MM-dd', $script:Culture)
}

function ConvertTo-IsoText {
    param([AllowNull()][object]$Value)
    if ($null -eq $Value) { return $null }
    ([DateTime]$Value).ToString('o', $script:Culture)
}

function ConvertFrom-IsoValue {
    # ConvertFrom-Json in PowerShell 7 may already return DateTime objects; accept either form.
    param([AllowNull()][object]$Value)
    if ($null -eq $Value) { return $null }
    if ($Value -is [DateTime]) { if ($Value.Kind -eq [DateTimeKind]::Utc) { return $Value.ToLocalTime() } return $Value }
    if ($Value -is [DateTimeOffset]) { return $Value.LocalDateTime }
    $text = [string]$Value
    if ([string]::IsNullOrWhiteSpace($text)) { return $null }
    $parsed = [DateTime]::MinValue
    if ([DateTime]::TryParse($text, $script:Culture, [Globalization.DateTimeStyles]::RoundtripKind, [ref]$parsed)) {
        if ($parsed.Kind -eq [DateTimeKind]::Utc) { return $parsed.ToLocalTime() }
        return $parsed
    }
    return $null
}

function Format-Time {
    param([AllowNull()][object]$Value, [string]$Format = 'HH:mm:ss')
    if ($null -eq $Value) { return '' }
    ([DateTime]$Value).ToString($Format, $script:Culture)
}

function Format-Duration {
    param([double]$Seconds)
    $span = [TimeSpan]::FromSeconds([math]::Max([double]0, [math]::Round($Seconds)))
    if ($span.TotalMinutes -lt 1) { return '0 min' }
    if ($span.TotalHours -lt 1) { return '{0} min' -f [int][math]::Floor($span.TotalMinutes) }
    '{0} h {1:00} min' -f [int][math]::Floor($span.TotalHours), $span.Minutes
}

function Get-RoundedHours {
    # Nearest 0.1 h, halves rounded up. Computed on seconds so it changes exactly when the countdown hits zero.
    param([double]$Seconds)
    $tenths = [math]::Floor(([math]::Max([double]0, $Seconds) + $script:TenthHourSeconds / 2) / $script:TenthHourSeconds)
    [double]$tenths / 10
}

function Get-SecondsToNextTenth {
    # The displayed value (rounded to 0.1 h, half away from zero) changes at 180 s + k * 360 s.
    param([double]$ActiveSeconds)
    $half = $script:TenthHourSeconds / 2
    $k = [math]::Floor(($ActiveSeconds - $half) / $script:TenthHourSeconds) + 1
    ($half + $k * $script:TenthHourSeconds) - $ActiveSeconds
}

# ---------------------------------------------------------------------------
# MARK: Settings
# ---------------------------------------------------------------------------

function Read-Settings {
    $script:Settings = [pscustomobject]@{ IdleThresholdMinutes = 480 }
    if (-not (Test-Path -LiteralPath $script:SettingsPath -PathType Leaf)) { return }
    try {
        $saved = Get-Content -LiteralPath $script:SettingsPath -Raw | ConvertFrom-Json
        $minutes = Get-PropertyValue $saved 'IdleThresholdMinutes'
        if ($null -ne $minutes -and [int]$minutes -ge 1 -and [int]$minutes -le 1440) {
            $script:Settings.IdleThresholdMinutes = [int]$minutes
        }
    }
    catch { Write-AppLog 'Could not read settings; using defaults.' $_ }
}

function Save-Settings {
    Write-FileAtomic -Path $script:SettingsPath -Content ($script:Settings | ConvertTo-Json)
}

# ---------------------------------------------------------------------------
# MARK: Single instance
# ---------------------------------------------------------------------------

function Initialize-SingleInstance {
    # One instance per Windows user (the state file is per user, so two copies would conflict).
    $bytes = [Text.Encoding]::UTF8.GetBytes("$($script:UserSid)|PresenceWorkTimer")
    $hash = [Convert]::ToHexString([Security.Cryptography.SHA256]::HashData($bytes)).Substring(0, 20)
    $mutexName = "Local\PresenceWorkTimer-$hash"
    $eventName = "Local\PresenceWorkTimer-$hash-Activate"

    [bool]$created = $false
    $script:Mutex = [Threading.Mutex]::new($true, $mutexName, [ref]$created)
    if (-not $created) {
        $script:Mutex.Dispose()
        $script:Mutex = $null
        try {
            $existing = [Threading.EventWaitHandle]::OpenExisting($eventName)
            try {
                [PresenceWorkTimer.V4.Native]::AllowAnyForeground()
                [void]$existing.Set()
            }
            finally { $existing.Dispose() }
        }
        catch {
            [void][Windows.Forms.MessageBox]::Show('Work Timer is already running.', $script:AppName, 'OK', 'Information')
        }
        return $false
    }

    $script:ActivateEvent = [Threading.EventWaitHandle]::new($false, [Threading.EventResetMode]::AutoReset, $eventName)
    return $true
}

function Close-SingleInstance {
    if ($null -ne $script:ActivateEvent) { $script:ActivateEvent.Dispose(); $script:ActivateEvent = $null }
    if ($null -eq $script:Mutex) { return }
    try { $script:Mutex.ReleaseMutex() } catch { }
    finally { $script:Mutex.Dispose(); $script:Mutex = $null }
}

# ---------------------------------------------------------------------------
# MARK: Workday state
# ---------------------------------------------------------------------------

function New-DayState {
    [pscustomobject]@{
        Version           = 3
        DateKey           = Get-DateKey
        RecordId          = $null
        TrackingStarted   = $false
        SessionStart      = $null
        StartReason       = $null
        Paused            = $false
        PauseStart        = $null
        Ended             = $false
        EndedAt           = $null
        EndReason         = $null
        LastSampleTime    = Get-Date
        CurrentIdleStart  = $null
        AdjustmentSeconds = [double]0   # Only non-zero for a day migrated from version 1.
        Periods           = New-List
    }
}

function New-Period {
    param([Parameter(Mandatory)][string]$Type, [Parameter(Mandatory)][DateTime]$Start, [Parameter(Mandatory)][DateTime]$End, [string]$Id)
    [pscustomobject]@{
        Id    = if ([string]::IsNullOrWhiteSpace($Id)) { New-Id } else { $Id }
        Type  = if ($script:PeriodTypes.Contains($Type)) { $Type } else { 'Manual' }
        Start = $Start
        End   = $End
    }
}

function Add-Period {
    param([Parameter(Mandatory)][string]$Type, [Parameter(Mandatory)][DateTime]$Start, [Parameter(Mandatory)][DateTime]$End)
    if ($End -le $Start) { return }
    [void]$script:State.Periods.Add((New-Period -Type $Type -Start $Start -End $End))
}

function Close-OpenIdle {
    param([Parameter(Mandatory)][DateTime]$End)
    $state = $script:State
    if ($null -eq $state.CurrentIdleStart) { return }
    $start = $state.CurrentIdleStart
    $state.CurrentIdleStart = $null
    Add-Period -Type Idle -Start $start -End $End
}

function Save-State {
    if ($null -eq $script:State) { return }
    $s = $script:State
    $saved = [ordered]@{
        Version           = 3
        DateKey           = $s.DateKey
        RecordId          = $s.RecordId
        TrackingStarted   = [bool]$s.TrackingStarted
        SessionStart      = ConvertTo-IsoText $s.SessionStart
        StartReason       = $s.StartReason
        Paused            = [bool]$s.Paused
        PauseStart        = ConvertTo-IsoText $s.PauseStart
        Ended             = [bool]$s.Ended
        EndedAt           = ConvertTo-IsoText $s.EndedAt
        EndReason         = $s.EndReason
        LastSampleTime    = ConvertTo-IsoText $s.LastSampleTime
        CurrentIdleStart  = ConvertTo-IsoText $s.CurrentIdleStart
        AdjustmentSeconds = [double]$s.AdjustmentSeconds
        Periods           = @(
            foreach ($p in $s.Periods) {
                [ordered]@{ Id = $p.Id; Type = $p.Type; Start = ConvertTo-IsoText $p.Start; End = ConvertTo-IsoText $p.End }
            }
        )
    }
    Write-FileAtomic -Path $script:StatePath -Content ($saved | ConvertTo-Json -Depth 5)
    $script:LastSavedAt = Get-Date
}

function Read-State {
    Initialize-Storage
    $path = if (Test-Path -LiteralPath $script:StatePath -PathType Leaf) { $script:StatePath }
            elseif (Test-Path -LiteralPath $script:LegacyStatePath -PathType Leaf) { $script:LegacyStatePath }
            else { $null }
    if ($null -eq $path) { $script:State = New-DayState; return }

    try {
        $saved = Get-Content -LiteralPath $path -Raw | ConvertFrom-Json
        $state = New-DayState
        $state.DateKey = [string](Get-PropertyValue $saved 'DateKey' $state.DateKey)
        $state.RecordId = [string](Get-PropertyValue $saved 'RecordId' '')
        $state.TrackingStarted = [bool](Get-PropertyValue $saved 'TrackingStarted' $false)
        $state.SessionStart = ConvertFrom-IsoValue (Get-PropertyValue $saved 'SessionStart')
        $state.StartReason = [string](Get-PropertyValue $saved 'StartReason' '')
        $state.Paused = [bool](Get-PropertyValue $saved 'Paused' $false)
        $state.PauseStart = ConvertFrom-IsoValue (Get-PropertyValue $saved 'PauseStart')
        $state.Ended = [bool](Get-PropertyValue $saved 'Ended' $false)
        $state.EndedAt = ConvertFrom-IsoValue (Get-PropertyValue $saved 'EndedAt')
        $state.EndReason = [string](Get-PropertyValue $saved 'EndReason' '')
        $state.CurrentIdleStart = ConvertFrom-IsoValue (Get-PropertyValue $saved 'CurrentIdleStart')
        $state.AdjustmentSeconds = [double](Get-PropertyValue $saved 'AdjustmentSeconds' 0)
        $lastSample = ConvertFrom-IsoValue (Get-PropertyValue $saved 'LastSampleTime')
        $state.LastSampleTime = if ($null -eq $lastSample) { Get-Date } else { $lastSample }

        if ([int](Get-PropertyValue $saved 'Version' 1) -lt 2) {
            Convert-LegacyState -Saved $saved -State $state
        }
        else {
            foreach ($p in @(Get-PropertyValue $saved 'Periods' @())) {
                $start = ConvertFrom-IsoValue (Get-PropertyValue $p 'Start')
                $end = ConvertFrom-IsoValue (Get-PropertyValue $p 'End')
                if ($null -eq $start -or $null -eq $end -or $end -le $start) { continue }
                [void]$state.Periods.Add((New-Period -Type ([string](Get-PropertyValue $p 'Type' 'Manual')) -Start $start -End $end -Id ([string](Get-PropertyValue $p 'Id' ''))))
            }
        }

        $savedVersion = [int](Get-PropertyValue $saved 'Version' 1)
        if ($savedVersion -lt 3) { Repair-FalseIdlePeriods -State $state }
        if ($state.TrackingStarted -and $null -eq $state.SessionStart) { $state.TrackingStarted = $false }
        if ($state.TrackingStarted -and [string]::IsNullOrWhiteSpace($state.RecordId)) { $state.RecordId = New-Id }
        if ($state.Paused -and $null -eq $state.PauseStart) { $state.PauseStart = $state.LastSampleTime }
        if ($state.Ended -and $null -eq $state.EndedAt) { $state.EndedAt = $state.LastSampleTime }
        $script:State = $state
    }
    catch {
        Write-AppLog "State file '$path' was unreadable; starting a new day." $_
        $backup = '{0}.invalid-{1}.json' -f $path, (Get-Date).ToString('yyyyMMdd-HHmmss', $script:Culture)
        Copy-Item -LiteralPath $path -Destination $backup -ErrorAction SilentlyContinue
        $script:StartupWarnings.Add("Today's saved timer state could not be read and was reset. A copy was kept at:`n$backup")
        $script:State = New-DayState
    }

    if ($path -eq $script:LegacyStatePath) {
        Save-State
        Move-Item -LiteralPath $script:LegacyStatePath -Destination "$($script:LegacyStatePath).migrated" -Force -ErrorAction SilentlyContinue
    }
}

function Repair-FalseIdlePeriods {
    # Versions before 3 could mis-read the input clock as "no input for ~49 days" while the user was
    # active, creating Idle periods clamped to start exactly at the workday start. Remove them once.
    param([Parameter(Mandatory)][object]$State)
    if ($null -eq $State.SessionStart) { return }
    $bogus = @($State.Periods | Where-Object { $_.Type -eq 'Idle' -and $_.Start -eq $State.SessionStart })
    foreach ($p in $bogus) { [void]$State.Periods.Remove($p) }
    $openBogus = $null -ne $State.CurrentIdleStart -and $State.CurrentIdleStart -eq $State.SessionStart
    if ($openBogus) { $State.CurrentIdleStart = $null }
    if ($bogus.Count -gt 0 -or $openBogus) {
        Write-AppLog ("Removed {0} false idle period(s) caused by the input-clock bug in an earlier version." -f ($bogus.Count + [int]$openBogus))
    }
}

function Convert-LegacyState {
    # Version 1 kept a running total. Rebuild periods and keep today's total exactly as it was.
    param([Parameter(Mandatory)][object]$Saved, [Parameter(Mandatory)][object]$State)
    foreach ($p in @(Get-PropertyValue $Saved 'IdlePeriods' @())) {
        $start = ConvertFrom-IsoValue (Get-PropertyValue $p 'Start')
        $end = ConvertFrom-IsoValue (Get-PropertyValue $p 'End')
        if ($null -eq $start -or $null -eq $end -or $end -le $start) { continue }
        $type = if ([string](Get-PropertyValue $p 'Source' '') -eq 'Manual') { 'Manual' } else { 'Idle' }
        [void]$State.Periods.Add((New-Period -Type $type -Start $start -End $end))
    }
    if ($State.Paused) { $State.PauseStart = $State.LastSampleTime }
    if ($State.TrackingStarted -and $null -ne $State.SessionStart) {
        $legacySeconds = [double](Get-PropertyValue $Saved 'AccumulatedActiveSeconds' 0)
        $computed = (Get-WorkSummary -State $State -At $State.LastSampleTime).ActiveSeconds
        $State.AdjustmentSeconds = $legacySeconds - $computed
    }
}

# ---------------------------------------------------------------------------
# MARK: Time calculation
# ---------------------------------------------------------------------------

function Measure-UnionSeconds {
    # Total length of the union of intervals, clipped to [RangeStart, RangeEnd]. Overlaps count once.
    param([object[]]$Intervals, [DateTime]$RangeStart, [DateTime]$RangeEnd)
    $clipped = [Collections.Generic.List[object]]::new()
    foreach ($i in $Intervals) {
        $s = if ($i.Start -lt $RangeStart) { $RangeStart } else { $i.Start }
        $e = if ($i.End -gt $RangeEnd) { $RangeEnd } else { $i.End }
        if ($e -gt $s) { $clipped.Add([pscustomobject]@{ Start = $s; End = $e }) }
    }
    if ($clipped.Count -eq 0) { return [double]0 }

    $total = [double]0
    $sorted = @($clipped | Sort-Object Start)
    $curStart = $sorted[0].Start
    $curEnd = $sorted[0].End
    for ($n = 1; $n -lt $sorted.Count; $n++) {
        $i = $sorted[$n]
        if ($i.Start -le $curEnd) {
            if ($i.End -gt $curEnd) { $curEnd = $i.End }
        }
        else {
            $total += ($curEnd - $curStart).TotalSeconds
            $curStart = $i.Start
            $curEnd = $i.End
        }
    }
    $total + ($curEnd - $curStart).TotalSeconds
}

function Get-WorkSummary {
    param([object]$State = $script:State, [AllowNull()][object]$At = $null)
    $summary = [pscustomobject]@{ ActiveSeconds = [double]0; InactiveSeconds = [double]0 }
    if (-not $State.TrackingStarted -or $null -eq $State.SessionStart) { return $summary }

    $start = [DateTime]$State.SessionStart
    $end = if ($State.Ended -and $null -ne $State.EndedAt) { [DateTime]$State.EndedAt }
           elseif ($null -ne $At) { [DateTime]$At }
           else { Get-Date }
    if ($end -le $start) { return $summary }

    $intervals = [Collections.Generic.List[object]]::new()
    foreach ($p in $State.Periods) { $intervals.Add($p) }
    if ($null -ne $State.CurrentIdleStart) { $intervals.Add([pscustomobject]@{ Start = $State.CurrentIdleStart; End = $end }) }
    if ($State.Paused -and $null -ne $State.PauseStart) { $intervals.Add([pscustomobject]@{ Start = $State.PauseStart; End = $end }) }

    $inactive = Measure-UnionSeconds -Intervals $intervals.ToArray() -RangeStart $start -RangeEnd $end
    $summary.InactiveSeconds = $inactive
    $summary.ActiveSeconds = [math]::Max([double]0, ($end - $start).TotalSeconds - $inactive + [double]$State.AdjustmentSeconds)
    $summary
}

function Get-PeriodSummaryText {
    # Describes the periods that overlap [RangeStart, RangeEnd], clipped to that range.
    param([AllowNull()][object[]]$Periods, [DateTime]$RangeStart, [DateTime]$RangeEnd)
    $items = foreach ($p in @($Periods | Where-Object { $null -ne $_ } | Sort-Object Start)) {
        $s = if ($p.Start -lt $RangeStart) { $RangeStart } else { $p.Start }
        $e = if ($p.End -gt $RangeEnd) { $RangeEnd } else { $p.End }
        if ($e -le $s) { continue }
        '{0}-{1} {2}' -f (Format-Time $s 'HH:mm'), (Format-Time $e 'HH:mm'), $script:PeriodTypes[[string]$p.Type]
    }
    @($items) -join '; '
}

function Copy-PeriodList {
    param([AllowNull()][object[]]$Periods)
    $list = New-List
    foreach ($p in @($Periods | Where-Object { $null -ne $_ })) {
        [void]$list.Add([pscustomobject]@{ Type = [string]$p.Type; Start = [DateTime]$p.Start; End = [DateTime]$p.End })
    }
    Write-Output -NoEnumerate $list
}

# ---------------------------------------------------------------------------
# MARK: Tracking
# ---------------------------------------------------------------------------

function Start-Tracking {
    param([Parameter(Mandatory)][string]$Reason, [AllowNull()][object]$At = $null)
    $s = $script:State
    if ($s.TrackingStarted -or $s.Ended) { return }
    $now = Get-Date
    $s.TrackingStarted = $true
    $s.SessionStart = if ($null -ne $At) { [DateTime]$At } else { $now }
    $s.StartReason = $Reason
    $s.RecordId = New-Id
    $s.LastSampleTime = $now
    Save-State
}

function Update-Sample {
    $s = $script:State
    $now = Get-Date
    $idleMs = [PresenceWorkTimer.V4.Native]::GetIdleMilliseconds()
    $present = $idleMs -lt ([double]$script:Settings.IdleThresholdMinutes * 60000)
    $script:IsPresent = $present
    $script:IdleMilliseconds = $idleMs

    if ($s.TrackingStarted -and -not $s.Ended -and -not $s.Paused) {
        if (($now - $s.LastSampleTime).TotalSeconds -gt $script:SleepGapSeconds) {
            # The process was suspended (sleep/hibernate): exclude the gap; editable later.
            Close-OpenIdle -End $s.LastSampleTime
            Add-Period -Type Asleep -Start $s.LastSampleTime -End $now
        }
        $lastInput = $now.AddMilliseconds(-$idleMs)
        if (-not $present -and $null -eq $s.CurrentIdleStart) {
            # Backdate to the actual last input, not the moment the threshold was crossed.
            $s.CurrentIdleStart = if ($lastInput -lt $s.SessionStart) { $s.SessionStart } else { $lastInput }
            # Record why, so an unexpected idle period can be diagnosed later (View Error Log).
            $session = (Get-Process -Id $PID).SessionId
            Write-AppLog ("Idle started: no keyboard/mouse input for {0} (threshold {1} min); idle counted from {2}. {3} session={4} remote={5}" -f
                (Format-Duration ($idleMs / 1000)), $script:Settings.IdleThresholdMinutes, (Format-Time $s.CurrentIdleStart),
                [PresenceWorkTimer.V4.Native]::DescribeInputClock(), $session, [Windows.Forms.SystemInformation]::TerminalServerSession)
        }
        elseif ($present -and $null -ne $s.CurrentIdleStart) {
            Close-OpenIdle -End $lastInput
        }
    }
    elseif ($null -ne $s.CurrentIdleStart) {
        Close-OpenIdle -End $now
    }
    $s.LastSampleTime = $now
}

function Switch-Pause {
    $s = $script:State
    if (-not $s.TrackingStarted -or $s.Ended) { return }
    Update-Sample
    $now = Get-Date
    if (-not $s.Paused) {
        Close-OpenIdle -End $now
        $s.Paused = $true
        $s.PauseStart = $now
    }
    else {
        if ($null -ne $s.PauseStart) { Add-Period -Type Paused -Start $s.PauseStart -End $now }
        $s.Paused = $false
        $s.PauseStart = $null
        $s.LastSampleTime = $now
    }
    Save-State
    Update-Display
}

function Complete-Workday {
    param([Parameter(Mandatory)][string]$Reason, [Parameter(Mandatory)][DateTime]$At, [switch]$NeedsReview)
    $s = $script:State
    if (-not $s.TrackingStarted -or $s.Ended) { return }
    if ($At -lt $s.SessionStart) { $At = $s.SessionStart }
    if ($null -ne $s.CurrentIdleStart) {
        $idleEnd = if ($s.CurrentIdleStart -gt $At) { $s.CurrentIdleStart } else { $At }
        Close-OpenIdle -End $idleEnd
    }
    if ($s.Paused -and $null -ne $s.PauseStart) { Add-Period -Type Paused -Start $s.PauseStart -End $At }
    $s.Paused = $false
    $s.PauseStart = $null
    $s.Ended = $true
    $s.EndedAt = $At
    $s.EndReason = $Reason
    Save-State
    Save-HistoryRecord -NeedsReview:$NeedsReview
}

function Invoke-DayRollover {
    param([string]$Reason = 'Automatic midnight rollover')
    $s = $script:State
    if ($s.DateKey -eq (Get-DateKey)) { return $false }
    if ($s.TrackingStarted -and -not $s.Ended) {
        $dayEnd = [DateTime]::ParseExact($s.DateKey, 'yyyy-MM-dd', $script:Culture).AddDays(1).AddSeconds(-1)
        $at = if ($s.LastSampleTime -lt $dayEnd) { $s.LastSampleTime } else { $dayEnd }
        # The workday was never ended, so the real end time is unknown: ask about it later.
        Complete-Workday -Reason $Reason -At $at -NeedsReview
        $script:ReviewPromptDone = $false
    }
    $script:State = New-DayState
    $script:LastInputTick = [PresenceWorkTimer.V4.Native]::GetLastInputTick()
    Save-State
    return $true
}

function Get-DowntimeGap {
    # Returns the period the app was not running during today's open workday, or $null.
    $s = $script:State
    if ($s.DateKey -ne (Get-DateKey) -or -not $s.TrackingStarted -or $s.Ended -or $s.Paused) { return $null }
    $now = Get-Date
    if (($now - $s.LastSampleTime).TotalSeconds -lt $script:RecoveryPromptSeconds) { return $null }
    [pscustomobject]@{ From = $s.LastSampleTime; To = $now }
}

function Resolve-DowntimeGap {
    param([Parameter(Mandatory)][object]$Gap)
    $s = $script:State
    Close-OpenIdle -End $Gap.From
    $message = ("Work Timer was not running from {0} to {1} ({2}). Today's workday started at {3}.`n`n" +
        "Count as Work: today's hours are calculated from the start of your workday, including that time.`n" +
        "Exclude: the gap is recorded as an inactive period. You can change this later in Idle Periods.") -f
        (Format-Time $Gap.From 'HH:mm'), (Format-Time $Gap.To 'HH:mm'), (Format-Duration ($Gap.To - $Gap.From).TotalSeconds), (Format-Time $s.SessionStart 'HH:mm')
    $count = Show-Choice -Title 'Resume Workday' -Heading 'Count the time Work Timer was closed?' -Message $message `
        -PrimaryText 'Count as Work' -SecondaryText 'Exclude'
    $now = Get-Date
    if (-not $count) { Add-Period -Type Offline -Start $Gap.From -End $now }
    $s.LastSampleTime = $now
    Save-State
}

# ---------------------------------------------------------------------------
# MARK: Work history (JSON source of truth + generated Excel workbook)
# ---------------------------------------------------------------------------

function ConvertTo-HistoryRecord {
    param([Parameter(Mandatory)][object]$Source)
    # Periods is $null for records imported from the old CSV log (their idle periods are unknown).
    $periods = $null
    $periodsProperty = $Source.PSObject.Properties['Periods']
    if ($null -ne $periodsProperty -and $null -ne $periodsProperty.Value) {
        $periods = New-List
        foreach ($p in @($periodsProperty.Value)) {
            $start = ConvertFrom-IsoValue (Get-PropertyValue $p 'Start')
            $end = ConvertFrom-IsoValue (Get-PropertyValue $p 'End')
            if ($null -eq $start -or $null -eq $end -or $end -le $start) { continue }
            [void]$periods.Add([pscustomobject]@{ Type = [string](Get-PropertyValue $p 'Type' 'Manual'); Start = $start; End = $end })
        }
    }
    $sessionStart = ConvertFrom-IsoValue (Get-PropertyValue $Source 'SessionStart')
    $rawDate = Get-PropertyValue $Source 'Date' ''
    $dateKey = ConvertTo-DateKeyText -Value $rawDate -Fallback $sessionStart
    if ($dateKey -ne [string]$rawDate) { $script:HistoryNeedsRewrite = $true }
    [pscustomobject]@{
        Id                = [string](Get-PropertyValue $Source 'Id' (New-Id))
        Date              = $dateKey
        SessionStart      = $sessionStart
        SessionEnd        = ConvertFrom-IsoValue (Get-PropertyValue $Source 'SessionEnd')
        ActiveSeconds     = [double](Get-PropertyValue $Source 'ActiveSeconds' 0)
        ActiveHours       = [double](Get-PropertyValue $Source 'ActiveHours' 0)
        InactiveSeconds   = [double](Get-PropertyValue $Source 'InactiveSeconds' 0)
        InactivePeriods   = [string](Get-PropertyValue $Source 'InactivePeriods' '')
        Reason            = [string](Get-PropertyValue $Source 'Reason' '')
        Periods           = $periods
        AdjustmentSeconds = [double](Get-PropertyValue $Source 'AdjustmentSeconds' 0)
        NeedsReview       = [bool](Get-PropertyValue $Source 'NeedsReview' $false)
    }
}

function Get-HistorySortKey {
    param([Parameter(Mandatory)][object]$Record)
    $start = if ($null -eq $Record.SessionStart) { '' } else { ([DateTime]$Record.SessionStart).ToString('HH:mm:ss', $script:Culture) }
    '{0}|{1}' -f $Record.Date, $start
}

function Sort-History {
    # Keeps history in date order (then by start time) so the JSON and the workbook are always sorted.
    $sorted = New-List
    foreach ($r in @($script:History | Sort-Object -Property @{ Expression = { Get-HistorySortKey $_ } })) { [void]$sorted.Add($r) }
    $script:History = $sorted
}

function Read-HistoryFile {
    param([Parameter(Mandatory)][string]$Path)
    $list = New-List
    $data = Get-Content -LiteralPath $Path -Raw | ConvertFrom-Json
    $script:HistoryFileVersion = [int](Get-PropertyValue $data 'Version' 1)
    foreach ($r in @(Get-PropertyValue $data 'Records' @())) { [void]$list.Add((ConvertTo-HistoryRecord $r)) }
    Write-Output -NoEnumerate $list
}

function Read-History {
    Initialize-Storage
    if (Test-Path -LiteralPath $script:HistoryPath -PathType Leaf) {
        try {
            $script:HistoryNeedsRewrite = $false
            $script:History = Read-HistoryFile $script:HistoryPath
            if ($script:HistoryFileVersion -lt 3 -and (Repair-HistoryFalseIdle)) { $script:HistoryNeedsRewrite = $true }
            if ($script:HistoryNeedsRewrite) {
                # Dates in a non-standard format were repaired: save them and rebuild the workbook.
                Save-History
                Export-HistoryWorkbook
            }
            return
        }
        catch { Write-AppLog 'Work history file was unreadable.' $_ }

        $stamp = (Get-Date).ToString('yyyyMMdd-HHmmss', $script:Culture)
        Copy-Item -LiteralPath $script:HistoryPath -Destination "$($script:HistoryPath).invalid-$stamp" -ErrorAction SilentlyContinue
        try {
            $script:History = Read-HistoryFile "$($script:HistoryPath).bak"
            $script:StartupWarnings.Add('The work history file was damaged and has been restored from its automatic backup.')
            Save-History
            return
        }
        catch {
            Write-AppLog 'Work history backup was unreadable.' $_
            $script:StartupWarnings.Add("The work history file could not be read. A copy was kept as work-history.json.invalid-$stamp in the data folder.")
            $script:History = New-List
            return
        }
    }

    $script:History = New-List
    if (Test-Path -LiteralPath $script:LegacyCsvPath -PathType Leaf) {
        Import-LegacyCsv
        Save-History
        $target = "$($script:LegacyCsvPath).migrated"
        if (Test-Path -LiteralPath $target) { $target = "$($script:LegacyCsvPath).migrated-$((Get-Date).ToString('yyyyMMddHHmmss', $script:Culture))" }
        try { Move-Item -LiteralPath $script:LegacyCsvPath -Destination $target } catch { Write-AppLog 'Could not rename the old CSV log after import.' $_ }
        Export-HistoryWorkbook
    }
}

function Import-LegacyCsv {
    $stream = [IO.FileStream]::new($script:LegacyCsvPath, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::ReadWrite)
    try { $reader = [IO.StreamReader]::new($stream, $script:Utf8, $true); $text = $reader.ReadToEnd(); $reader.Dispose() }
    finally { $stream.Dispose() }
    if ([string]::IsNullOrWhiteSpace($text)) { return }

    foreach ($row in @($text | ConvertFrom-Csv)) {
        $hours = ConvertFrom-LegacyNumber (Get-PropertyValue $row 'ActiveHours')
        $idleHours = ConvertFrom-LegacyNumber (Get-PropertyValue $row 'IdleHours')
        $sessionStart = ConvertFrom-LegacyDate (Get-PropertyValue $row 'SessionStart')
        [void]$script:History.Add([pscustomobject]@{
            Id              = New-Id
            Date            = ConvertTo-DateKeyText -Value (Get-PropertyValue $row 'Date' '') -Fallback $sessionStart
            SessionStart    = $sessionStart
            SessionEnd      = ConvertFrom-LegacyDate (Get-PropertyValue $row 'CompletedAt')
            ActiveSeconds   = $hours * 3600
            ActiveHours     = $hours
            InactiveSeconds = $idleHours * 3600
            InactivePeriods = [string](Get-PropertyValue $row 'IdlePeriods' '')
            Reason          = [string](Get-PropertyValue $row 'CompletionReason' '')
            Periods         = $null
            AdjustmentSeconds = [double]0
            NeedsReview     = $false
        })
    }
}

function ConvertTo-DateKeyText {
    # Normalizes a work date to yyyy-MM-dd. Accepts ISO dates, dates re-saved by Excel
    # (e.g. 10/1/2026), and DateTime values; falls back to the session start's date.
    param([AllowNull()][object]$Value, [AllowNull()][object]$Fallback)
    if ($Value -is [DateTime]) { return Get-DateKey $Value }
    $text = ([string]$Value).Trim()
    $parsed = [DateTime]::MinValue
    $none = [Globalization.DateTimeStyles]::None
    $formats = [string[]]@('yyyy-MM-dd', 'yyyy-M-d', 'yyyy/MM/dd', 'yyyy-MM-dd HH:mm:ss', 'yyyy-MM-ddTHH:mm:ss')
    if ($text -and [DateTime]::TryParseExact($text, $formats, $script:Culture, $none, [ref]$parsed)) { return Get-DateKey $parsed }
    if ($null -ne $Fallback) {
        # When the start time is known it decides the day, which resolves month/day order (10/1 vs 1/10).
        return Get-DateKey ([DateTime]$Fallback)
    }
    $current = [Globalization.CultureInfo]::CurrentCulture
    if ($text -and [DateTime]::TryParse($text, $current, $none, [ref]$parsed)) { return Get-DateKey $parsed }
    if ($text -and [DateTime]::TryParseExact($text, [string[]]@('M/d/yyyy', 'M/d/yyyy H:mm', 'M/d/yyyy h:mm tt'), $script:Culture, $none, [ref]$parsed)) { return Get-DateKey $parsed }
    $text
}

function ConvertFrom-LegacyDate {
    param([AllowNull()][object]$Value)
    $text = [string]$Value
    $parsed = [DateTime]::MinValue
    if ([DateTime]::TryParseExact($text, 'yyyy-MM-dd HH:mm:ss', $script:Culture, [Globalization.DateTimeStyles]::None, [ref]$parsed)) { return $parsed }
    if ([DateTime]::TryParse($text, [ref]$parsed)) { return $parsed }
    return $null
}

function ConvertFrom-LegacyNumber {
    # Version 1 wrote numbers with the current culture ('{0:N1}'), so try that first.
    param([AllowNull()][object]$Value)
    $text = [string]$Value
    $parsed = [double]0
    if ([double]::TryParse($text, [Globalization.NumberStyles]::Number, [Globalization.CultureInfo]::CurrentCulture, [ref]$parsed)) { return $parsed }
    if ([double]::TryParse($text, [Globalization.NumberStyles]::Number, $script:Culture, [ref]$parsed)) { return $parsed }
    return [double]0
}

function Save-History {
    Sort-History
    $payload = [ordered]@{
        Version = 3
        Records = @(
            foreach ($r in $script:History) {
                $periods = $null   # $null = unknown (imported from the old CSV); @() = no inactive periods
                if ($null -ne $r.Periods) {
                    $periods = @(foreach ($p in $r.Periods) { [ordered]@{ Type = $p.Type; Start = ConvertTo-IsoText $p.Start; End = ConvertTo-IsoText $p.End } })
                }
                [ordered]@{
                    Id                = $r.Id
                    Date              = $r.Date
                    SessionStart      = ConvertTo-IsoText $r.SessionStart
                    SessionEnd        = ConvertTo-IsoText $r.SessionEnd
                    ActiveSeconds     = [math]::Round([double]$r.ActiveSeconds, 1)
                    ActiveHours       = [double]$r.ActiveHours
                    InactiveSeconds   = [math]::Round([double]$r.InactiveSeconds, 1)
                    InactivePeriods   = $r.InactivePeriods
                    Reason            = $r.Reason
                    AdjustmentSeconds = [double]$r.AdjustmentSeconds
                    NeedsReview       = [bool]$r.NeedsReview
                    Periods           = $periods
                }
            }
        )
    }
    Write-FileAtomic -Path $script:HistoryPath -Content ($payload | ConvertTo-Json -Depth 6) -KeepBackup
}

function Find-HistoryRecord {
    param([Parameter(Mandatory)][AllowEmptyString()][string]$Id)
    foreach ($r in $script:History) { if ($r.Id -eq $Id) { return $r } }
    return $null
}

function Save-HistoryRecord {
    # Inserts or updates today's record (keyed by RecordId), then refreshes the workbook.
    param([switch]$NeedsReview)
    $s = $script:State
    if (-not $s.TrackingStarted -or -not $s.Ended) { return }
    $summary = Get-WorkSummary -State $s
    $existing = Find-HistoryRecord $s.RecordId
    $record = [pscustomobject]@{
        Id                = $s.RecordId
        Date              = $s.DateKey
        SessionStart      = $s.SessionStart
        SessionEnd        = $s.EndedAt
        ActiveSeconds     = $summary.ActiveSeconds
        ActiveHours       = Get-RoundedHours $summary.ActiveSeconds
        InactiveSeconds   = $summary.InactiveSeconds
        InactivePeriods   = Get-PeriodSummaryText -Periods @($s.Periods) -RangeStart $s.SessionStart -RangeEnd $s.EndedAt
        Reason            = $s.EndReason
        Periods           = Copy-PeriodList @($s.Periods)
        AdjustmentSeconds = [double]$s.AdjustmentSeconds
        NeedsReview       = $NeedsReview.IsPresent -or ($null -ne $existing -and [bool]$existing.NeedsReview)
    }
    if ($null -ne $existing) { $script:History[$script:History.IndexOf($existing)] = $record } else { [void]$script:History.Add($record) }
    Save-History
    Export-HistoryWorkbook
}

# --- Correcting a past day (e.g. the timer rolled over at midnight because nobody ended it) ---

function Repair-HistoryFalseIdle {
    # One-time cleanup of saved days affected by the input-clock bug (see Repair-FalseIdlePeriods).
    $fixed = 0
    foreach ($r in $script:History) {
        if ($null -eq $r.Periods -or $null -eq $r.SessionStart -or $null -eq $r.SessionEnd) { continue }
        $bogus = @($r.Periods | Where-Object { $_.Type -eq 'Idle' -and $_.Start -eq $r.SessionStart })
        if ($bogus.Count -eq 0) { continue }
        foreach ($p in $bogus) { [void]$r.Periods.Remove($p) }
        $active = Measure-RecordActiveSeconds -Record $r -Start $r.SessionStart -End $r.SessionEnd
        $r.ActiveSeconds = $active
        $r.ActiveHours = Get-RoundedHours $active
        $r.InactiveSeconds = [math]::Max([double]0, ($r.SessionEnd - $r.SessionStart).TotalSeconds - $active)
        $r.InactivePeriods = Get-PeriodSummaryText -Periods @($r.Periods) -RangeStart $r.SessionStart -RangeEnd $r.SessionEnd
        $fixed++
        Write-AppLog ('Recalculated {0} after removing false idle time: now {1} h.' -f $r.Date, $r.ActiveHours)
    }
    return $fixed -gt 0
}

function Get-RecordInactiveLump {
    # Records imported from the old CSV have no period list: treat their inactive time as one fixed amount.
    param([Parameter(Mandatory)][object]$Record)
    if ($null -ne $Record.SessionStart -and $null -ne $Record.SessionEnd) {
        return [math]::Max([double]0, ($Record.SessionEnd - $Record.SessionStart).TotalSeconds - [double]$Record.ActiveSeconds)
    }
    [double]$Record.InactiveSeconds
}

function Measure-RecordActiveSeconds {
    param([Parameter(Mandatory)][object]$Record, [Parameter(Mandatory)][DateTime]$Start, [Parameter(Mandatory)][DateTime]$End)
    if ($End -le $Start) { return [double]0 }
    $span = ($End - $Start).TotalSeconds
    if ($null -eq $Record.Periods) { return [math]::Max([double]0, $span - (Get-RecordInactiveLump $Record)) }
    $inactive = Measure-UnionSeconds -Intervals @($Record.Periods) -RangeStart $Start -RangeEnd $End
    [math]::Max([double]0, $span - $inactive + [double]$Record.AdjustmentSeconds)
}

function Get-EndTimeForActiveSeconds {
    # The time at which active work (start minus inactive periods) reaches TargetSeconds.
    param([Parameter(Mandatory)][object]$Record, [Parameter(Mandatory)][DateTime]$Start, [Parameter(Mandatory)][double]$TargetSeconds)
    if ($null -eq $Record.Periods) { return $Start.AddSeconds($TargetSeconds + (Get-RecordInactiveLump $Record)) }
    $remaining = $TargetSeconds - [double]$Record.AdjustmentSeconds
    $cursor = $Start
    foreach ($p in @($Record.Periods | Sort-Object Start)) {
        if ($p.End -le $cursor) { continue }
        $gapStart = if ($p.Start -gt $cursor) { $p.Start } else { $cursor }
        $workable = ($gapStart - $cursor).TotalSeconds
        if ($workable -ge $remaining) { return $cursor.AddSeconds([math]::Max([double]0, $remaining)) }
        $remaining -= $workable
        $cursor = $p.End
    }
    $cursor.AddSeconds([math]::Max([double]0, $remaining))
}

function Set-HistoryRecordTimes {
    param(
        [Parameter(Mandatory)][object]$Record, [Parameter(Mandatory)][DateTime]$Start,
        [Parameter(Mandatory)][DateTime]$End, [Parameter(Mandatory)][string]$Reason
    )
    $active = Measure-RecordActiveSeconds -Record $Record -Start $Start -End $End
    $Record.SessionStart = $Start
    $Record.SessionEnd = $End
    $Record.ActiveSeconds = $active
    $Record.ActiveHours = Get-RoundedHours $active
    $Record.InactiveSeconds = [math]::Max([double]0, ($End - $Start).TotalSeconds - $active)
    if ($null -ne $Record.Periods) {
        $Record.InactivePeriods = Get-PeriodSummaryText -Periods @($Record.Periods) -RangeStart $Start -RangeEnd $End
    }
    $Record.Reason = $Reason
    $Record.NeedsReview = $false

    # Keep today's live state in step if this is today's (already ended) workday.
    $s = $script:State
    if ($s.Ended -and $s.RecordId -eq $Record.Id) {
        $s.SessionStart = $Start
        $s.EndedAt = $End
        $s.EndReason = $Reason
        Save-State
    }
    Save-History
    Export-HistoryWorkbook
}

function Get-PendingReviews {
    @($script:History | Where-Object { $_.NeedsReview } | Sort-Object -Property @{ Expression = { Get-HistorySortKey $_ } })
}

function ConvertTo-XmlText {
    param([AllowNull()][object]$Value)
    $text = if ($null -eq $Value) { '' } else { [string]$Value }
    $text = [regex]::Replace($text, '[\x00-\x08\x0B\x0C\x0E-\x1F]', '')
    [Security.SecurityElement]::Escape($text)
}

function Format-XlsxTextCell {
    # Inline strings are never evaluated as formulas, which prevents spreadsheet formula injection.
    param([string]$Ref, [AllowNull()][object]$Text, [int]$Style = 0)
    if ($null -eq $Text -or [string]$Text -eq '') { return "<c r=`"$Ref`" s=`"$Style`"/>" }
    "<c r=`"$Ref`" s=`"$Style`" t=`"inlineStr`"><is><t xml:space=`"preserve`">$(ConvertTo-XmlText $Text)</t></is></c>"
}

function Format-XlsxNumberCell {
    param([string]$Ref, [AllowNull()][object]$Number, [int]$Style = 0)
    if ($null -eq $Number) { return "<c r=`"$Ref`" s=`"$Style`"/>" }
    "<c r=`"$Ref`" s=`"$Style`"><v>$(([double]$Number).ToString('R', $script:Culture))</v></c>"
}

function Write-ZipPackage {
    param([Parameter(Mandatory)][string]$Path, [Parameter(Mandatory)][Collections.Specialized.OrderedDictionary]$Parts)
    $stream = [IO.File]::Open($Path, [IO.FileMode]::CreateNew, [IO.FileAccess]::ReadWrite, [IO.FileShare]::None)
    try {
        $zip = [IO.Compression.ZipArchive]::new($stream, [IO.Compression.ZipArchiveMode]::Create, $true)
        try {
            foreach ($name in $Parts.Keys) {
                $entryStream = $zip.CreateEntry([string]$name, [IO.Compression.CompressionLevel]::Optimal).Open()
                try {
                    $bytes = $script:Utf8.GetBytes([string]$Parts[$name])
                    $entryStream.Write($bytes, 0, $bytes.Length)
                }
                finally { $entryStream.Dispose() }
            }
        }
        finally { $zip.Dispose() }
    }
    finally { $stream.Dispose() }
}

function Format-XlsxFormulaCell {
    # Formula cell with a cached result so the value shows even before Excel recalculates.
    param([string]$Ref, [string]$Formula, [AllowNull()][object]$Cached, [int]$Style = 0, [switch]$TextResult)
    $f = "<f>$(ConvertTo-XmlText $Formula)</f>"
    if ($TextResult) {
        $v = if ($null -eq $Cached) { '' } else { "<v>$(ConvertTo-XmlText $Cached)</v>" }
        return "<c r=`"$Ref`" s=`"$Style`" t=`"str`">$f$v</c>"
    }
    $v = if ($null -eq $Cached) { '' } else { "<v>$(([double]$Cached).ToString('R', $script:Culture))</v>" }
    "<c r=`"$Ref`" s=`"$Style`">$f$v</c>"
}

function Export-HistoryWorkbook {
    # Builds the .xlsx (Office Open XML) directly - no Excel or extra module needed to write it.
    # The data is an Excel table named WorkHistory, sorted by date, with formula columns:
    #   Day          = TEXT(Date, "ddd")
    #   Hours Worked = ROUND((End - Start) * 24 - Inactive Hours, 1)
    # If the workbook is open in Excel, the update is retried automatically once it is closed.
    $script:LastExportAttempt = Get-Date
    Sort-History
    $records = @($script:History)
    $table = 'WorkHistory'
    $columns = @(
        @{ Header = 'Date'; Width = 12 }, @{ Header = 'Day'; Width = 8 }, @{ Header = 'Start'; Width = 10 },
        @{ Header = 'End'; Width = 10 }, @{ Header = 'Inactive Hours'; Width = 16 }, @{ Header = 'Hours Worked'; Width = 15 },
        @{ Header = 'Inactive Periods'; Width = 70 }, @{ Header = 'Completion'; Width = 40 }
    )
    $dayFormula = "IF($table[[#This Row],[Date]]=`"`",`"`",TEXT($table[[#This Row],[Date]],`"ddd`"))"
    $hoursFormula = "IF(OR($table[[#This Row],[Start]]=`"`",$table[[#This Row],[End]]=`"`"),`"`"," +
        "ROUND(($table[[#This Row],[End]]-$table[[#This Row],[Start]])*24-$table[[#This Row],[Inactive Hours]],1))"

    $dataRows = [math]::Max(1, $records.Count)   # an Excel table needs at least one data row
    $lastDataRow = $dataRows + 1
    $totalRow = $lastDataRow + 1
    $ns = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'

    $sb = [Text.StringBuilder]::new()
    [void]$sb.Append("<?xml version=`"1.0`" encoding=`"UTF-8`" standalone=`"yes`"?><worksheet $ns>")
    [void]$sb.Append("<dimension ref=`"A1:H$totalRow`"/>")
    [void]$sb.Append('<sheetViews><sheetView tabSelected="1" workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>')
    [void]$sb.Append('<sheetFormatPr defaultRowHeight="15"/><cols>')
    for ($c = 0; $c -lt $columns.Count; $c++) {
        [void]$sb.Append(('<col min="{0}" max="{0}" width="{1}" customWidth="1"/>' -f ($c + 1), $columns[$c].Width))
    }
    [void]$sb.Append('</cols><sheetData><row r="1" ht="20" customHeight="1">')
    $letters = @('A', 'B', 'C', 'D', 'E', 'F', 'G', 'H')
    for ($c = 0; $c -lt $columns.Count; $c++) { [void]$sb.Append((Format-XlsxTextCell "$($letters[$c])1" $columns[$c].Header 0)) }
    [void]$sb.Append('</row>')

    $row = 1
    foreach ($r in $records) {
        $row++
        $date = [DateTime]::MinValue
        $hasDate = [DateTime]::TryParseExact($r.Date, 'yyyy-MM-dd', $script:Culture, [Globalization.DateTimeStyles]::None, [ref]$date)
        $hasTimes = $null -ne $r.SessionStart -and $null -ne $r.SessionEnd
        # Inactive Hours is written so that the Hours Worked formula reproduces the timer's own total.
        $inactiveHours = if ($hasTimes) { [math]::Max([double]0, ($r.SessionEnd - $r.SessionStart).TotalHours - [double]$r.ActiveSeconds / 3600) }
                         else { [double]$r.InactiveSeconds / 3600 }
        $completion = if ($r.NeedsReview) { "Needs end time - $($r.Reason)" } else { $r.Reason }

        [void]$sb.Append("<row r=`"$row`">")
        if ($hasDate) { [void]$sb.Append((Format-XlsxNumberCell "A$row" $date.ToOADate() 2)) }
        else { [void]$sb.Append((Format-XlsxTextCell "A$row" $r.Date 0)) }
        [void]$sb.Append((Format-XlsxFormulaCell "B$row" $dayFormula $(if ($hasDate) { $date.ToString('ddd', $script:Culture) } else { '' }) 0 -TextResult))
        [void]$sb.Append((Format-XlsxNumberCell "C$row" $(if ($null -ne $r.SessionStart) { $r.SessionStart.ToOADate() }) 3))
        [void]$sb.Append((Format-XlsxNumberCell "D$row" $(if ($null -ne $r.SessionEnd) { $r.SessionEnd.ToOADate() }) 3))
        [void]$sb.Append((Format-XlsxNumberCell "E$row" $inactiveHours 5))
        if ($hasTimes) { [void]$sb.Append((Format-XlsxFormulaCell "F$row" $hoursFormula ([double]$r.ActiveHours) 4)) }
        else { [void]$sb.Append((Format-XlsxNumberCell "F$row" ([double]$r.ActiveHours) 4)) }   # old record without times
        [void]$sb.Append((Format-XlsxTextCell "G$row" $r.InactivePeriods 0))
        [void]$sb.Append((Format-XlsxTextCell "H$row" $completion 0))
        [void]$sb.Append('</row>')
    }
    if ($records.Count -eq 0) {
        [void]$sb.Append("<row r=`"2`"><c r=`"A2`" s=`"2`"/>$(Format-XlsxFormulaCell 'B2' $dayFormula '' 0 -TextResult)<c r=`"C2`" s=`"3`"/><c r=`"D2`" s=`"3`"/><c r=`"E2`" s=`"5`"/>$(Format-XlsxFormulaCell 'F2' $hoursFormula '' 4 -TextResult)</row>")
    }

    $totalInactive = [double]0; $totalWorked = [double]0
    foreach ($r in $records) {
        $totalWorked += [double]$r.ActiveHours
        $totalInactive += if ($null -ne $r.SessionStart -and $null -ne $r.SessionEnd) { [math]::Max([double]0, ($r.SessionEnd - $r.SessionStart).TotalHours - [double]$r.ActiveSeconds / 3600) } else { [double]$r.InactiveSeconds / 3600 }
    }
    [void]$sb.Append("<row r=`"$totalRow`">")
    [void]$sb.Append((Format-XlsxTextCell "A$totalRow" 'Total' 0))
    [void]$sb.Append((Format-XlsxFormulaCell "E$totalRow" "SUBTOTAL(109,$table[Inactive Hours])" $totalInactive 5))
    [void]$sb.Append((Format-XlsxFormulaCell "F$totalRow" "SUBTOTAL(109,$table[Hours Worked])" $totalWorked 4))
    [void]$sb.Append('</row></sheetData>')
    [void]$sb.Append('<pageMargins left="0.5" right="0.5" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>')
    [void]$sb.Append('<tableParts count="1"><tablePart r:id="rId1"/></tableParts></worksheet>')

    $tableXml = "<?xml version=`"1.0`" encoding=`"UTF-8`" standalone=`"yes`"?>" +
        "<table xmlns=`"http://schemas.openxmlformats.org/spreadsheetml/2006/main`" id=`"1`" name=`"$table`" displayName=`"$table`" ref=`"A1:H$totalRow`" totalsRowCount=`"1`">" +
        "<autoFilter ref=`"A1:H$lastDataRow`"/><tableColumns count=`"8`">" +
        '<tableColumn id="1" name="Date" totalsRowLabel="Total"/>' +
        "<tableColumn id=`"2`" name=`"Day`"><calculatedColumnFormula>$(ConvertTo-XmlText $dayFormula)</calculatedColumnFormula></tableColumn>" +
        '<tableColumn id="3" name="Start"/><tableColumn id="4" name="End"/>' +
        '<tableColumn id="5" name="Inactive Hours" totalsRowFunction="sum"/>' +
        "<tableColumn id=`"6`" name=`"Hours Worked`" totalsRowFunction=`"sum`"><calculatedColumnFormula>$(ConvertTo-XmlText $hoursFormula)</calculatedColumnFormula></tableColumn>" +
        '<tableColumn id="7" name="Inactive Periods"/><tableColumn id="8" name="Completion"/></tableColumns>' +
        '<tableStyleInfo name="TableStyleMedium2" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/></table>'

    $parts = [ordered]@{
        '[Content_Types].xml' = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/tables/table1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>'
        '_rels/.rels' = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'
        'xl/workbook.xml' = "<?xml version=`"1.0`" encoding=`"UTF-8`" standalone=`"yes`"?><workbook $ns><sheets><sheet name=`"Work History`" sheetId=`"1`" r:id=`"rId1`"/></sheets><calcPr calcId=`"191029`" fullCalcOnLoad=`"1`"/></workbook>"
        'xl/_rels/workbook.xml.rels' = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>'
        'xl/worksheets/_rels/sheet1.xml.rels' = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"/></Relationships>'
        'xl/tables/table1.xml' = $tableXml
        'xl/styles.xml' = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><numFmts count="3"><numFmt numFmtId="164" formatCode="yyyy-mm-dd"/><numFmt numFmtId="165" formatCode="hh:mm"/><numFmt numFmtId="166" formatCode="0.0"/></numFmts><fonts count="1"><font><sz val="11"/><name val="Calibri"/><family val="2"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="6"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/><xf numFmtId="165" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/><xf numFmtId="166" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/><xf numFmtId="2" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>'
        'xl/worksheets/sheet1.xml' = $sb.ToString()
    }

    Initialize-Storage
    $temp = Join-Path $script:AppDirectory ('.workbook.{0}.tmp' -f (New-Id))
    try {
        Write-ZipPackage -Path $temp -Parts $parts
        try {
            [IO.File]::Move($temp, $script:WorkbookPath, $true)
            $script:ExcelPending = $false
        }
        catch [IO.IOException], [UnauthorizedAccessException] {
            $script:ExcelPending = $true
        }
    }
    finally {
        if (Test-Path -LiteralPath $temp) { Remove-Item -LiteralPath $temp -Force -ErrorAction SilentlyContinue }
    }
}

# ---------------------------------------------------------------------------
# MARK: Start with Windows (per-user Startup folder; no admin rights or Task Scheduler needed)
# ---------------------------------------------------------------------------

function Get-LaunchCommand {
    if ($null -eq $script:ScriptPath) { throw 'Save the script to a file first.' }
    $pwsh = (Get-Process -Id $PID).Path
    # Store-installed PowerShell lives in a versioned, protected folder that changes on every
    # update. Use the stable per-user alias instead.
    if ($pwsh -like '*\WindowsApps\*') {
        $alias = Join-Path $env:LOCALAPPDATA 'Microsoft\WindowsApps\pwsh.exe'
        if (Test-Path -LiteralPath $alias) { $pwsh = $alias }
    }
    $pwshArguments = '-NoProfile -STA -WindowStyle Hidden -File "{0}"' -f $script:ScriptPath
    $conhost = Join-Path $env:SystemRoot 'System32\conhost.exe'
    # conhost --headless runs PowerShell with no console window at all.
    $useHeadless = (Test-Path -LiteralPath $conhost) -and [Environment]::OSVersion.Version.Build -ge 19041
    [pscustomobject]@{
        Target           = if ($useHeadless) { $conhost } else { $pwsh }
        Arguments        = if ($useHeadless) { '--headless "{0}" {1}' -f $pwsh, $pwshArguments } else { $pwshArguments }
        WorkingDirectory = Split-Path -Parent $script:ScriptPath
        Icon             = "$(Get-AppIconPath),0"
    }
}

function New-AppShortcut {
    param([Parameter(Mandatory)][string]$Path)
    $launch = Get-LaunchCommand
    $shell = New-Object -ComObject WScript.Shell
    $shortcut = $null
    try {
        $shortcut = $shell.CreateShortcut($Path)
        $shortcut.TargetPath = $launch.Target
        $shortcut.Arguments = $launch.Arguments
        $shortcut.WorkingDirectory = $launch.WorkingDirectory
        $shortcut.IconLocation = $launch.Icon
        $shortcut.Description = 'Presence-aware work timer'
        $shortcut.WindowStyle = 1
        $shortcut.Save()
    }
    finally {
        if ($null -ne $shortcut) { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($shortcut) }
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($shell)
    }
    # Tag the shortcut with the app's taskbar identity so the running window matches it.
    try { [PresenceWorkTimer.V4.AppIdentity]::SetShortcutAppId($Path, $script:AppUserModelId) }
    catch { Write-AppLog "Could not set the taskbar identity on '$Path'." $_ }
}

function Get-StartupShortcutPath { Join-Path ([Environment]::GetFolderPath('Startup')) $script:ShortcutName }
function Get-StartMenuShortcutPath { Join-Path ([Environment]::GetFolderPath('Programs')) $script:ShortcutName }

function Test-StartWithWindows { Test-Path -LiteralPath (Get-StartupShortcutPath) -PathType Leaf }

function Set-StartWithWindows {
    param([Parameter(Mandatory)][bool]$Enabled)
    Remove-LegacyStartupTask
    $path = Get-StartupShortcutPath
    if ($Enabled) { New-AppShortcut -Path $path }
    elseif (Test-Path -LiteralPath $path) { Remove-Item -LiteralPath $path -Force }
}

function Remove-LegacyStartupTask {
    # Version 1 registered a Task Scheduler task that could not start Store-installed PowerShell.
    $service = $folder = $null
    try {
        $service = New-Object -ComObject 'Schedule.Service'
        $service.Connect()
        $folder = $service.GetFolder('\')
        $folder.DeleteTask($script:LegacyTaskName, 0)
    }
    catch { }
    finally {
        foreach ($com in @($folder, $service)) {
            if ($null -ne $com) { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($com) }
        }
    }
}

# ---------------------------------------------------------------------------
# MARK: UI toolkit
# ---------------------------------------------------------------------------

function Initialize-Theme {
    $graphics = [Drawing.Graphics]::FromHwnd([IntPtr]::Zero)
    try { $script:Scale = $graphics.DpiX / 96.0 } finally { $graphics.Dispose() }

    $palette = @{
        Window = '#F3F5F8'; Surface = '#FFFFFF'; Border = '#E3E8EF'; Header = '#14213D'; HeaderMuted = '#A9B4C8'
        Text = '#111827'; Muted = '#6B7280'; Subtle = '#F8FAFC'; Accent = '#2563EB'; AccentHover = '#1D4ED8'
        AccentSoft = '#DBEAFE'; Success = '#059669'; Warning = '#D97706'; Danger = '#DC2626'; DangerSoft = '#FEF2F2'
        DangerBorder = '#FCA5A5'; Neutral = '#9CA3AF'; Track = '#E5E7EB'; ButtonBorder = '#CBD5E1'; Hover = '#F1F5F9'
    }
    foreach ($key in $palette.Keys) { $script:Colors[$key] = [Drawing.ColorTranslator]::FromHtml($palette[$key]) }

    $installed = @{}
    foreach ($family in [Drawing.FontFamily]::Families) { $installed[$family.Name] = $true }
    $font = {
        param([string]$Family, [single]$Size, [Drawing.FontStyle]$Style = [Drawing.FontStyle]::Regular)
        if ($installed.ContainsKey($Family)) { return [Drawing.Font]::new($Family, $Size, $Style) }
        $fallbackStyle = if ($Family -like '*Semibold*') { [Drawing.FontStyle]::Bold } else { $Style }
        [Drawing.Font]::new('Segoe UI', $Size, $fallbackStyle)
    }
    $script:Fonts = @{
        Title     = & $font 'Segoe UI Semibold' 15
        Heading   = & $font 'Segoe UI Semibold' 12.5
        Body      = & $font 'Segoe UI' 9.75
        Strong    = & $font 'Segoe UI Semibold' 9.75
        Small     = & $font 'Segoe UI' 8.75
        Caption   = & $font 'Segoe UI Semibold' 8
        Hero      = & $font 'Segoe UI Light' 40
        HeroUnit  = & $font 'Segoe UI' 11
        Stat      = & $font 'Segoe UI Semibold' 12
        Clock     = & $font 'Segoe UI Light' 18
        Countdown = & $font 'Segoe UI Semibold' 10
    }
    $script:AppIcon = New-AppIcon
}

function New-AppIconBitmap {
    # Draws the Work Timer icon: a stopwatch on a blue gradient tile with a green progress arc.
    param([Parameter(Mandatory)][int]$Size)
    $s = [single]$Size
    $bitmap = [Drawing.Bitmap]::new($Size, $Size, [Drawing.Imaging.PixelFormat]::Format32bppArgb)
    $g = [Drawing.Graphics]::FromImage($bitmap)
    $disposables = [Collections.Generic.List[IDisposable]]::new()
    try {
        $g.SmoothingMode = [Drawing.Drawing2D.SmoothingMode]::AntiAlias
        $g.PixelOffsetMode = [Drawing.Drawing2D.PixelOffsetMode]::HighQuality
        $g.Clear([Drawing.Color]::Transparent)
        $color = { param([string]$Html, [int]$Alpha = 255) [Drawing.Color]::FromArgb($Alpha, [Drawing.ColorTranslator]::FromHtml($Html)) }
        $track = { param($obj) $disposables.Add($obj); $obj }

        # Rounded-square tile with a diagonal gradient.
        $inset = [math]::Max([single]0.5, $s * 0.03)
        $tile = [Drawing.RectangleF]::new($inset, $inset, $s - 2 * $inset, $s - 2 * $inset)
        $radius = $tile.Width * 0.24
        $path = & $track ([Drawing.Drawing2D.GraphicsPath]::new())
        $d = $radius * 2
        $path.AddArc($tile.X, $tile.Y, $d, $d, 180, 90)
        $path.AddArc($tile.Right - $d, $tile.Y, $d, $d, 270, 90)
        $path.AddArc($tile.Right - $d, $tile.Bottom - $d, $d, $d, 0, 90)
        $path.AddArc($tile.X, $tile.Bottom - $d, $d, $d, 90, 90)
        $path.CloseFigure()
        $gradient = & $track ([Drawing.Drawing2D.LinearGradientBrush]::new($tile, (& $color '#1E3A8A'), (& $color '#0EA5E9'), [single]45))
        $g.FillPath($gradient, $path)
        # Soft top sheen for depth (skipped at tiny sizes where it just looks muddy).
        if ($Size -ge 32) {
            $sheen = & $track ([Drawing.Drawing2D.LinearGradientBrush]::new(
                [Drawing.RectangleF]::new($tile.X, $tile.Y - 1, $tile.Width, $tile.Height * 0.6 + 2),
                (& $color '#FFFFFF' 46), (& $color '#FFFFFF' 0), [single]90))
            $g.SetClip($path)
            $g.FillRectangle($sheen, $tile.X, $tile.Y, $tile.Width, $tile.Height * 0.6)
            $g.ResetClip()
        }

        $cx = $s * 0.5
        $cy = $s * 0.555
        $white = & $color '#FFFFFF'

        # Crown and stem on top of the stopwatch.
        $whiteBrush = & $track ([Drawing.SolidBrush]::new($white))
        $g.FillRectangle($whiteBrush, $cx - $s * 0.045, $s * 0.175, $s * 0.09, $s * 0.08)
        $crown = & $track ([Drawing.Drawing2D.GraphicsPath]::new())
        $cw = $s * 0.20; $ch = $s * 0.075; $cr = $ch
        $crown.AddArc($cx - $cw / 2, $s * 0.125, $cr, $ch, 90, 180)
        $crown.AddArc($cx + $cw / 2 - $cr, $s * 0.125, $cr, $ch, 270, 180)
        $crown.CloseFigure()
        $g.FillPath($whiteBrush, $crown)
        # Side button at roughly the 1:30 position.
        if ($Size -ge 24) {
            $state = $g.Save()
            $g.TranslateTransform($cx, $cy)
            $g.RotateTransform(45)
            $g.FillRectangle($whiteBrush, -$s * 0.035, -$s * 0.375, $s * 0.07, $s * 0.07)
            $g.Restore($state)
        }

        # Outer ring.
        $ringRadius = $s * 0.29
        $ringPen = & $track ([Drawing.Pen]::new($white, $(if ($Size -lt 20) { [single]1.6 } else { [math]::Max([single]1.2, $s * 0.062) })))
        $g.DrawEllipse($ringPen, $cx - $ringRadius, $cy - $ringRadius, 2 * $ringRadius, 2 * $ringRadius)

        # Green progress arc: "time worked so far".
        $arcRadius = if ($Size -lt 20) { $s * 0.17 } else { $s * 0.19 }
        $arcPen = & $track ([Drawing.Pen]::new((& $color '#34D399'), $(if ($Size -lt 20) { [single]2.2 } else { [math]::Max([single]1.2, $s * 0.085) })))
        $arcPen.StartCap = [Drawing.Drawing2D.LineCap]::Round
        $arcPen.EndCap = [Drawing.Drawing2D.LineCap]::Round
        $g.DrawArc($arcPen, $cx - $arcRadius, $cy - $arcRadius, 2 * $arcRadius, 2 * $arcRadius, [single]-90, [single]250)

        if ($Size -lt 20) { return $bitmap }
        # Hand pointing to the end of the arc, plus hub.
        $angle = (-90 + 250) * [math]::PI / 180
        $handPen = & $track ([Drawing.Pen]::new($white, [math]::Max([single]1, $s * 0.045)))
        $handPen.StartCap = [Drawing.Drawing2D.LineCap]::Round
        $handPen.EndCap = [Drawing.Drawing2D.LineCap]::Round
        $handLength = $s * 0.145
        $g.DrawLine($handPen, $cx, $cy, [single]($cx + [math]::Cos($angle) * $handLength), [single]($cy + [math]::Sin($angle) * $handLength))
        $hub = [math]::Max([single]1, $s * 0.04)
        $g.FillEllipse($whiteBrush, $cx - $hub, $cy - $hub, 2 * $hub, 2 * $hub)
    }
    finally {
        foreach ($item in $disposables) { $item.Dispose() }
        $g.Dispose()
    }
    $bitmap
}

function ConvertTo-IconImageBytes {
    # One .ico frame as a classic 32-bit DIB (readable by Explorer, Win32 and .NET alike).
    param([Parameter(Mandatory)][Drawing.Bitmap]$Bitmap)
    $size = $Bitmap.Width
    $stream = [IO.MemoryStream]::new()
    try {
        $data = $Bitmap.LockBits([Drawing.Rectangle]::new(0, 0, $size, $size), [Drawing.Imaging.ImageLockMode]::ReadOnly, [Drawing.Imaging.PixelFormat]::Format32bppArgb)
        try {
            $stride = $data.Stride
            $pixels = [byte[]]::new($stride * $size)
            [Runtime.InteropServices.Marshal]::Copy($data.Scan0, $pixels, 0, $pixels.Length)
        }
        finally { $Bitmap.UnlockBits($data) }

        $writer = [IO.BinaryWriter]::new($stream)
        # BITMAPINFOHEADER: height is doubled to cover the (unused) AND mask.
        $writer.Write([int32]40); $writer.Write([int32]$size); $writer.Write([int32]($size * 2))
        $writer.Write([uint16]1); $writer.Write([uint16]32); $writer.Write([int32]0)
        $writer.Write([int32]0); $writer.Write([int32]0); $writer.Write([int32]0); $writer.Write([int32]0); $writer.Write([int32]0)
        for ($y = $size - 1; $y -ge 0; $y--) { $writer.Write($pixels, $y * $stride, $size * 4) }   # bottom-up BGRA
        $maskRow = [int]([math]::Ceiling($size / 32.0) * 4)
        $writer.Write([byte[]]::new($maskRow * $size))   # alpha channel is used, so the mask is all zero
        $writer.Flush()
        return , $stream.ToArray()
    }
    finally { $stream.Dispose() }
}

function Save-AppIconFile {
    param([Parameter(Mandatory)][string]$Path)
    $sizes = @(16, 20, 24, 32, 40, 48, 64, 256)
    $images = foreach ($size in $sizes) {
        $bitmap = New-AppIconBitmap $size
        try { , (ConvertTo-IconImageBytes $bitmap) } finally { $bitmap.Dispose() }
    }
    $stream = [IO.MemoryStream]::new()
    try {
        $writer = [IO.BinaryWriter]::new($stream)
        $writer.Write([uint16]0); $writer.Write([uint16]1); $writer.Write([uint16]$sizes.Count)   # ICONDIR
        $offset = 6 + 16 * $sizes.Count
        for ($n = 0; $n -lt $sizes.Count; $n++) {
            $dimension = if ($sizes[$n] -ge 256) { 0 } else { $sizes[$n] }   # 0 means 256
            $writer.Write([byte]$dimension); $writer.Write([byte]$dimension); $writer.Write([byte]0); $writer.Write([byte]0)
            $writer.Write([uint16]1); $writer.Write([uint16]32)
            $writer.Write([uint32]$images[$n].Length); $writer.Write([uint32]$offset)
            $offset += $images[$n].Length
        }
        foreach ($image in $images) { $writer.Write([byte[]]$image) }
        $writer.Flush()
        Initialize-Storage
        $temp = "$Path.$(New-Id).tmp"
        try { [IO.File]::WriteAllBytes($temp, $stream.ToArray()); [IO.File]::Move($temp, $Path, $true) }
        finally { if (Test-Path -LiteralPath $temp) { Remove-Item -LiteralPath $temp -Force -ErrorAction SilentlyContinue } }
    }
    finally { $stream.Dispose() }
}

function Get-AppIconPath {
    # The version is in the file name so a redesigned icon gets a new path (Windows caches icons by path).
    $path = Join-Path $script:AppDirectory 'Work Timer v1.ico'
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { Save-AppIconFile -Path $path }
    $path
}

function New-AppIcon {
    try { return [Drawing.Icon]::new((Get-AppIconPath)) }
    catch {
        Write-AppLog 'Could not create the icon file; using a built-in icon.' $_
        $bitmap = New-AppIconBitmap 64
        try { return [Drawing.Icon]::FromHandle($bitmap.GetHicon()) } finally { $bitmap.Dispose() }
    }
}

function Update-ExistingShortcuts {
    # Refreshes this app's shortcuts (icon, PowerShell path) if they exist and point at this script.
    if ([string]::IsNullOrWhiteSpace($script:ScriptPath)) { return }
    foreach ($path in @((Get-StartupShortcutPath), (Get-StartMenuShortcutPath))) {
        if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { continue }
        try {
            $shell = New-Object -ComObject WScript.Shell
            try {
                $existing = $shell.CreateShortcut($path)
                $ours = ([string]$existing.Arguments).Contains([string]$script:ScriptPath)
                [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($existing)
            }
            finally { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($shell) }
            if ($ours) { New-AppShortcut -Path $path }
        }
        catch { Write-AppLog "Could not refresh shortcut '$path'." $_ }
    }
}

function Px { param([double]$Value) [int][math]::Round($Value * $script:Scale) }   # design units -> pixels
function New-Rect { param([double]$X, [double]$Y, [double]$W, [double]$H) [Drawing.Rectangle]::new((Px $X), (Px $Y), (Px $W), (Px $H)) }

function New-Label {
    param(
        [string]$Text = '', [double]$X = 0, [double]$Y = 0,
        [Drawing.Font]$Font = $script:Fonts.Body, [Drawing.Color]$Color = $script:Colors.Text,
        [double]$Width = 0, [double]$Height = 0, [Drawing.ContentAlignment]$Align = [Drawing.ContentAlignment]::TopLeft
    )
    $label = [Windows.Forms.Label]::new()
    $label.Text = $Text
    $label.Font = $Font
    $label.ForeColor = $Color
    $label.BackColor = [Drawing.Color]::Transparent
    $label.UseMnemonic = $false
    if ($Width -gt 0) {
        $label.AutoSize = $false
        $label.Bounds = New-Rect $X $Y $Width $(if ($Height -gt 0) { $Height } else { 22 })
        $label.TextAlign = $Align
    }
    else {
        $label.AutoSize = $true
        $label.Location = [Drawing.Point]::new((Px $X), (Px $Y))
    }
    $label
}

function Set-ButtonStyle {
    param([Parameter(Mandatory)][Windows.Forms.Button]$Button)
    $c = $script:Colors
    $variant = [string]$Button.Tag
    if (-not $Button.Enabled) {
        $Button.BackColor = $c.Hover; $Button.ForeColor = $c.Neutral; $Button.FlatAppearance.BorderColor = $c.Border
        return
    }
    switch ($variant) {
        'Primary' { $back = $c.Accent; $fore = [Drawing.Color]::White; $border = $c.Accent; $hover = $c.AccentHover }
        'Danger' { $back = $c.Surface; $fore = $c.Danger; $border = $c.DangerBorder; $hover = $c.DangerSoft }
        default { $back = $c.Surface; $fore = $c.Text; $border = $c.ButtonBorder; $hover = $c.Hover }
    }
    $Button.BackColor = $back
    $Button.ForeColor = $fore
    $Button.FlatAppearance.BorderColor = $border
    $Button.FlatAppearance.MouseOverBackColor = $hover
    $Button.FlatAppearance.MouseDownBackColor = $hover
}

function New-Button {
    param(
        [string]$Text, [double]$X, [double]$Y, [double]$Width, [double]$Height = 40,
        [ValidateSet('Primary', 'Secondary', 'Danger')][string]$Variant = 'Secondary'
    )
    $button = [Windows.Forms.Button]::new()
    $button.Text = $Text
    $button.Bounds = New-Rect $X $Y $Width $Height
    $button.FlatStyle = [Windows.Forms.FlatStyle]::Flat
    $button.FlatAppearance.BorderSize = 1
    $button.Font = $script:Fonts.Strong
    $button.Cursor = [Windows.Forms.Cursors]::Hand
    $button.UseMnemonic = $false
    $button.Tag = $Variant
    Set-ButtonStyle $button
    $button.Add_EnabledChanged({ param($source, $e) Set-ButtonStyle $source })
    $button
}

function New-Card {
    param([double]$X, [double]$Y, [double]$Width, [double]$Height)
    $card = [Windows.Forms.Panel]::new()
    $card.Bounds = New-Rect $X $Y $Width $Height
    $card.BackColor = $script:Colors.Surface
    $card.Add_Paint({
        param($source, $e)
        $pen = [Drawing.Pen]::new($script:Colors.Border)
        try { $e.Graphics.DrawRectangle($pen, 0, 0, $source.Width - 1, $source.Height - 1) } finally { $pen.Dispose() }
    })
    $card
}

function New-TimePicker {
    param([double]$X, [double]$Y, [double]$Width = 130)
    $picker = [Windows.Forms.DateTimePicker]::new()
    $picker.Format = [Windows.Forms.DateTimePickerFormat]::Custom
    $picker.CustomFormat = 'HH:mm:ss'
    $picker.ShowUpDown = $true
    $picker.Font = $script:Fonts.Body
    $picker.Bounds = New-Rect $X $Y $Width 28
    $picker
}

function New-Grid {
    param([double]$X, [double]$Y, [double]$Width, [double]$Height, [string[]]$HiddenColumns = @())
    $grid = [Windows.Forms.DataGridView]::new()
    $grid.Bounds = New-Rect $X $Y $Width $Height
    $grid.ReadOnly = $true
    $grid.AllowUserToAddRows = $false
    $grid.AllowUserToDeleteRows = $false
    $grid.AllowUserToResizeRows = $false
    $grid.MultiSelect = $false
    $grid.SelectionMode = [Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
    $grid.AutoSizeColumnsMode = [Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill
    $grid.RowHeadersVisible = $false
    $grid.BorderStyle = [Windows.Forms.BorderStyle]::None
    $grid.BackgroundColor = $script:Colors.Surface
    $grid.GridColor = $script:Colors.Border
    $grid.CellBorderStyle = [Windows.Forms.DataGridViewCellBorderStyle]::SingleHorizontal
    $grid.EnableHeadersVisualStyles = $false
    $grid.ColumnHeadersBorderStyle = [Windows.Forms.DataGridViewHeaderBorderStyle]::None
    $grid.ColumnHeadersHeightSizeMode = [Windows.Forms.DataGridViewColumnHeadersHeightSizeMode]::DisableResizing
    $grid.ColumnHeadersHeight = Px 34
    $grid.ColumnHeadersDefaultCellStyle.BackColor = $script:Colors.Subtle
    $grid.ColumnHeadersDefaultCellStyle.ForeColor = $script:Colors.Muted
    $grid.ColumnHeadersDefaultCellStyle.Font = $script:Fonts.Caption
    $grid.ColumnHeadersDefaultCellStyle.SelectionBackColor = $script:Colors.Subtle
    $grid.ColumnHeadersDefaultCellStyle.Padding = [Windows.Forms.Padding]::new((Px 6), 0, 0, 0)
    $grid.DefaultCellStyle.Font = $script:Fonts.Body
    $grid.DefaultCellStyle.ForeColor = $script:Colors.Text
    $grid.DefaultCellStyle.SelectionBackColor = $script:Colors.AccentSoft
    $grid.DefaultCellStyle.SelectionForeColor = $script:Colors.Text
    $grid.DefaultCellStyle.Padding = [Windows.Forms.Padding]::new((Px 6), 0, 0, 0)
    $grid.AlternatingRowsDefaultCellStyle.BackColor = $script:Colors.Subtle
    $grid.RowTemplate.Height = Px 30
    $grid.Tag = $HiddenColumns
    $grid.Add_DataBindingComplete({
        param($source, $e)
        foreach ($name in @($source.Tag)) { if ($source.Columns.Contains($name)) { $source.Columns[$name].Visible = $false } }
        foreach ($column in $source.Columns) { $column.SortMode = [Windows.Forms.DataGridViewColumnSortMode]::Automatic }
    })
    $grid
}

function New-DialogForm {
    param([string]$Title, [double]$Width, [double]$Height, [switch]$Resizable)
    $form = [Windows.Forms.Form]::new()
    $form.Text = $Title
    $form.AutoScaleMode = [Windows.Forms.AutoScaleMode]::None
    $form.Font = $script:Fonts.Body
    $form.BackColor = $script:Colors.Surface
    $form.ForeColor = $script:Colors.Text
    $form.ClientSize = [Drawing.Size]::new((Px $Width), (Px $Height))
    $form.StartPosition = [Windows.Forms.FormStartPosition]::CenterParent
    $form.ShowInTaskbar = $false
    $form.MinimizeBox = $false
    $form.Icon = $script:AppIcon
    if ($Resizable) {
        $form.FormBorderStyle = [Windows.Forms.FormBorderStyle]::Sizable
        $form.MinimumSize = $form.Size
    }
    else {
        $form.FormBorderStyle = [Windows.Forms.FormBorderStyle]::FixedDialog
        $form.MaximizeBox = $false
    }
    $form
}

function Show-Modal {
    param([Parameter(Mandatory)][Windows.Forms.Form]$Form)
    $owner = [Windows.Forms.Form]::ActiveForm
    if ($null -eq $owner -and $null -ne $script:MainForm -and -not $script:MainForm.IsDisposed -and $script:MainForm.Visible) {
        $owner = $script:MainForm
    }
    if ($null -eq $owner) {
        $Form.StartPosition = [Windows.Forms.FormStartPosition]::CenterScreen
        $Form.TopMost = $true
        return $Form.ShowDialog()
    }
    $Form.ShowDialog($owner)
}

function Show-Choice {
    # Styled replacement for MessageBox. Returns $true when the primary button is chosen.
    param(
        [string]$Title = $script:AppName, [Parameter(Mandatory)][string]$Heading, [string]$Message = '',
        [string]$PrimaryText = 'OK', [string]$SecondaryText = '',
        [ValidateSet('Info', 'Warning', 'Error', 'Danger')][string]$Kind = 'Info'
    )
    $accent = switch ($Kind) { 'Warning' { $script:Colors.Warning } 'Error' { $script:Colors.Danger } 'Danger' { $script:Colors.Danger } default { $script:Colors.Accent } }
    $form = New-DialogForm -Title $Title -Width 460 -Height 200
    $stripe = [Windows.Forms.Panel]::new(); $stripe.Dock = 'Top'; $stripe.Height = Px 4; $stripe.BackColor = $accent

    $headingLabel = New-Label $Heading 24 22 $script:Fonts.Heading
    $headingLabel.MaximumSize = [Drawing.Size]::new((Px 412), 0)
    $messageLabel = New-Label $Message 24 0 $script:Fonts.Body $script:Colors.Text
    $messageLabel.MaximumSize = [Drawing.Size]::new((Px 412), 0)
    $form.Controls.AddRange(@($stripe, $headingLabel, $messageLabel))
    $messageLabel.Top = $headingLabel.Bottom + (Px 10)

    $footer = [Windows.Forms.Panel]::new(); $footer.Dock = 'Bottom'; $footer.Height = Px 64; $footer.BackColor = $script:Colors.Subtle
    $primaryVariant = if ($Kind -eq 'Danger') { 'Danger' } else { 'Primary' }
    $primary = New-Button $PrimaryText 0 14 130 36 $primaryVariant
    $primary.DialogResult = [Windows.Forms.DialogResult]::OK
    $primary.Left = $form.ClientSize.Width - $primary.Width - (Px 20)
    $footer.Controls.Add($primary)
    $form.AcceptButton = $primary
    if ($SecondaryText) {
        $secondary = New-Button $SecondaryText 0 14 110 36 'Secondary'
        $secondary.DialogResult = [Windows.Forms.DialogResult]::Cancel
        $secondary.Left = $primary.Left - $secondary.Width - (Px 8)
        $footer.Controls.Add($secondary)
        $form.CancelButton = $secondary
    }
    else { $form.CancelButton = $primary }
    $form.Controls.Add($footer)
    $bodyBottom = if ($Message) { $messageLabel.Bottom } else { $headingLabel.Bottom }
    $form.ClientSize = [Drawing.Size]::new($form.ClientSize.Width, $bodyBottom + (Px 26) + $footer.Height)

    try { return (Show-Modal $form) -eq [Windows.Forms.DialogResult]::OK }
    finally { $form.Dispose() }
}

function Invoke-UiAction {
    # Runs a UI action and reports failures instead of crashing the timer.
    param([Parameter(Mandatory)][string]$Context, [Parameter(Mandatory)][scriptblock]$Action)
    try { & $Action }
    catch {
        Write-AppLog "$Context failed." $_
        [void](Show-Choice -Heading "$Context failed" -Message $_.Exception.Message -Kind Error)
    }
}

# ---------------------------------------------------------------------------
# MARK: Main window
# ---------------------------------------------------------------------------

function Initialize-MainWindow {
    $c = $script:Colors
    $f = $script:Fonts
    $ui = $script:Ui

    $form = [Windows.Forms.Form]::new()
    $form.Text = $script:AppName
    $form.AutoScaleMode = [Windows.Forms.AutoScaleMode]::None
    $form.ClientSize = [Drawing.Size]::new((Px 460), (Px 522))
    $form.FormBorderStyle = [Windows.Forms.FormBorderStyle]::FixedSingle
    $form.MaximizeBox = $false
    $form.StartPosition = [Windows.Forms.FormStartPosition]::CenterScreen
    $form.BackColor = $c.Window
    $form.Font = $f.Body
    $form.Icon = $script:AppIcon
    $form.KeyPreview = $true
    $script:MainForm = $form

    # Header
    $header = [Windows.Forms.Panel]::new(); $header.Dock = 'Top'; $header.Height = Px 72; $header.BackColor = $c.Header
    $ui.Title = New-Label 'Work Timer' 22 11 $f.Title ([Drawing.Color]::White)
    $ui.Date = New-Label '' 23 43 $f.Small $c.HeaderMuted
    $ui.Clock = New-Label '' 270 12 $f.Clock ([Drawing.Color]::White) 170 34 ([Drawing.ContentAlignment]::MiddleRight)
    $header.Controls.AddRange(@($ui.Title, $ui.Date, $ui.Clock))

    # Hero card: status, hours, countdown
    $hero = New-Card 16 88 428 196
    $ui.StatusDot = New-Label ([string][char]0x25CF) 16 13 $f.Strong $c.Neutral
    $ui.Status = New-Label '' 32 14 $f.Strong $c.Text
    $ui.Hours = New-Label '0.0' 12 32 $f.Hero $c.Text
    $ui.HoursUnit = New-Label 'hours active' 120 78 $f.HeroUnit $c.Muted
    $ui.CountdownCaption = New-Label '' 20 122 $f.Small $c.Muted
    $ui.CountdownValue = New-Label '' 290 119 $f.Countdown $c.Text 118 24 ([Drawing.ContentAlignment]::MiddleRight)
    $ui.Track = [Windows.Forms.Panel]::new(); $ui.Track.Bounds = New-Rect 20 148 388 8; $ui.Track.BackColor = $c.Track
    $ui.TrackFill = [Windows.Forms.Panel]::new(); $ui.TrackFill.Bounds = [Drawing.Rectangle]::new(0, 0, 0, $ui.Track.Height); $ui.TrackFill.BackColor = $c.Accent
    $ui.Track.Controls.Add($ui.TrackFill)
    $ui.Exact = New-Label '' 20 166 $f.Small $c.Muted
    $hero.Controls.AddRange(@($ui.StatusDot, $ui.Status, $ui.Hours, $ui.HoursUnit, $ui.CountdownCaption, $ui.CountdownValue, $ui.Track, $ui.Exact))

    # Stats card
    $stats = New-Card 16 296 428 78
    $captions = @('STARTED', 'INACTIVE', 'PRESENCE')
    $values = @()
    for ($n = 0; $n -lt 3; $n++) {
        $x = 20 + ($n * 140)
        $stats.Controls.Add((New-Label $captions[$n] $x 14 $f.Caption $c.Muted))
        $value = New-Label '-' $x 34 $f.Stat $c.Text
        $stats.Controls.Add($value)
        $values += $value
    }
    $ui.Started = $values[0]; $ui.Inactive = $values[1]; $ui.Presence = $values[2]

    # Actions
    $ui.EndButton = New-Button 'End Workday' 16 388 210 42 'Primary'
    $ui.PauseButton = New-Button 'Pause Timer' 234 388 210 42
    $ui.IdleButton = New-Button 'Idle Periods' 16 438 210 42
    $ui.HistoryButton = New-Button 'Work History' 234 438 210 42
    $ui.Footer = New-Label '' 18 494 $f.Small $c.Muted 424 20

    $ui.EndButton.Add_Click({ Invoke-UiAction 'End workday' { Request-EndWorkday } })
    $ui.PauseButton.Add_Click({ Invoke-UiAction 'Pause' { Switch-Pause } })
    $ui.IdleButton.Add_Click({ Invoke-UiAction 'Idle periods' { Show-PeriodEditor } })
    $ui.HistoryButton.Add_Click({ Invoke-UiAction 'Work history' { Show-WorkHistory } })
    $ui.Footer.Add_Click({
        if (@(Get-PendingReviews).Count -gt 0) { Invoke-UiAction 'Previous workday' { Show-PendingReviews } }
    })

    $form.Controls.AddRange(@($hero, $stats, $ui.EndButton, $ui.PauseButton, $ui.IdleButton, $ui.HistoryButton, $ui.Footer, $header))
    $form.Add_KeyDown({
        param($source, $e)
        if ($e.Control -and $e.Shift -and $e.KeyCode -eq [Windows.Forms.Keys]::A) {
            $e.SuppressKeyPress = $true
            Invoke-UiAction 'Advanced tools' { Show-AdvancedTools }
        }
    })
    $form.Add_FormClosing({
        param($source, $e)
        $s = $script:State
        if ($e.CloseReason -eq [Windows.Forms.CloseReason]::UserClosing -and $s.TrackingStarted -and -not $s.Ended) {
            $message = "Time isn't tracked while Work Timer is closed. If you reopen it today, you'll be asked whether to count the time it was closed."
            if (-not (Show-Choice -Heading 'Close Work Timer?' -Message $message -PrimaryText 'Close' -SecondaryText 'Keep Running' -Kind Warning)) {
                $e.Cancel = $true
                return
            }
        }
        $script:ShuttingDown = $true
        if ($null -ne $script:Timer) { $script:Timer.Stop() }
        try { Update-Sample; Save-State } catch { Write-AppLog 'Could not save state while closing.' $_ }
    })
    $form.Add_Shown({ Invoke-UiAction 'Startup' { Complete-Startup } })
    Set-WindowAppIdentity -Form $form
}

function Set-WindowAppIdentity {
    # Must run before the window is shown so the taskbar button gets Work Timer's icon, not PowerShell's.
    param([Parameter(Mandatory)][Windows.Forms.Form]$Form)
    try {
        $relaunch = $null
        if ($null -ne $script:ScriptPath) {
            $launch = Get-LaunchCommand
            $relaunch = '"{0}" {1}' -f $launch.Target, $launch.Arguments
        }
        $iconPath = try { Get-AppIconPath } catch { $null }
        $rejected = [PresenceWorkTimer.V4.AppIdentity]::SetWindowAppId($Form.Handle, $script:AppUserModelId, $relaunch, $iconPath, $script:AppName)
        if ($rejected) { Write-AppLog "Taskbar pin details not accepted by Windows: $rejected (the taskbar icon still works)." }
    }
    catch { Write-AppLog 'Could not set the taskbar identity for the window.' $_ }
}

function Complete-Startup {
    try {
        $script:MainForm.Activate()
        foreach ($warning in $script:StartupWarnings) { [void](Show-Choice -Heading 'Work Timer recovered your data' -Message $warning -Kind Warning) }
        $gap = Get-DowntimeGap
        if ($null -ne $gap) { Resolve-DowntimeGap -Gap $gap }
        if (@(Get-PendingReviews).Count -gt 0) { Show-PendingReviews }
    }
    finally {
        # Always start ticking, even if a startup prompt failed.
        $script:LastInputTick = [PresenceWorkTimer.V4.Native]::GetLastInputTick()
        $script:State.LastSampleTime = Get-Date
        $script:Timer.Start()
        Update-Display
    }
}

function Show-MainWindow {
    $form = $script:MainForm
    if ($null -eq $form -or $form.IsDisposed) { return }
    if ($form.WindowState -eq [Windows.Forms.FormWindowState]::Minimized) { $form.WindowState = [Windows.Forms.FormWindowState]::Normal }
    $form.TopMost = $true
    $form.TopMost = $false
    $form.Activate()
}

function Update-Display {
    $form = $script:MainForm
    if ($null -eq $form -or $form.IsDisposed) { return }
    $ui = $script:Ui
    $c = $script:Colors
    $s = $script:State
    $now = Get-Date
    $summary = Get-WorkSummary
    $active = $summary.ActiveSeconds

    $ui.Clock.Text = $now.ToString('HH:mm:ss', $script:Culture)
    $ui.Date.Text = $now.ToString('D', [Globalization.CultureInfo]::CurrentCulture)
    $ui.Hours.Text = (Get-RoundedHours $active).ToString('0.0', [Globalization.CultureInfo]::CurrentCulture)
    $ui.HoursUnit.Left = $ui.Hours.Right
    $ui.HoursUnit.Top = $ui.Hours.Bottom - $ui.HoursUnit.Height - (Px 14)

    $counting = $s.TrackingStarted -and -not $s.Ended -and -not $s.Paused -and $null -eq $s.CurrentIdleStart
    if ($s.Ended) { $status = 'Workday complete'; $dot = $c.Accent; $caption = "Ended at $(Format-Time $s.EndedAt 'HH:mm')" }
    elseif ($s.Paused) { $status = 'Paused'; $dot = $c.Warning; $caption = 'Timer paused - resume to keep counting' }
    elseif (-not $s.TrackingStarted) { $status = 'Waiting for activity'; $dot = $c.Neutral; $caption = 'Starts on your first keyboard or mouse input' }
    elseif (-not $counting) { $status = 'Idle - not counting'; $dot = $c.Warning; $caption = 'Counting resumes when you return' }
    else { $status = 'Tracking'; $dot = $c.Success; $caption = 'Next 0.1 hour in' }
    $ui.Status.Text = $status
    $ui.StatusDot.ForeColor = $dot
    $ui.CountdownCaption.Text = $caption

    $remaining = Get-SecondsToNextTenth $active
    $progress = [math]::Min([double]1, [math]::Max([double]0, 1 - ($remaining / $script:TenthHourSeconds)))
    $ui.TrackFill.Width = [int]($ui.Track.Width * $(if ($s.TrackingStarted) { $progress } else { 0 }))
    $ui.TrackFill.BackColor = if ($counting) { $c.Accent } else { $c.Neutral }
    $ui.CountdownValue.Text = if ($counting) {
        $span = [TimeSpan]::FromSeconds([math]::Ceiling($remaining))
        '{0}:{1:00}' -f [int][math]::Floor($span.TotalMinutes), $span.Seconds
    } else { '' }
    $ui.Exact.Text = if ($s.TrackingStarted) { 'Exact: {0}' -f (Format-Duration $active) } else { '' }

    $ui.Started.Text = if ($s.TrackingStarted) { Format-Time $s.SessionStart 'HH:mm' } else { '-' }
    $ui.Inactive.Text = if ($s.TrackingStarted) { Format-Duration $summary.InactiveSeconds } else { '-' }
    $ui.Presence.Text = if ($script:IsPresent) { 'Active' } else { 'Away {0}' -f (Format-Duration ($script:IdleMilliseconds / 1000)) }
    $ui.Presence.ForeColor = if ($script:IsPresent) { $c.Success } else { $c.Warning }

    $open = $s.TrackingStarted -and -not $s.Ended
    $ui.EndButton.Enabled = $open
    $ui.PauseButton.Enabled = $open
    $ui.PauseButton.Text = if ($s.Paused) { 'Resume Timer' } else { 'Pause Timer' }
    $ui.IdleButton.Enabled = [bool]$s.TrackingStarted

    $pending = @($script:History | Where-Object { $_.NeedsReview }).Count
    $ui.Footer.Cursor = if ($pending -gt 0) { [Windows.Forms.Cursors]::Hand } else { [Windows.Forms.Cursors]::Default }
    if ($null -ne $script:LastTickError) {
        $ui.Footer.Text = "Problem: $($script:LastTickError)"; $ui.Footer.ForeColor = $c.Danger
    }
    elseif ($pending -gt 0) {
        $ui.Footer.Text = 'A previous workday was never ended - click here to enter when you stopped.'; $ui.Footer.ForeColor = $c.Warning
    }
    elseif ($script:ExcelPending) {
        $ui.Footer.Text = 'Work History.xlsx is open elsewhere - it will update when closed.'; $ui.Footer.ForeColor = $c.Warning
    }
    else {
        $ui.Footer.Text = 'History is saved to Work History.xlsx'; $ui.Footer.ForeColor = $c.Muted
    }
}

function Invoke-Tick {
    if ($null -ne $script:ActivateEvent -and $script:ActivateEvent.WaitOne(0)) { Show-MainWindow }
    [void](Invoke-DayRollover)

    $s = $script:State
    $tick = [PresenceWorkTimer.V4.Native]::GetLastInputTick()
    $sessionEvent = [PresenceWorkTimer.V4.SessionWatcher]::Consume()
    if (-not $s.TrackingStarted -and -not $s.Ended) {
        if ($sessionEvent) { Start-Tracking 'Windows sign-in or unlock' }
        elseif ($script:LastInputTick -ne 0 -and $tick -ne 0 -and $tick -ne $script:LastInputTick) { Start-Tracking 'First keyboard or mouse input' }
    }
    if ($tick -ne 0) { $script:LastInputTick = $tick }

    Update-Sample
    $now = Get-Date
    if (($now - $script:LastSavedAt).TotalSeconds -ge $script:SaveIntervalSeconds) { Save-State }
    if ($script:ExcelPending -and ($now - $script:LastExportAttempt).TotalSeconds -ge $script:ExportRetrySeconds) { Export-HistoryWorkbook }
    $script:LastTickError = $null
    # Ask about an unfinished previous day once you're back at the computer (first input of the new day).
    if (-not $script:ReviewPromptDone -and -not $script:ReviewDialogOpen -and $script:State.TrackingStarted -and
        @($script:History | Where-Object { $_.NeedsReview }).Count -gt 0) {
        $script:ReviewPromptDue = $true
    }
    Update-Display
}

function Request-EndWorkday {
    $s = $script:State
    if (-not $s.TrackingStarted -or $s.Ended) { return }
    Update-Sample
    $hours = (Get-RoundedHours (Get-WorkSummary).ActiveSeconds).ToString('0.0', [Globalization.CultureInfo]::CurrentCulture)
    $confirmed = Show-Choice -Title 'End Workday' -Heading 'End today''s workday?' `
        -Message "$hours hours of active work will be saved to your work history. You can still edit idle periods today, and the history will update." `
        -PrimaryText 'End Workday' -SecondaryText 'Cancel'
    if (-not $confirmed) { return }
    Complete-Workday -Reason 'Ended by user' -At (Get-Date)
    Update-Display
}

# ---------------------------------------------------------------------------
# MARK: Idle & inactive period editor
# ---------------------------------------------------------------------------

function Show-PeriodEditor {
    if (-not $script:State.TrackingStarted) { return }
    if (-not $script:State.Ended) { Update-Sample }
    $day = [DateTime]::ParseExact($script:State.DateKey, 'yyyy-MM-dd', $script:Culture)
    $typeKeys = @($script:PeriodTypes.Keys)

    $form = New-DialogForm -Title 'Idle Periods' -Width 660 -Height 500
    $heading = New-Label 'Idle & Inactive Periods' 20 16 $script:Fonts.Heading
    $intro = New-Label 'Time inside these periods is not counted as work. Delete a period to count it, or select one to change its times.' 21 46 $script:Fonts.Small $script:Colors.Muted 620 20
    $grid = New-Grid 20 76 620 230 -HiddenColumns @('Id')
    $typeCaption = New-Label 'TYPE' 20 322 $script:Fonts.Caption $script:Colors.Muted
    $startCaption = New-Label 'START' 206 322 $script:Fonts.Caption $script:Colors.Muted
    $endCaption = New-Label 'END' 350 322 $script:Fonts.Caption $script:Colors.Muted
    $typeBox = [Windows.Forms.ComboBox]::new()
    $typeBox.DropDownStyle = [Windows.Forms.ComboBoxStyle]::DropDownList
    $typeBox.Bounds = New-Rect 20 342 172 28
    foreach ($key in $typeKeys) { [void]$typeBox.Items.Add($script:PeriodTypes[$key]) }
    $typeBox.SelectedIndex = 0
    $startPicker = New-TimePicker 206 342
    $endPicker = New-TimePicker 350 342
    $add = New-Button 'Add Period' 20 390 140 38 'Primary'
    $save = New-Button 'Save Changes' 168 390 140 38
    $delete = New-Button 'Delete' 316 390 110 38 'Danger'
    $summary = New-Label '' 20 458 $script:Fonts.Body $script:Colors.Muted 420 22
    $close = New-Button 'Close' 540 450 100 36
    $close.DialogResult = [Windows.Forms.DialogResult]::OK
    $form.AcceptButton = $close
    $form.CancelButton = $close
    $separator = [Windows.Forms.Panel]::new(); $separator.Bounds = New-Rect 0 440 660 1; $separator.BackColor = $script:Colors.Border
    $form.Controls.AddRange(@($heading, $intro, $grid, $typeCaption, $startCaption, $endCaption, $typeBox, $startPicker, $endPicker, $add, $save, $delete, $separator, $summary, $close))

    $setDefaults = {
        $limit = if ($script:State.Ended) { $script:State.EndedAt } else { Get-Date }
        $start = $limit.AddMinutes(-30)
        if ($start -lt $script:State.SessionStart) { $start = $script:State.SessionStart }
        $startPicker.Value = $start
        $endPicker.Value = $limit
        $typeBox.SelectedIndex = 0
    }

    $getSelected = {
        if ($grid.SelectedRows.Count -eq 0) { return $null }
        $id = [string]$grid.SelectedRows[0].Cells['Id'].Value
        $script:State.Periods | Where-Object { $_.Id -eq $id } | Select-Object -First 1
    }

    $refresh = {
        $table = [Data.DataTable]::new()
        foreach ($name in @('Id', 'Type', 'Start', 'End', 'Duration')) { [void]$table.Columns.Add($name, [string]) }
        foreach ($p in @($script:State.Periods | Sort-Object Start)) {
            [void]$table.Rows.Add($p.Id, $script:PeriodTypes[$p.Type], (Format-Time $p.Start), (Format-Time $p.End), (Format-Duration ($p.End - $p.Start).TotalSeconds))
        }
        $now = Get-Date
        if ($null -ne $script:State.CurrentIdleStart) {
            [void]$table.Rows.Add('', ('Idle - no input for {0}+ min (in progress)' -f $script:Settings.IdleThresholdMinutes),(Format-Time $script:State.CurrentIdleStart), 'now', (Format-Duration ($now - $script:State.CurrentIdleStart).TotalSeconds))
        }
        if ($script:State.Paused -and $null -ne $script:State.PauseStart) {
            [void]$table.Rows.Add('', 'Paused (in progress)', (Format-Time $script:State.PauseStart), 'now', (Format-Duration ($now - $script:State.PauseStart).TotalSeconds))
        }
        $grid.DataSource = $table
        $grid.ClearSelection()
        $result = Get-WorkSummary
        $summary.Text = 'Active {0} h  |  Inactive {1}' -f (Get-RoundedHours $result.ActiveSeconds).ToString('0.0'), (Format-Duration $result.InactiveSeconds)
        $save.Enabled = $false
        $delete.Enabled = $false
        & $setDefaults
    }

    $commit = {
        Save-State
        if ($script:State.Ended) { Save-HistoryRecord }
        & $refresh
        Update-Display
    }

    $readInput = {
        $start = $day.Add($startPicker.Value.TimeOfDay)
        $end = $day.Add($endPicker.Value.TimeOfDay)
        $limit = if ($script:State.Ended) { $script:State.EndedAt } else { Get-Date }
        $problem = if ($end -le $start) { 'The end time must be after the start time.' }
                   elseif ($start -lt $script:State.SessionStart) { "The period can't start before the workday started ($(Format-Time $script:State.SessionStart))." }
                   elseif ($end -gt $limit) { "The period can't end after $(Format-Time $limit)." }
                   else { $null }
        if ($null -ne $problem) {
            [void](Show-Choice -Heading 'Check the period times' -Message $problem -Kind Warning)
            return $null
        }
        [pscustomobject]@{ Type = $typeKeys[$typeBox.SelectedIndex]; Start = $start; End = $end }
    }

    $grid.Add_SelectionChanged({
        $period = & $getSelected
        $save.Enabled = $null -ne $period
        $delete.Enabled = $null -ne $period
        if ($null -ne $period) {
            $typeBox.SelectedIndex = [array]::IndexOf($typeKeys, $period.Type)
            $startPicker.Value = $period.Start
            $endPicker.Value = $period.End
        }
    })
    $add.Add_Click({
        Invoke-UiAction 'Add period' {
            $value = & $readInput
            if ($null -eq $value) { return }
            Add-Period -Type $value.Type -Start $value.Start -End $value.End
            & $commit
        }
    })
    $save.Add_Click({
        Invoke-UiAction 'Save period' {
            $period = & $getSelected
            if ($null -eq $period) { return }
            $value = & $readInput
            if ($null -eq $value) { return }
            $period.Type = $value.Type; $period.Start = $value.Start; $period.End = $value.End
            & $commit
        }
    })
    $delete.Add_Click({
        Invoke-UiAction 'Delete period' {
            $period = & $getSelected
            if ($null -eq $period) { return }
            $message = '{0}, {1} to {2}. This time will count as active work.' -f $script:PeriodTypes[$period.Type], (Format-Time $period.Start), (Format-Time $period.End)
            if (-not (Show-Choice -Heading 'Delete this period?' -Message $message -PrimaryText 'Delete' -SecondaryText 'Cancel' -Kind Danger)) { return }
            [void]$script:State.Periods.Remove($period)
            & $commit
        }
    })

    $form.Add_Shown({ & $refresh })
    & $setDefaults
    try { [void](Show-Modal $form) } finally { $form.Dispose() }
}

# ---------------------------------------------------------------------------
# MARK: Work history viewer
# ---------------------------------------------------------------------------

function Show-WorkHistory {
    $form = New-DialogForm -Title 'Work History' -Width 960 -Height 580 -Resizable
    $heading = New-Label 'Work History' 20 16 $script:Fonts.Heading
    $sub = New-Label $script:WorkbookPath 21 46 $script:Fonts.Small $script:Colors.Muted
    $filter = [Windows.Forms.TextBox]::new()
    $filter.Bounds = New-Rect 20 76 300 28
    $filter.PlaceholderText = 'Filter by date, weekday or note'
    $grid = New-Grid 20 114 920 380 -HiddenColumns @('Id')
    $grid.Anchor = 'Top, Bottom, Left, Right'
    $summary = New-Label '' 20 520 $script:Fonts.Body $script:Colors.Muted 400 22
    $summary.Anchor = 'Bottom, Left'
    $editDay = New-Button 'Edit Day' 552 512 120 38
    $editDay.Anchor = 'Bottom, Right'
    $editDay.Enabled = $false
    $openExcel = New-Button 'Open in Excel' 680 512 150 38 'Primary'
    $openExcel.Anchor = 'Bottom, Right'
    $close = New-Button 'Close' 840 512 100 38
    $close.Anchor = 'Bottom, Right'
    $close.DialogResult = [Windows.Forms.DialogResult]::OK
    $form.CancelButton = $close
    $form.Controls.AddRange(@($heading, $sub, $filter, $grid, $summary, $editDay, $openExcel, $close))

    $grid.Add_DataBindingComplete({
        param($source, $e)
        foreach ($name in @('Hours worked', 'Inactive (h)')) {
            if ($source.Columns.Contains($name)) {
                $source.Columns[$name].DefaultCellStyle.Format = if ($name -eq 'Hours worked') { 'N1' } else { 'N2' }
                $source.Columns[$name].DefaultCellStyle.Alignment = [Windows.Forms.DataGridViewContentAlignment]::MiddleRight
            }
        }
        if ($source.Columns.Contains('Inactive periods')) { $source.Columns['Inactive periods'].FillWeight = 220 }
        if ($source.Columns.Contains('Completion')) { $source.Columns['Completion'].FillWeight = 160 }
        foreach ($gridRow in $source.Rows) {
            if ([string]$gridRow.Cells['Completion'].Value -like 'Needs end time*') {
                $gridRow.Cells['Completion'].Style.ForeColor = $script:Colors.Warning
                $gridRow.Cells['Completion'].Style.Font = $script:Fonts.Strong
            }
        }
        $source.ClearSelection()
    })

    $getSelected = {
        if ($grid.SelectedRows.Count -eq 0) { return $null }
        Find-HistoryRecord ([string]$grid.SelectedRows[0].Cells['Id'].Value)
    }
    $editSelected = {
        $record = & $getSelected
        if ($null -eq $record) { return }
        if ((Show-EditDayRecord -Record $record) -eq 'Saved') { & $refresh; Update-Display }
    }

    $refresh = {
        $table = [Data.DataTable]::new()
        foreach ($name in @('Id', 'Date', 'Day', 'Start', 'End')) { [void]$table.Columns.Add($name, [string]) }
        [void]$table.Columns.Add('Hours worked', [double])
        [void]$table.Columns.Add('Inactive (h)', [double])
        foreach ($name in @('Inactive periods', 'Completion')) { [void]$table.Columns.Add($name, [string]) }

        $term = $filter.Text.Trim()
        $pattern = '*' + [Management.Automation.WildcardPattern]::Escape($term) + '*'
        $total = [double]0
        $count = 0
        $records = @($script:History | Sort-Object -Property @{ Expression = { $_.Date } }, @{ Expression = { if ($null -eq $_.SessionStart) { [DateTime]::MinValue } else { $_.SessionStart } } } -Descending)
        foreach ($r in $records) {
            $date = [DateTime]::MinValue
            $dayName = if ([DateTime]::TryParseExact($r.Date, 'yyyy-MM-dd', $script:Culture, [Globalization.DateTimeStyles]::None, [ref]$date)) { $date.ToString('ddd', [Globalization.CultureInfo]::CurrentCulture) } else { '' }
            if ($term -and -not (($r.Date -like $pattern) -or ($dayName -like $pattern) -or ($r.Reason -like $pattern) -or ($r.InactivePeriods -like $pattern))) { continue }
            [void]$table.Rows.Add($r.Id, $r.Date, $dayName, (Format-Time $r.SessionStart 'HH:mm'), (Format-Time $r.SessionEnd 'HH:mm'),
                [double]$r.ActiveHours, [math]::Round([double]$r.InactiveSeconds / 3600, 2), $r.InactivePeriods,
                $(if ($r.NeedsReview) { "Needs end time - $($r.Reason)" } else { $r.Reason }))
            $total += [double]$r.ActiveHours
            $count++
        }
        $grid.DataSource = $table
        $summary.Text = if ($count -eq 0) { 'No records' } else {
            '{0} day(s)  |  {1:N1} h total  |  {2:N1} h average' -f $count, $total, ($total / $count)
        }
    }

    $filter.Add_TextChanged({ Invoke-UiAction 'Filter history' { & $refresh } })
    $grid.Add_SelectionChanged({ $editDay.Enabled = $grid.SelectedRows.Count -gt 0 })
    $grid.Add_CellDoubleClick({ param($source, $e) if ($e.RowIndex -ge 0) { Invoke-UiAction 'Edit day' { & $editSelected } } })
    $editDay.Add_Click({ Invoke-UiAction 'Edit day' { & $editSelected } })
    $openExcel.Add_Click({
        Invoke-UiAction 'Open in Excel' {
            if (-not (Test-Path -LiteralPath $script:WorkbookPath -PathType Leaf)) { Export-HistoryWorkbook }
            Invoke-Item -LiteralPath $script:WorkbookPath
        }
    })
    & $refresh
    try { [void](Show-Modal $form) } finally { $form.Dispose() }
}

# ---------------------------------------------------------------------------
# MARK: Correct a past day: enter the real end time or the hours worked
# ---------------------------------------------------------------------------

function Show-EditDayRecord {
    # Returns 'Saved', 'Kept' (review accepted as recorded), 'Later' or 'Cancel'.
    param([Parameter(Mandatory)][object]$Record, [switch]$Review)
    $c = $script:Colors
    $f = $script:Fonts
    $culture = [Globalization.CultureInfo]::CurrentCulture
    $day = [DateTime]::MinValue
    if (-not [DateTime]::TryParseExact($Record.Date, 'yyyy-MM-dd', $script:Culture, [Globalization.DateTimeStyles]::None, [ref]$day)) {
        throw "The record date '$($Record.Date)' is not valid."
    }
    $dayEnd = $day.AddDays(1).AddSeconds(-1)
    $recordedStart = if ($null -ne $Record.SessionStart) { [DateTime]$Record.SessionStart } else { $day.AddHours(8) }
    $recordedEnd = if ($null -ne $Record.SessionEnd) { [DateTime]$Record.SessionEnd } else { $recordedStart.AddSeconds([double]$Record.ActiveSeconds) }
    $hoursText = ([double]$Record.ActiveHours).ToString('0.0', $culture)

    $form = New-DialogForm -Title $(if ($Review) { 'Finish Previous Workday' } else { 'Edit Day' }) -Width 460 -Height 394
    $headingText = if ($Review) { "$($day.ToString('dddd', $culture)) wasn't ended" } else { $day.ToString('D', $culture) }
    $heading = New-Label $headingText 24 18 $f.Heading
    $infoText = if ($Review) {
        "Work Timer kept counting until $(Format-Time $recordedEnd 'HH:mm') on $($day.ToString('d', $culture)) and recorded $hoursText hours. Enter when you actually stopped, or how many hours you worked."
    } else {
        "Recorded $(Format-Time $recordedStart 'HH:mm') to $(Format-Time $recordedEnd 'HH:mm'), $hoursText hours. Change the start, and either the end time or the hours worked."
    }
    $info = New-Label $infoText 24 52 $f.Body $c.Text
    $info.MaximumSize = [Drawing.Size]::new((Px 412), 0)

    $startCaption = New-Label 'START TIME' 24 130 $f.Caption $c.Muted
    $startPicker = New-TimePicker 24 150 130
    $startPicker.Value = $recordedStart
    $modeCaption = New-Label 'WORKDAY ENDED' 24 194 $f.Caption $c.Muted
    $byEnd = [Windows.Forms.RadioButton]::new(); $byEnd.Text = 'At this time'; $byEnd.Font = $f.Body; $byEnd.Bounds = New-Rect 24 216 130 26; $byEnd.Checked = $true
    $endPicker = New-TimePicker 160 214 130
    $endPicker.Value = $recordedEnd
    $byHours = [Windows.Forms.RadioButton]::new(); $byHours.Text = 'After working'; $byHours.Font = $f.Body; $byHours.Bounds = New-Rect 24 254 130 26
    $hoursBox = [Windows.Forms.NumericUpDown]::new()
    $hoursBox.Bounds = New-Rect 160 252 90 28
    $hoursBox.DecimalPlaces = 2; $hoursBox.Increment = [decimal]0.1; $hoursBox.Minimum = 0; $hoursBox.Maximum = 24
    $hoursBox.Value = [decimal][math]::Min(24, [math]::Round([double]$Record.ActiveHours, 2))
    $hoursBox.Enabled = $false
    $hoursSuffix = New-Label 'hours' 256 256 $f.Body $c.Text
    $preview = New-Label '' 24 298 $f.Strong $c.Accent 412 22

    $footer = [Windows.Forms.Panel]::new(); $footer.Dock = 'Bottom'; $footer.Height = Px 64; $footer.BackColor = $c.Subtle
    $save = New-Button 'Save' 330 14 110 36 'Primary'
    $save.DialogResult = [Windows.Forms.DialogResult]::OK
    $footer.Controls.Add($save)
    if ($Review) {
        $keep = New-Button 'Keep as Recorded' 178 14 144 36
        $keep.DialogResult = [Windows.Forms.DialogResult]::Ignore
        $later = New-Button 'Later' 20 14 90 36
        $later.DialogResult = [Windows.Forms.DialogResult]::Cancel
        $footer.Controls.AddRange(@($keep, $later))
        $form.CancelButton = $later
    }
    else {
        $cancel = New-Button 'Cancel' 220 14 102 36
        $cancel.DialogResult = [Windows.Forms.DialogResult]::Cancel
        $footer.Controls.Add($cancel)
        $form.CancelButton = $cancel
    }
    $form.AcceptButton = $save
    $form.Controls.AddRange(@($heading, $info, $startCaption, $startPicker, $modeCaption, $byEnd, $endPicker, $byHours, $hoursBox, $hoursSuffix, $preview, $footer))

    # Works out the new start/end from the inputs. Returns $null and sets $preview on a problem.
    $calculate = {
        $start = $day.Add($startPicker.Value.TimeOfDay)
        if ($byEnd.Checked) {
            $end = $day.Add($endPicker.Value.TimeOfDay)
            if ($end -le $start) { $preview.Text = 'The end time must be after the start time.'; $preview.ForeColor = $c.Danger; return $null }
        }
        else {
            $end = Get-EndTimeForActiveSeconds -Record $Record -Start $start -TargetSeconds ([double]$hoursBox.Value * 3600)
            if ($end -gt $dayEnd) { $preview.Text = "That many hours doesn't fit between the start time and midnight."; $preview.ForeColor = $c.Danger; return $null }
        }
        $active = Measure-RecordActiveSeconds -Record $Record -Start $start -End $end
        $preview.ForeColor = $c.Accent
        $preview.Text = '{0} to {1}  |  {2} hours worked' -f (Format-Time $start 'HH:mm'), (Format-Time $end 'HH:mm'), (Get-RoundedHours $active).ToString('0.0', $culture)
        [pscustomobject]@{ Start = $start; End = $end }
    }
    $update = {
        $endPicker.Enabled = $byEnd.Checked
        $hoursBox.Enabled = $byHours.Checked
        $save.Enabled = $null -ne (& $calculate)
    }
    foreach ($control in @($startPicker, $endPicker)) { $control.Add_ValueChanged({ & $update }) }
    $hoursBox.Add_ValueChanged({ & $update })
    $byEnd.Add_CheckedChanged({ & $update })
    & $update

    try {
        $result = Show-Modal $form
        $values = & $calculate
        $useHours = $byHours.Checked
    }
    finally { $form.Dispose() }

    switch ($result) {
        ([Windows.Forms.DialogResult]::OK) {
            if ($null -eq $values) { return 'Cancel' }
            $base = ([string]$Record.Reason -replace '; corrected .*$', '')
            $how = if ($useHours) { 'hours worked' } else { 'end time' }
            Set-HistoryRecordTimes -Record $Record -Start $values.Start -End $values.End -Reason "$base; corrected ($how entered)"
            return 'Saved'
        }
        ([Windows.Forms.DialogResult]::Ignore) {
            $Record.NeedsReview = $false
            Save-History
            Export-HistoryWorkbook
            return 'Kept'
        }
        ([Windows.Forms.DialogResult]::Cancel) { if ($Review) { return 'Later' } return 'Cancel' }
    }
    return 'Cancel'
}

function Show-PendingReviews {
    if ($script:ReviewDialogOpen) { return }
    $script:ReviewDialogOpen = $true
    try {
        foreach ($record in @(Get-PendingReviews)) {
            if ((Show-EditDayRecord -Record $record -Review) -eq 'Later') { break }
        }
    }
    finally {
        $script:ReviewDialogOpen = $false
        $script:ReviewPromptDone = $true
        Update-Display
    }
}

# ---------------------------------------------------------------------------
# MARK: Advanced tools (Ctrl+Shift+A)
# ---------------------------------------------------------------------------

function Show-AdvancedTools {
    $c = $script:Colors
    $f = $script:Fonts
    $form = New-DialogForm -Title 'Advanced Tools' -Width 480 -Height 480
    $heading = New-Label 'Advanced Tools' 24 16 $f.Heading
    $sub = New-Label 'Opened with Ctrl+Shift+A' 25 46 $f.Small $c.Muted

    $workdayCaption = New-Label 'WORKDAY' 24 82 $f.Caption $c.Muted
    $editStart = New-Button 'Edit Start Time' 24 104 138 38
    $reopen = New-Button 'Reopen Workday' 170 104 138 38
    $reset = New-Button 'Reset Today' 316 104 138 38 'Danger'

    $startupCaption = New-Label 'START WITH WINDOWS' 24 166 $f.Caption $c.Muted
    $startupStatus = New-Label '' 24 186 $f.Body $c.Text
    $startupToggle = New-Button '' 24 214 210 38 'Primary'
    $startMenu = New-Button 'Add Start Menu Shortcut' 242 214 212 38

    $idleCaption = New-Label 'IDLE DETECTION' 24 276 $f.Caption $c.Muted
    $idleText = New-Label 'Mark me idle after' 24 304 $f.Body $c.Text
    $idleMinutes = [Windows.Forms.NumericUpDown]::new()
    $idleMinutes.Bounds = New-Rect 150 300 70 28
    $idleMinutes.Minimum = 1; $idleMinutes.Maximum = 1440
    $idleMinutes.Value = [decimal]$script:Settings.IdleThresholdMinutes
    $idleSuffix = New-Label 'minutes without input' 228 304 $f.Body $c.Text
    $idleApply = New-Button 'Apply' 384 298 70 32

    $dataCaption = New-Label 'DATA' 24 352 $f.Caption $c.Muted
    $openFolder = New-Button 'Open Data Folder' 24 374 138 38
    $rebuild = New-Button 'Rebuild Excel Log' 170 374 138 38
    $openLog = New-Button 'View Error Log' 316 374 138 38

    $separator = [Windows.Forms.Panel]::new(); $separator.Bounds = New-Rect 0 428 480 1; $separator.BackColor = $c.Border
    $close = New-Button 'Close' 354 436 100 36
    $close.DialogResult = [Windows.Forms.DialogResult]::OK
    $form.CancelButton = $close
    $form.Controls.AddRange(@($heading, $sub, $workdayCaption, $editStart, $reopen, $reset, $startupCaption, $startupStatus,
        $startupToggle, $startMenu, $idleCaption, $idleText, $idleMinutes, $idleSuffix, $idleApply, $dataCaption,
        $openFolder, $rebuild, $openLog, $separator, $close))

    $refresh = {
        $enabled = Test-StartWithWindows
        $startupStatus.Text = if ($enabled) { 'Work Timer starts automatically when you sign in.' } else { 'Work Timer does not start automatically.' }
        $startupToggle.Text = if ($enabled) { 'Stop Starting with Windows' } else { 'Start with Windows' }
        $reopen.Enabled = [bool]$script:State.Ended
    }

    $editStart.Add_Click({ Invoke-UiAction 'Edit start time' { Show-EditStartTime; & $refresh } })
    $reopen.Add_Click({ Invoke-UiAction 'Reopen workday' { Invoke-ReopenWorkday; & $refresh } })
    $reset.Add_Click({ Invoke-UiAction 'Reset today' { Invoke-ResetToday; & $refresh } })
    $startupToggle.Add_Click({
        Invoke-UiAction 'Start with Windows' {
            Set-StartWithWindows -Enabled (-not (Test-StartWithWindows))
            & $refresh
        }
    })
    $startMenu.Add_Click({
        Invoke-UiAction 'Start menu shortcut' {
            New-AppShortcut -Path (Get-StartMenuShortcutPath)
            [void](Show-Choice -Heading 'Shortcut added' -Message 'Search for "Work Timer" in the Start menu. It opens the timer without a PowerShell window, and you can pin it to the taskbar.')
        }
    })
    $idleApply.Add_Click({
        Invoke-UiAction 'Idle threshold' {
            $script:Settings.IdleThresholdMinutes = [int]$idleMinutes.Value
            Save-Settings
            [void](Show-Choice -Heading 'Idle detection updated' -Message "You'll be marked idle after $([int]$idleMinutes.Value) minute(s) without keyboard or mouse input. Idle time starts from your last input.")
        }
    })
    $openFolder.Add_Click({ Invoke-UiAction 'Open data folder' { Initialize-Storage; Invoke-Item -LiteralPath $script:AppDirectory } })
    $rebuild.Add_Click({
        Invoke-UiAction 'Rebuild Excel log' {
            Export-HistoryWorkbook
            if ($script:ExcelPending) { [void](Show-Choice -Heading 'Workbook is in use' -Message 'Close Work History.xlsx in Excel. It will update automatically.' -Kind Warning) }
            else { [void](Show-Choice -Heading 'Excel log rebuilt' -Message $script:WorkbookPath) }
        }
    })
    $openLog.Add_Click({
        Invoke-UiAction 'Open error log' {
            if (Test-Path -LiteralPath $script:ErrorLogPath) { Invoke-Item -LiteralPath $script:ErrorLogPath }
            else { [void](Show-Choice -Heading 'No errors logged' -Message 'The error log is empty.') }
        }
    })

    & $refresh
    try { [void](Show-Modal $form) } finally { $form.Dispose() }
    Update-Display
}

function Show-EditStartTime {
    $s = $script:State
    $form = New-DialogForm -Title 'Edit Start Time' -Width 400 -Height 236
    $heading = New-Label 'When did your workday start?' 24 18 $script:Fonts.Heading
    $note = New-Label "Today, $((Get-Date).ToString('D', [Globalization.CultureInfo]::CurrentCulture)). Time before the start is not counted; inactive periods are kept." 25 50 $script:Fonts.Small $script:Colors.Muted
    $note.MaximumSize = [Drawing.Size]::new((Px 350), 0)
    $picker = New-TimePicker 24 102 150
    $picker.Value = if ($null -ne $s.SessionStart) { $s.SessionStart } else { Get-Date }
    $ok = New-Button 'Apply' 166 176 110 38 'Primary'; $ok.DialogResult = [Windows.Forms.DialogResult]::OK
    $cancel = New-Button 'Cancel' 284 176 92 38; $cancel.DialogResult = [Windows.Forms.DialogResult]::Cancel
    $form.Controls.AddRange(@($heading, $note, $picker, $ok, $cancel))
    $form.AcceptButton = $ok
    $form.CancelButton = $cancel
    try {
        if ((Show-Modal $form) -ne [Windows.Forms.DialogResult]::OK) { return }
        $newStart = (Get-Date).Date.Add($picker.Value.TimeOfDay)
    }
    finally { $form.Dispose() }

    $limit = if ($s.Ended) { $s.EndedAt } else { Get-Date }
    if ($newStart -gt $limit) {
        [void](Show-Choice -Heading 'Start time not changed' -Message "The start time can't be later than $(Format-Time $limit)." -Kind Warning)
        return
    }
    if (-not $s.TrackingStarted) {
        Start-Tracking -Reason 'Start time set manually' -At $newStart
    }
    else {
        $s.SessionStart = $newStart
        Save-State
        if ($s.Ended) { Save-HistoryRecord }
    }
    Update-Display
}

function Invoke-ReopenWorkday {
    $s = $script:State
    if (-not $s.Ended) { return }
    $message = "You ended the workday at $(Format-Time $s.EndedAt 'HH:mm'). The time since then will be recorded as an inactive period (delete it in Idle Periods to count it). Today's history entry updates when you end the day again."
    if (-not (Show-Choice -Heading 'Reopen today''s workday?' -Message $message -PrimaryText 'Reopen' -SecondaryText 'Cancel')) { return }
    $now = Get-Date
    Add-Period -Type Stopped -Start $s.EndedAt -End $now
    $s.Ended = $false
    $s.EndedAt = $null
    $s.EndReason = $null
    $s.LastSampleTime = $now
    Save-State
    Update-Display
}

function Invoke-ResetToday {
    $message = "Today's start time, idle periods and active time will be cleared and the timer will wait for your next input. Entries already saved to your work history are kept."
    if (-not (Show-Choice -Heading 'Reset today?' -Message $message -PrimaryText 'Reset' -SecondaryText 'Cancel' -Kind Danger)) { return }
    $script:State = New-DayState
    $script:LastInputTick = [PresenceWorkTimer.V4.Native]::GetLastInputTick()
    Save-State
    Update-Display
}

# ---------------------------------------------------------------------------
# MARK: Startup
# ---------------------------------------------------------------------------

function Show-FatalStartupError {
    param([Parameter(Mandatory)][object]$ErrorObject)
    $message = if ($ErrorObject -is [Management.Automation.ErrorRecord]) { $ErrorObject.Exception.Message } else { [string]$ErrorObject }
    if ($ErrorObject -is [Management.Automation.ErrorRecord]) { Write-AppLog 'Fatal error.' $ErrorObject } else { Write-AppLog "Fatal error: $message" }
    $path = if ($null -eq $script:ScriptPath) { '(unsaved script)' } else { $script:ScriptPath }
    [void][Windows.Forms.MessageBox]::Show("Work Timer stopped because of an error.`n`nScript: $path`n`n$message`n`nDetails: $($script:ErrorLogPath)", 'Work Timer Error', 'OK', 'Error')
}

try {
    [void][PresenceWorkTimer.V4.Native]::HideOwnConsole()
    Initialize-Storage
    if (-not (Initialize-SingleInstance)) { return }

    try { [void][PresenceWorkTimer.V4.AppIdentity]::SetProcessAppId($script:AppUserModelId) } catch { }
    [Windows.Forms.Application]::EnableVisualStyles()
    try { [Windows.Forms.Application]::SetCompatibleTextRenderingDefault($false) } catch { }
    try { [void][Windows.Forms.Application]::SetHighDpiMode([Windows.Forms.HighDpiMode]::SystemAware) } catch { }
    [Windows.Forms.Application]::SetUnhandledExceptionMode([Windows.Forms.UnhandledExceptionMode]::CatchException)
    [Windows.Forms.Application]::add_ThreadException({
        param($source, $e)
        Write-AppLog "Unhandled UI error: $($e.Exception.Message)"
    })

    Initialize-Theme
    Update-ExistingShortcuts
    Read-Settings
    Read-History
    Read-State
    [void](Invoke-DayRollover -Reason 'Automatic rollover at startup')
    Save-State

    Initialize-MainWindow
    $script:Timer = [Windows.Forms.Timer]::new()
    $script:Timer.Interval = 1000
    $script:Timer.Add_Tick({
        if ($script:ShuttingDown) { return }
        try {
            Invoke-Tick
            if ($script:ReviewPromptDue) {
                # Shown after the tick's work is finished so the dialog never interrupts a state update.
                $script:ReviewPromptDue = $false
                Invoke-UiAction 'Previous workday' { Show-PendingReviews }
            }
        }
        catch {
            $script:LastTickError = $_.Exception.Message
            if (((Get-Date) - $script:LastTickErrorLoggedAt).TotalMinutes -ge 1) {
                Write-AppLog 'Timer tick failed.' $_
                $script:LastTickErrorLoggedAt = Get-Date
            }
            try { Update-Display } catch { }
        }
    })

    [PresenceWorkTimer.V4.SessionWatcher]::Start()
    Update-Display
    [Windows.Forms.Application]::Run($script:MainForm)
}
catch {
    Show-FatalStartupError $_
}
finally {
    $script:ShuttingDown = $true
    try { [PresenceWorkTimer.V4.SessionWatcher]::Stop() } catch { }
    if ($null -ne $script:Timer) { $script:Timer.Stop(); $script:Timer.Dispose(); $script:Timer = $null }
    if ($null -ne $script:MainForm) { $script:MainForm.Dispose(); $script:MainForm = $null }
    Close-SingleInstance
}
