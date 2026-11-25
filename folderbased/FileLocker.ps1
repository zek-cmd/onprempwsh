#Requires -Version 5.1
<#
.SYNOPSIS
    Shows which processes are locking a file or a folder, using the Windows
    Restart Manager API (the same Windows helper that says "this file is in use").

.DESCRIPTION
    Picks one folder (through the Windows "Browse for Folder" dialog) or one file
    (through the "Open" dialog) when -Path is not given, then asks the Windows
    Restart Manager which processes hold an open handle on it.

    For a folder the script registers every file inside it, because Windows Restart
    Manager works on files and refuses a bare folder path. Use -Recurse to look
    inside subfolders as well.

    Every locking process is shown with its process ID, name, the application name
    Windows reports, the way Windows sees it (window, service, console, Explorer),
    its session, when it started, and the program file it runs from.

    -Stop ends the processes that are locking the path, one clear line in the log
    per process. Nothing is stopped unless -Stop is given on purpose.

    What the Restart Manager cannot see (this is Windows, not the script):
      - a file that only *denies access* because of permissions, not because a
        program has it open;
      - handles held by kernel drivers and some system processes;
      - locks inside a virtual machine, a container or a remote share reached
        over a protocol that hides them.

    In those cases Resource Monitor (resmon) or Sysinternals "handle.exe" can show
    the remaining holder.

.PARAMETER Path
    The file or folder to check. When left out, the picker dialog opens.

.PARAMETER File
    Force the "Open file" dialog when -Path is not given (default is the folder
    dialog; the console asks which one when neither switch is used).

.PARAMETER Recurse
    With a folder: also check every file in its subfolders.

.PARAMETER MaxFiles
    Safety stop for very large folders. Files past this number are not checked.
    Default 5000.

.PARAMETER Stop
    End the processes that hold the path. Each one is confirmed first (use -WhatIf
    to only see what would be stopped).

.EXAMPLE
    .\FileLocker.ps1

.EXAMPLE
    .\FileLocker.ps1 -Path 'D:\Shares\Finance\Report.xlsx'

.EXAMPLE
    .\FileLocker.ps1 -Path 'D:\Shares\Finance' -Recurse

.EXAMPLE
    .\FileLocker.ps1 -Path 'D:\Shares\Finance\Report.xlsx' -Stop -WhatIf
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [string]$Path,
    [switch]$File,
    [switch]$Recurse,
    [ValidateRange(1, 200000)]
    [int]$MaxFiles = 5000,
    [switch]$Stop
)

$ErrorActionPreference = 'Stop'

# ============================================
# C# HELPER  -  Restart Manager API
# ============================================

if (-not ('FileLockCheck.LockFinder' -as [type])) {
    Add-Type -TypeDefinition @'
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;

namespace FileLockCheck
{
    public class LockHolder
    {
        public int ProcessId;
        public string AppName;
        public string ServiceName;
        public string AppType;
        public uint SessionId;
        public bool Restartable;
        public DateTime StartTime;
    }

    public static class LockFinder
    {
        [StructLayout(LayoutKind.Sequential)]
        private struct RM_UNIQUE_PROCESS
        {
            public int dwProcessId;
            public System.Runtime.InteropServices.ComTypes.FILETIME ProcessStartTime;
        }

        private const int RmRebootReasonNone = 0;
        private const int CCH_RM_MAX_APP_NAME = 255;
        private const int CCH_RM_MAX_SVC_NAME = 63;
        private const int CCH_RM_SESSION_KEY = 32;
        private const int ERROR_SUCCESS = 0;
        private const int ERROR_ACCESS_DENIED = 5;
        private const int ERROR_MORE_DATA = 234;

        private enum RM_APP_TYPE
        {
            RmUnknownApp = 0,
            RmMainWindow = 1,
            RmOtherWindow = 2,
            RmService = 3,
            RmExplorer = 4,
            RmConsole = 5,
            RmCritical = 1000
        }

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        private struct RM_PROCESS_INFO
        {
            public RM_UNIQUE_PROCESS Process;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = CCH_RM_MAX_APP_NAME + 1)]
            public string strAppName;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = CCH_RM_MAX_SVC_NAME + 1)]
            public string strServiceShortName;
            public RM_APP_TYPE ApplicationType;
            public uint AppStatus;
            public uint TSSessionId;
            [MarshalAs(UnmanagedType.Bool)]
            public bool bRestartable;
        }

        [DllImport("rstrtmgr.dll", CharSet = CharSet.Unicode)]
        private static extern int RmStartSession(out uint pSessionHandle, int dwSessionFlags, StringBuilder strSessionKey);

        [DllImport("rstrtmgr.dll")]
        private static extern int RmEndSession(uint pSessionHandle);

        [DllImport("rstrtmgr.dll", CharSet = CharSet.Unicode)]
        private static extern int RmRegisterResources(uint pSessionHandle, uint nFiles, string[] rgsFilenames,
            uint nApplications, [In] RM_UNIQUE_PROCESS[] rgApplications, uint nServices, string[] rgsServiceNames);

        [DllImport("rstrtmgr.dll")]
        private static extern int RmGetList(uint dwSessionHandle, out uint pnProcInfoNeeded,
            ref uint pnProcInfo, [In, Out] RM_PROCESS_INFO[] rgAffectedApps, out uint lpdwRebootReasons);

        private static string ErrorText(int code)
        {
            if (code == ERROR_ACCESS_DENIED)
            {
                return "access denied (Windows error 5) - the path was still being registered, or the"
                    + " program holding it is a driver, a service or a protected process Windows will not name";
            }
            if (code == 87)
            {
                return "Windows error 87 (bad arguments) - Restart Manager accepts files, not folders";
            }
            return "Windows error " + code;
        }

        // Windows frequently answers ERROR_ACCESS_DENIED on the first RmGetList call made right
        // after RmRegisterResources, while the resource is still being registered. The documented
        // handling is to ask again, so this waits about a second in small steps instead of failing
        // at once. That is what made the script say "run as administrator" on an admin prompt.
        private static int GetListWithRetry(uint handle, out uint needed, ref uint count,
            RM_PROCESS_INFO[] info, out uint reasons)
        {
            int res = ERROR_SUCCESS;
            needed = 0;
            reasons = RmRebootReasonNone;

            for (int attempt = 0; attempt < 40; attempt++)
            {
                if (info != null) { count = (uint)info.Length; }
                res = RmGetList(handle, out needed, ref count, info, out reasons);

                if (res != ERROR_ACCESS_DENIED)
                {
                    return res;
                }

                Thread.Sleep(25);
            }

            return res;
        }

        public static List<LockHolder> Find(string[] resources)
        {
            List<LockHolder> holders = new List<LockHolder>();
            if (resources == null || resources.Length == 0)
            {
                return holders;
            }

            uint handle;
            StringBuilder key = new StringBuilder(CCH_RM_SESSION_KEY + 1);
            int res = RmStartSession(out handle, 0, key);
            if (res != ERROR_SUCCESS)
            {
                throw new InvalidOperationException("Could not start a Restart Manager session: " + ErrorText(res) + ".");
            }

            try
            {
                res = RmRegisterResources(handle, (uint)resources.Length, resources, 0, null, 0, null);
                if (res != ERROR_SUCCESS)
                {
                    throw new InvalidOperationException("Could not register the path with the Restart Manager: " + ErrorText(res) + ".");
                }

                // Ask once for the number of holders, then again with a buffer that size. The buffer
                // can still be one short when a program opens the file in between, so a bigger buffer
                // is taken and Windows is asked again instead of failing on ERROR_MORE_DATA.
                uint needed = 0;
                uint count = 0;
                uint reasons = RmRebootReasonNone;
                RM_PROCESS_INFO[] info = null;

                for (int round = 0; round < 8; round++)
                {
                    res = GetListWithRetry(handle, out needed, ref count, info, out reasons);

                    if (res == ERROR_SUCCESS)
                    {
                        break;
                    }

                    if (res != ERROR_MORE_DATA)
                    {
                        throw new InvalidOperationException("Could not read the list of locking processes: " + ErrorText(res) + ".");
                    }

                    if (needed == 0)
                    {
                        return holders;
                    }

                    info = new RM_PROCESS_INFO[needed];
                    count = needed;
                }

                if (res != ERROR_SUCCESS)
                {
                    throw new InvalidOperationException("Could not read the list of locking processes: " + ErrorText(res) + ".");
                }

                for (uint i = 0; i < count; i++)
                {
                    LockHolder holder = new LockHolder();
                    holder.ProcessId = info[i].Process.dwProcessId;
                    holder.AppName = (info[i].strAppName ?? string.Empty).TrimEnd('\0');
                    holder.ServiceName = (info[i].strServiceShortName ?? string.Empty).TrimEnd('\0');
                    holder.AppType = info[i].ApplicationType.ToString();
                    holder.SessionId = info[i].TSSessionId;
                    holder.Restartable = info[i].bRestartable;

                    long fileTime = ((long)info[i].Process.ProcessStartTime.dwHighDateTime << 32)
                        | (uint)info[i].Process.ProcessStartTime.dwLowDateTime;

                    try
                    {
                        holder.StartTime = DateTime.FromFileTimeUtc(fileTime).ToLocalTime();
                    }
                    catch
                    {
                        holder.StartTime = DateTime.MinValue;
                    }

                    holders.Add(holder);
                }
            }
            finally
            {
                RmEndSession(handle);
            }

            return holders;
        }
    }
}
'@ -Language CSharp
}

# ============================================
# HELPERS
# ============================================

function Write-ColorOutput {
    param(
        [string]$Message = '',
        [string]$Color = 'White'
    )
    Write-Host $Message -ForegroundColor $Color
}

function Show-Banner {
    Write-Host ''
    Write-ColorOutput '========================================' 'Cyan'
    Write-ColorOutput '   FILE LOCKER  -  WHO IS USING IT?' 'Yellow'
    Write-ColorOutput '========================================' 'Cyan'
    Write-ColorOutput '   Engine  : Windows Restart Manager' 'Gray'
    Write-ColorOutput '   Look at : files and folders in use by a program' 'Gray'
    Write-Host ''
}

function Get-ShortText {
    # Cuts a long text down so a table row stays on one line.
    param(
        [string]$Text,
        [int]$Width
    )

    if ([string]::IsNullOrEmpty($Text)) { return '' }
    if ($Width -le 1) { return '' }
    if ($Text.Length -le $Width) { return $Text }
    if ($Width -le 4) { return $Text.Substring(0, $Width) }
    return $Text.Substring(0, $Width - 3) + '...'
}

function Select-TargetPath {
    <#
        Shows the Windows folder picker or file picker and returns the chosen path,
        or $null when the user cancels. The dialog runs on a single-threaded (STA)
        thread, which is what it needs under both Windows PowerShell 5.1 and
        PowerShell 7.
    #>
    param(
        [ValidateSet('Folder', 'File')]
        [string]$Kind = 'Folder',
        [string]$Description,
        [string]$InitialDirectory
    )

    $dialogCode = {
        param($Kind, $Description, $InitialDirectory)

        Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop

        # A hidden, always-on-top owner keeps the dialog in front of this window.
        $owner = New-Object System.Windows.Forms.Form
        $owner.TopMost = $true
        $owner.ShowInTaskbar = $false
        $owner.Opacity = 0
        $owner.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen

        $dialog = $null
        $chosen = $null

        try {
            if ($Kind -eq 'File') {
                $dialog = New-Object System.Windows.Forms.OpenFileDialog
                $dialog.Title = $Description
                $dialog.CheckFileExists = $true
                $dialog.CheckPathExists = $true
                $dialog.Multiselect = $false
                $dialog.DereferenceLinks = $false
                $dialog.Filter = 'All files (*.*)|*.*'
            }
            else {
                $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
                $dialog.Description = $Description
                try { $dialog.ShowNewFolderButton = $false } catch { }
            }

            if (-not [string]::IsNullOrWhiteSpace($InitialDirectory)) {
                if ($Kind -eq 'File') {
                    if (Test-Path -LiteralPath $InitialDirectory -PathType Container) {
                        $dialog.InitialDirectory = $InitialDirectory
                    }
                }
                elseif (Test-Path -LiteralPath $InitialDirectory -PathType Container) {
                    $dialog.SelectedPath = $InitialDirectory
                }
            }

            $answer = $dialog.ShowDialog($owner)
            if ($answer -eq [System.Windows.Forms.DialogResult]::OK) {
                if ($Kind -eq 'File') { $chosen = $dialog.FileName }
                else { $chosen = $dialog.SelectedPath }
            }
        }
        finally {
            if ($null -ne $dialog) { $dialog.Dispose() }
            $owner.Dispose()
        }

        return $chosen
    }

    $apartmentState = [System.Threading.Thread]::CurrentThread.GetApartmentState()
    if ($apartmentState -eq [System.Threading.ApartmentState]::STA) {
        return (& $dialogCode $Kind $Description $InitialDirectory)
    }

    # Single-threaded apartment not available here (PowerShell 7): use a temporary
    # runspace that has one.
    $runspace = [runspacefactory]::CreateRunspace()
    $runspace.ApartmentState = [System.Threading.ApartmentState]::STA
    $runspace.ThreadOptions = 'ReuseThread'
    $runspace.Open()

    $shell = [powershell]::Create()
    $shell.Runspace = $runspace
    [void]$shell.AddScript($dialogCode.ToString()).AddArgument($Kind).AddArgument($Description).AddArgument($InitialDirectory)

    $chosen = $null
    try {
        $output = $shell.Invoke()
        foreach ($item in $output) {
            if (-not [string]::IsNullOrWhiteSpace([string]$item)) {
                $chosen = [string]$item
            }
        }
    }
    catch {
        throw ("The dialog could not be shown: {0}" -f $_.Exception.Message)
    }
    finally {
        $shell.Dispose()
        $runspace.Dispose()
    }

    return $chosen
}

# ============================================
# MAIN
# ============================================

Show-Banner

# ---------- 1. Pick what to check ----------
$wantFile = [bool]$File

if ([string]::IsNullOrWhiteSpace($Path)) {
    if (-not $File) {
        Write-ColorOutput 'What do you want to check?' 'White'
        Write-ColorOutput '   1 = a folder' 'Gray'
        Write-ColorOutput '   2 = one file' 'Gray'
        Write-ColorOutput '   0 = cancel' 'Gray'
        $answer = ''
        try { $answer = (Read-Host '   Choose') } catch { $answer = '' }
        Write-Host ''

        switch ($answer.Trim()) {
            '2' { $wantFile = $true }
            '1' { $wantFile = $false }
            '0' { Write-ColorOutput 'Nothing was checked.' 'Yellow'; return }
            '' { $wantFile = $false }
            default { $wantFile = $false }
        }
    }

    if ($wantFile) {
        Write-ColorOutput 'Opening the file dialog...' 'Gray'
        $Path = Select-TargetPath -Kind 'File' -Description 'Select the file you want to check'
    }
    else {
        Write-ColorOutput 'Opening the folder dialog...' 'Gray'
        $Path = Select-TargetPath -Kind 'Folder' -Description 'Select the folder you want to check'
    }

    if ([string]::IsNullOrWhiteSpace($Path)) {
        Write-Host ''
        Write-ColorOutput 'Nothing selected. Nothing was checked.' 'Yellow'
        return
    }
}

$Path = $Path.Trim('"').Trim()
if (-not (Test-Path -LiteralPath $Path)) {
    Write-ColorOutput ("Path not found: {0}" -f $Path) 'Red'
    return
}

$target = (Resolve-Path -LiteralPath $Path).ProviderPath
$isFolder = Test-Path -LiteralPath $target -PathType Container
$kindText = if ($isFolder) { 'Folder' } else { 'File' }

if (-not $isFolder) { $wantFile = $true }

# ---------- 2. Build the list of paths to check ----------
# Restart Manager tracks files. A folder path handed to it comes back as "access denied",
# so a folder is checked through the files inside it - which is what a program holds anyway.
$resources = New-Object System.Collections.Generic.List[string]
$fileCount = 0
$truncated = $false

if ($isFolder) {
    $getFiles = @{
        LiteralPath = $target
        File        = $true
        Force       = $true
        Recurse     = [bool]$Recurse
        ErrorAction = 'SilentlyContinue'
    }

    $inner = @(Get-ChildItem @getFiles)
    if ($inner.Count -gt $MaxFiles) {
        $inner = @($inner | Select-Object -First $MaxFiles)
        $truncated = $true
    }

    foreach ($item in $inner) {
        $resources.Add($item.FullName) | Out-Null
        $fileCount++
    }

    if ($fileCount -eq 0) {
        Write-ColorOutput ("Folder: {0}" -f $target) 'White'
        Write-Host ''
        Write-ColorOutput 'That folder has no files in it, so there is nothing to check.' 'Yellow'
        Write-Host ''
        return
    }
}
else {
    $resources.Add($target) | Out-Null
}

Write-ColorOutput ("{0,-7}: {1}" -f $kindText, $target) 'White'
if ($isFolder) {
    $scopeText = if ($Recurse) { 'all files, subfolders included' } else { 'files in the top level only' }
    Write-ColorOutput ("Checked : {0} file(s)  ({1})" -f $fileCount, $scopeText) 'Gray'
    if ($truncated) {
        Write-ColorOutput ("Note    : stopped at {0} files. Use -MaxFiles to check more." -f $MaxFiles) 'DarkYellow'
    }
}
Write-Host ''

# ---------- 3. Ask Windows who is holding it ----------
Write-ColorOutput 'Asking the Restart Manager...' 'Gray'

try {
    $locks = @([FileLockCheck.LockFinder]::Find([string[]]$resources.ToArray()))
}
catch {
    Write-Host ''
    Write-ColorOutput ('Could not check the path: {0}' -f $_.Exception.Message) 'Red'
    return
}

Write-Host ''

if ($locks.Count -eq 0) {
    Write-ColorOutput 'Nothing is locking it. No program has this open right now.' 'Green'
    Write-Host ''
    Write-ColorOutput 'If a copy still fails, the cause is not an open handle:' 'Gray'
    Write-ColorOutput '  - permissions (the account is not allowed), or' 'Gray'
    Write-ColorOutput '  - a lock only the kernel or a share can see.' 'Gray'
    Write-ColorOutput 'Resource Monitor (resmon) and Sysinternals handle.exe show those.' 'Gray'
    Write-Host ''
    return
}

# ---------- 4. Add the details of each process ----------
$rows = foreach ($lock in $locks) {
    $process = $null
    try { $process = Get-Process -Id $lock.ProcessId -ErrorAction SilentlyContinue } catch { }

    $processName = '(gone)'
    $windowTitle = ''
    $exePath = ''
    if ($null -ne $process) {
        $processName = $process.ProcessName
        try { $windowTitle = $process.MainWindowTitle } catch { }
        try { if ($process.Path) { $exePath = $process.Path } } catch { }
    }

    $appName = $lock.AppName
    if ([string]::IsNullOrWhiteSpace($appName)) { $appName = $windowTitle }
    if ([string]::IsNullOrWhiteSpace($appName)) { $appName = '-' }

    $typeText = ($lock.AppType -replace '^Rm', '')

    $started = '-'
    if ($lock.StartTime -ne [datetime]::MinValue) {
        $started = $lock.StartTime.ToString('yyyy-MM-dd HH:mm:ss')
    }

    [PSCustomObject]@{
        ProcessId  = $lock.ProcessId
        Process    = $processName
        App        = $appName
        Type       = $typeText
        Session    = $lock.SessionId
        Started    = $started
        Service    = $lock.ServiceName
        Executable = $exePath
    }
}

$rows = @($rows | Sort-Object -Property ProcessId -Unique)

# ---------- 5. Show the table ----------
$width = 120
try { $width = $Host.UI.RawUI.WindowSize.Width } catch { }

$pidWidth = 7
$procWidth = 22
$appWidth = 24
$typeWidth = 10
$sessWidth = 7
$startedWidth = 19
$usedWidth = $pidWidth + $procWidth + $appWidth + $typeWidth + $sessWidth + $startedWidth + 6
$execWidth = [Math]::Max(12, $width - $usedWidth)

Write-ColorOutput ('  {0,-7} {1,-22} {2,-24} {3,-10} {4,-7} {5,-19} {6}' -f `
        'PID', 'Process', 'App (Restart Manager)', 'Type', 'Session', 'Started', 'Runs from') 'Cyan'
Write-ColorOutput ('  {0}' -f ('-' * [Math]::Min($width - 2, 190))) 'DarkGray'

foreach ($row in $rows) {
    $line = '  {0,-7} {1,-22} {2,-24} {3,-10} {4,-7} {5,-19} {6}' -f `
        $row.ProcessId,
        (Get-ShortText $row.Process $procWidth),
        (Get-ShortText $row.App $appWidth),
        (Get-ShortText $row.Type $typeWidth),
        $row.Session,
        $row.Started,
        (Get-ShortText $row.Executable $execWidth)
    Write-ColorOutput $line 'White'
}

Write-Host ''

$processWord = if ($rows.Count -eq 1) { 'process' } else { 'processes' }
Write-ColorOutput ('  {0} {1} holding this path.' -f $rows.Count, $processWord) 'Yellow'

foreach ($row in $rows) {
    if (-not [string]::IsNullOrWhiteSpace($row.Service)) {
        Write-ColorOutput ('    - service: {0}' -f $row.Service) 'Gray'
    }
}

Write-Host ''
Write-ColorOutput '  A handle can also belong to a program that is not the file''s owner:' 'Gray'
Write-ColorOutput '  Explorer thumbnails, antivirus scanners and indexers count too.' 'Gray'
Write-Host ''

# ---------- 6. Optionally end those processes ----------
if ($Stop) {
    Write-ColorOutput 'Stopping the processes that hold it...' 'Yellow'

    $stopped = 0
    $failed = 0

    foreach ($row in $rows) {
        $label = ('PID {0} ({1})' -f $row.ProcessId, $row.Process)

        if (-not $PSCmdlet.ShouldProcess($label, 'Stop process')) {
            Write-ColorOutput ('  [would stop] {0}' -f $label) 'DarkYellow'
            continue
        }

        try {
            Stop-Process -Id $row.ProcessId -Force -ErrorAction Stop
            $stopped++
            Write-ColorOutput ('  [stopped]    {0}' -f $label) 'Green'
        }
        catch {
            $failed++
            Write-ColorOutput ('  [failed]     {0} - {1}' -f $label, $_.Exception.Message) 'Red'
        }
    }

    Write-Host ''
    Write-ColorOutput ('  Stopped : {0}' -f $stopped) 'Green'
    Write-ColorOutput ('  Failed  : {0}' -f $failed) $(if ($failed -gt 0) { 'Red' } else { 'Gray' })

    if ($WhatIfPreference) {
        Write-Host ''
        Write-ColorOutput 'DRY RUN - no process was stopped.' 'Yellow'
        Write-ColorOutput 'Run it again without -WhatIf to stop them for real.' 'Gray'
    }

    Write-Host ''
}

Write-ColorOutput 'Done.' 'Green'
Write-Host ''

# The table is also returned, so it can be saved or piped:
#   .\FileLocker.ps1 -Path 'D:\Downloads' | Export-Csv locks.csv -NoTypeInformation
$rows
