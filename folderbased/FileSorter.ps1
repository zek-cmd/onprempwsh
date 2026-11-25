#Requires -Version 5.1
<#
.SYNOPSIS
    Sorts the loose files of a folder into category subfolders, IDM style.

.DESCRIPTION
    Picks one folder (through the Windows "Browse for Folder" dialog when -Path is
    not given) and moves every loose file in it into a category subfolder built
    the way Internet Download Manager builds its folders:

        Compressed, Documents, Music, Pictures, Programs, Video, Other

    A wide list of extensions is recognised, covering files made on Windows, macOS
    and Linux. Anything unknown goes into "Other", so no file is left behind.

    Rules the script follows:
      - Subfolders that already exist in the folder are never touched; only loose
        files in the top level are sorted.
      - The category folders themselves are never touched, so running the script
        twice is safe.
      - A file that would land on an existing name gets " (1)", " (2)" ... added,
        so nothing is ever overwritten.
      - Hidden and system files are skipped, and so are desktop.ini, thumbs.db
        and .DS_Store. Use -IncludeHidden to sort them too.
      - A file that is locked or in use is reported and left where it is.

.PARAMETER Path
    The folder to sort. When left out, the folder dialog opens.

.PARAMETER Copy
    Copy each file instead of moving it. The original folder keeps its files.

.PARAMETER IncludeHidden
    Sort hidden and system files as well.

.PARAMETER Open
    Open the sorted folder in File Explorer when the sort finishes.

.PARAMETER WhatIf
    Show what would happen and the folders that would be made, without moving or
    copying anything.

.EXAMPLE
    .\FileSorter.ps1

.EXAMPLE
    .\FileSorter.ps1 -Path 'D:\Downloads' -WhatIf

.EXAMPLE
    .\FileSorter.ps1 -Path 'D:\Downloads' -Copy -Open
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [string]$Path,
    [switch]$Copy,
    [switch]$IncludeHidden,
    [switch]$Open
)

$ErrorActionPreference = 'Stop'

# ============================================
# CATEGORIES  (IDM style)
# ============================================

$CategoryOrder = @('Compressed', 'Documents', 'Music', 'Pictures', 'Programs', 'Video', 'Other')

$CategoryColor = @{
    Compressed = 'Magenta'
    Documents  = 'Cyan'
    Music      = 'Yellow'
    Pictures   = 'DarkYellow'
    Programs   = 'Green'
    Video      = 'Blue'
    Other      = 'Gray'
}

$UnknownCategory = 'Other'

# Extensions grouped per category. Every entry starts with a dot and is lower case.
$ExtensionMap = @{
    'Compressed' = @(
        # archives
        '.zip', '.zipx', '.rar', '.7z', '.tar', '.gz', '.tgz', '.gzip', '.bz2', '.tbz', '.tbz2', '.bzip2',
        '.xz', '.txz', '.lz', '.lzma', '.zst', '.zstd', '.lz4', '.br', '.z', '.arj', '.ace', '.lzh', '.lha',
        '.arc', '.pak', '.cpio', '.shar', '.sit', '.sitx', '.hqx', '.zoo', '.cpt', '.pea',
        # disk images and packaged images
        '.iso', '.img', '.dmg', '.vhd', '.vhdx', '.vmdk', '.wim', '.swm', '.esd', '.toast',
        # package blobs from various ecosystems
        '.cab', '.nupkg', '.whl', '.gem', '.crate', '.xpi', '.zi', '.bundle'
    )
    'Documents'  = @(
        # text and office documents
        '.doc', '.docx', '.docm', '.dot', '.dotx', '.dotm', '.rtf', '.txt', '.text', '.md', '.markdown',
        '.rst', '.adoc', '.asciidoc', '.log', '.nfo', '.readme', '.pdf', '.xps', '.oxps', '.djvu', '.djv',
        '.epub', '.mobi', '.azw', '.azw3', '.kf8', '.fb2', '.cbr', '.cbz', '.cb7', '.cbt', '.lit', '.lrf',
        '.tex', '.latex', '.bib', '.wpd', '.wps', '.wri', '.odt', '.ott', '.sxw', '.stw', '.abw', '.zabw',
        '.pages', '.pap', '.scriv', '.manuscript',
        # spreadsheets, data and tables
        '.xls', '.xlsx', '.xlsm', '.xlsb', '.xlt', '.xltx', '.xltm', '.xlam', '.csv', '.tsv', '.tab',
        '.ods', '.ots', '.sxc', '.stc', '.numbers', '.gnumeric', '.dif', '.slk', '.dbf', '.accdb', '.mdb',
        '.sqlite', '.sqlite3', '.db', '.sql', '.parquet', '.feather', '.hdf', '.hdf5', '.nc',
        # presentations
        '.ppt', '.pptx', '.pptm', '.pps', '.ppsx', '.ppsm', '.pot', '.potx', '.potm', '.odp', '.otp',
        '.sxi', '.sti', '.key', '.kth', '.shw',
        # notes, mail and contacts
        '.one', '.onetoc2', '.onenote', '.msg', '.eml', '.emlx', '.mbox', '.oft', '.pst', '.ost', '.vcf',
        '.vcard', '.ics', '.ical',
        # plain data and settings files
        '.json', '.jsonc', '.json5', '.xml', '.yaml', '.yml', '.toml', '.ini', '.inf', '.cfg', '.conf',
        '.config', '.properties', '.env', '.reg', '.plist', '.diff', '.patch', '.lock', '.chm', '.hlp'
    )
    'Music'      = @(
        '.mp3', '.mp2', '.mpga', '.wav', '.wave', '.flac', '.aac', '.m4a', '.m4b', '.m4p', '.m4r', '.alac',
        '.ogg', '.oga', '.opus', '.wma', '.aiff', '.aif', '.aifc', '.ape', '.wv', '.tta', '.tak', '.dsf',
        '.dff', '.mid', '.midi', '.mod', '.xm', '.it', '.s3m', '.amr', '.awb', '.au', '.snd', '.ra', '.ram',
        '.cda', '.ac3', '.dts', '.mka', '.spx', '.ofr', '.ofs', '.shn', '.vgm', '.vgz', '.nsf', '.gbs',
        '.m3u', '.m3u8', '.pls', '.wpl', '.asx', '.cue', '.xspf', '.zpl'
    )
    'Pictures'   = @(
        # everyday photos and screenshots
        '.jpg', '.jpeg', '.jpe', '.jfif', '.jif', '.pjpeg', '.pjp', '.png', '.apng', '.gif', '.bmp', '.dib',
        '.tif', '.tiff', '.webp', '.heic', '.heif', '.hif', '.avif', '.svg', '.svgz', '.ico', '.cur',
        '.ani', '.tga', '.icb', '.vda', '.vst', '.pcx', '.pcc', '.ppm', '.pgm', '.pbm', '.pnm', '.pfm',
        '.xbm', '.xpm', '.wbmp', '.dds', '.ktx', '.ktx2', '.exr', '.hdr', '.pic', '.sgi', '.rgb', '.rgba',
        '.jp2', '.j2k', '.jpf', '.jpx', '.jpm', '.jbig', '.jb2', '.wdp', '.hdp', '.jxl',
        # camera raw files
        '.raw', '.cr2', '.cr3', '.crw', '.nef', '.nrw', '.arw', '.srf', '.sr2', '.orf', '.rw2', '.rwl',
        '.raf', '.pef', '.dng', '.3fr', '.fff', '.iiq', '.mos', '.mrw', '.x3f', '.erf', '.kdc', '.dcr',
        # artwork and design files
        '.psd', '.psb', '.ai', '.eps', '.epsf', '.indd', '.idml', '.xcf', '.kra', '.ora', '.clip', '.csp',
        '.sai', '.sai2', '.afphoto', '.afdesign', '.procreate', '.sketch', '.fig', '.cdr', '.cmx', '.fh',
        '.vsd', '.vsdx', '.wmf', '.emf', '.cgm', '.xwd', '.gd', '.gd2', '.pat', '.gbr', '.abr'
    )
    'Programs'   = @(
        # Windows executables, installers and scripts
        '.exe', '.msi', '.msix', '.msixbundle', '.appx', '.appxbundle', '.appinstaller', '.com', '.scr',
        '.pif', '.cpl', '.dll', '.ocx', '.drv', '.sys', '.ax', '.mui', '.tsp', '.efi', '.bin',
        '.bat', '.cmd', '.ps1', '.psm1', '.psd1', '.vbs', '.vbe', '.js', '.jse', '.wsf', '.wsh', '.hta',
        '.ahk', '.au3', '.py', '.pyw', '.pyc', '.pyo', '.rb', '.pl', '.sh', '.bash', '.zsh', '.fish',
        '.csh', '.ksh', '.jar', '.war', '.ear', '.class', '.apk', '.aab', '.apks', '.xapk', '.ipa',
        # other operating systems' installers and packages
        '.app', '.pkg', '.mpkg', '.deb', '.rpm', '.snap', '.flatpak', '.appimage', '.run',
        '.xdg', '.pet', '.slp', '.pup', '.ebuild',
        # developer project and code files
        '.cs', '.vb', '.fs', '.java', '.kt', '.kts', '.swift', '.m', '.mm', '.h', '.hpp', '.c', '.cpp',
        '.cc', '.cxx', '.go', '.rs', '.php', '.asp', '.aspx', '.jsp', '.htm', '.html', '.xhtml', '.css',
        '.scss', '.sass', '.less', '.tsx', '.jsx', '.vue', '.svelte', '.lua', '.r', '.scala',
        '.clj', '.ex', '.exs', '.dart', '.sol', '.asm', '.sln', '.csproj', '.vbproj',
        '.vcxproj', '.vcproj', '.fsproj', '.props', '.targets', '.gradle', '.cmake', '.mk', '.makefile',
        '.dockerfile', '.gitignore', '.gitattributes', '.editorconfig'
    )
    'Video'      = @(
        '.mp4', '.m4v', '.mkv', '.mk3d', '.avi', '.mov', '.qt', '.wmv', '.asf', '.flv', '.f4v', '.webm',
        '.mpg', '.mpeg', '.mpe', '.m2v', '.m1v', '.mpv', '.vob', '.ifo', '.bup', '.ts', '.mts', '.m2ts',
        '.m2t', '.3gp', '.3g2', '.3gpp', '.3gpp2', '.ogv', '.rm', '.rmvb', '.divx', '.xvid',
        '.dv', '.mxf', '.r3d', '.braw', '.amv', '.svi', '.nsv', '.roq', '.y4m', '.yuv',
        '.h264', '.h265', '.hevc', '.avc', '.vp8', '.vp9', '.av1', '.gifv',
        # subtitles and video side files
        '.srt', '.vtt', '.ass', '.ssa', '.sub', '.sbv', '.smi', '.sami', '.ttml', '.dfxp'
    )
}

# Build one lookup table: extension -> category.
$ExtToCategory = @{}
foreach ($category in $ExtensionMap.Keys) {
    foreach ($ext in $ExtensionMap[$category]) {
        if (-not $ExtToCategory.ContainsKey($ext)) {
            $ExtToCategory[$ext] = $category
        }
    }
}

# Files that are never worth moving around.
$SkipFileNames = @('desktop.ini', 'thumbs.db', 'ehthumbs.db', '.ds_store', '.directory', 'icon')

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
    Write-ColorOutput '   FILE SORTER  -  IDM STYLE FOLDERS' 'Yellow'
    Write-ColorOutput '========================================' 'Cyan'
    $mode = if ($Copy) { 'Copy files' } else { 'Move files' }
    if ($WhatIfPreference) { $mode += '  (DRY RUN - nothing is changed)' }
    Write-ColorOutput ("   Mode    : {0}" -f $mode) 'Gray'
    Write-ColorOutput ("   Folders : {0}" -f ($CategoryOrder -join ', ')) 'Gray'
    Write-ColorOutput ("   Known   : {0} file extensions" -f $ExtToCategory.Count) 'Gray'
    Write-Host ''
}

function Select-SortFolder {
    <#
        Shows the Windows "Browse for Folder" dialog and returns the chosen path,
        or $null when the user cancels. The dialog runs on a single-threaded (STA)
        thread, which is what it needs under both Windows PowerShell 5.1 and
        PowerShell 7.
    #>
    param(
        [string]$Description = 'Select the folder whose files you want to sort',
        [string]$InitialDirectory
    )

    $dialogCode = {
        param($Description, $InitialDirectory)

        Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop

        $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
        $dialog.Description = $Description
        try { $dialog.ShowNewFolderButton = $false } catch { }

        if (-not [string]::IsNullOrWhiteSpace($InitialDirectory) -and
            (Test-Path -LiteralPath $InitialDirectory -PathType Container)) {
            $dialog.SelectedPath = $InitialDirectory
        }

        # A hidden, always-on-top owner keeps the dialog in front of this window.
        $owner = New-Object System.Windows.Forms.Form
        $owner.TopMost = $true
        $owner.ShowInTaskbar = $false
        $owner.Opacity = 0
        $owner.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen

        $chosen = $null
        try {
            $answer = $dialog.ShowDialog($owner)
            if ($answer -eq [System.Windows.Forms.DialogResult]::OK) {
                $chosen = $dialog.SelectedPath
            }
        }
        finally {
            $owner.Dispose()
            $dialog.Dispose()
        }

        return $chosen
    }

    $apartmentState = [System.Threading.Thread]::CurrentThread.GetApartmentState()
    if ($apartmentState -eq [System.Threading.ApartmentState]::STA) {
        return (& $dialogCode $Description $InitialDirectory)
    }

    # Single-threaded apartment not available here (PowerShell 7): use a temporary
    # runspace that has one.
    $runspace = [runspacefactory]::CreateRunspace()
    $runspace.ApartmentState = [System.Threading.ApartmentState]::STA
    $runspace.ThreadOptions = 'ReuseThread'
    $runspace.Open()

    $shell = [powershell]::Create()
    $shell.Runspace = $runspace
    [void]$shell.AddScript($dialogCode.ToString()).AddArgument($Description).AddArgument($InitialDirectory)

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
        throw ("The folder dialog could not be shown: {0}" -f $_.Exception.Message)
    }
    finally {
        $shell.Dispose()
        $runspace.Dispose()
    }

    return $chosen
}

function Get-UniqueTarget {
    # Returns a full path in $Folder that does not exist yet, so no file is lost.
    param(
        [string]$Folder,
        [string]$FileName
    )

    $candidate = Join-Path $Folder $FileName
    if (-not (Test-Path -LiteralPath $candidate)) { return $candidate }

    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($FileName)
    $extension = [System.IO.Path]::GetExtension($FileName)
    if ([string]::IsNullOrWhiteSpace($baseName)) { $baseName = 'file' }

    $number = 1
    while ($true) {
        $candidate = Join-Path $Folder ('{0} ({1}){2}' -f $baseName, $number, $extension)
        if (-not (Test-Path -LiteralPath $candidate)) { return $candidate }
        $number++
    }
}

function Get-Category {
    param([string]$Extension)

    if ([string]::IsNullOrWhiteSpace($Extension)) { return $UnknownCategory }
    $ext = $Extension.ToLowerInvariant()
    if ($ExtToCategory.ContainsKey($ext)) { return $ExtToCategory[$ext] }
    return $UnknownCategory
}

# ============================================
# MAIN
# ============================================

Show-Banner

# ---------- 1. Pick the folder ----------
if ([string]::IsNullOrWhiteSpace($Path)) {
    Write-ColorOutput 'Opening the folder dialog...' 'Gray'
    $Path = Select-SortFolder -Description 'Select the folder whose files you want to sort'

    if ([string]::IsNullOrWhiteSpace($Path)) {
        Write-Host ''
        Write-ColorOutput 'No folder selected. Nothing was sorted.' 'Yellow'
        return
    }
}

$Path = $Path.Trim('"').Trim()
if (-not (Test-Path -LiteralPath $Path)) {
    Write-ColorOutput ("Folder not found: {0}" -f $Path) 'Red'
    return
}

$root = (Resolve-Path -LiteralPath $Path).ProviderPath
if (-not (Test-Path -LiteralPath $root -PathType Container)) {
    Write-ColorOutput ("That is a file, not a folder: {0}" -f $root) 'Red'
    return
}

Write-ColorOutput ("Folder : {0}" -f $root) 'White'
Write-Host ''

# ---------- 2. Collect the loose files ----------
$getFiles = @{
    LiteralPath = $root
    File        = $true
    ErrorAction = 'Stop'
}
if ($IncludeHidden) { $getFiles['Force'] = $true }

try {
    $files = @(Get-ChildItem @getFiles)
}
catch {
    Write-ColorOutput ("The folder could not be read: {0}" -f $_.Exception.Message) 'Red'
    return
}

if ($files.Count -eq 0) {
    Write-ColorOutput 'No loose files to sort in this folder. Nothing was changed.' 'Yellow'
    if ($Open) { Invoke-Item -LiteralPath $root }
    return
}

# ---------- 3. Sort ----------
$counts = @{}
foreach ($category in $CategoryOrder) { $counts[$category] = 0 }

$movedCount = 0
$skippedCount = 0
$failedCount = 0
$failures = New-Object System.Collections.Generic.List[string]
$verb = if ($Copy) { 'Copied' } else { 'Moved' }
$actionName = if ($Copy) { 'Copy' } else { 'Move' }

foreach ($file in $files) {
    # Never move the category folders or anything already sorted.
    if ($file.PSIsContainer) { continue }

    $lowerName = $file.Name.ToLowerInvariant()
    if ($SkipFileNames -contains $lowerName) {
        $skippedCount++
        continue
    }
    if (-not $IncludeHidden -and ($file.Attributes -band [System.IO.FileAttributes]::Hidden)) {
        $skippedCount++
        continue
    }

    $category = Get-Category -Extension $file.Extension
    $color = $CategoryColor[$category]
    $destinationFolder = Join-Path $root $category

    try {
        if ($PSCmdlet.ShouldProcess($file.FullName, ('{0} to [{1}]' -f $actionName, $category))) {

            if (-not (Test-Path -LiteralPath $destinationFolder -PathType Container)) {
                New-Item -Path $destinationFolder -ItemType Directory -Force | Out-Null
            }

            $target = Get-UniqueTarget -Folder $destinationFolder -FileName $file.Name

            if ($Copy) {
                Copy-Item -LiteralPath $file.FullName -Destination $target -ErrorAction Stop
            }
            else {
                Move-Item -LiteralPath $file.FullName -Destination $target -ErrorAction Stop
            }

            $counts[$category]++
            $movedCount++
            Write-ColorOutput ('  [{0,-10}] {1}' -f $category, $file.Name) $color
        }
        else {
            # Dry run: report it, change nothing.
            $counts[$category]++
            $movedCount++
            Write-ColorOutput ('  [{0,-10}] {1}  ->  would go to "{2}"' -f $category, $file.Name, $category) $color
        }
    }
    catch {
        $failedCount++
        $reason = $_.Exception.Message
        $failures.Add(('{0} - {1}' -f $file.Name, $reason)) | Out-Null
        Write-ColorOutput ('  [{0,-10}] {1}  FAILED: {2}' -f $category, $file.Name, $reason) 'Red'
    }
}

# ---------- 4. Summary ----------
Write-Host ''
Write-ColorOutput '----------------------------------------' 'Cyan'
Write-ColorOutput '   SUMMARY' 'Yellow'
Write-ColorOutput '----------------------------------------' 'Cyan'

foreach ($category in $CategoryOrder) {
    $number = $counts[$category]
    if ($number -eq 0) { continue }

    $folderPath = Join-Path $root $category
    $sizeText = ''
    if (Test-Path -LiteralPath $folderPath) {
        try {
            $bytes = (Get-ChildItem -LiteralPath $folderPath -File -Recurse -ErrorAction SilentlyContinue |
                Measure-Object -Property Length -Sum).Sum
            if ($null -ne $bytes) { $sizeText = '  ({0:N1} MB)' -f ($bytes / 1MB) }
        }
        catch { }
    }
    Write-ColorOutput ('  {0,-12} {1,4} file(s){2}' -f $category, $number, $sizeText) $CategoryColor[$category]
}

Write-Host ''
Write-ColorOutput ('  {0}   : {1}' -f $verb, $movedCount) 'Green'
Write-ColorOutput ('  Skipped : {0}' -f $skippedCount) 'Gray'
Write-ColorOutput ('  Failed  : {0}' -f $failedCount) $(if ($failedCount -gt 0) { 'Red' } else { 'Gray' })

if ($WhatIfPreference) {
    Write-Host ''
    Write-ColorOutput 'DRY RUN - no file was moved or copied and no folder was created.' 'Yellow'
    Write-ColorOutput 'Run the script again without -WhatIf to do it for real.' 'Gray'
}

if ($failures.Count -gt 0) {
    Write-Host ''
    Write-ColorOutput '  Files that could not be sorted (still in place):' 'Red'
    foreach ($line in $failures) {
        Write-ColorOutput ("    - {0}" -f $line) 'DarkYellow'
    }
}

if ($Open -and -not $WhatIfPreference) {
    try { Invoke-Item -LiteralPath $root } catch { }
}

Write-Host ''
Write-ColorOutput 'Done.' 'Green'
