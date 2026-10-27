$UserFolders = Get-ChildItem -Path . -Directory
foreach ($Folder in $UserFolders) {
    $Path = $Folder.FullName
    takeown /F $Path /R /D Y
    icacls $Path /grant "COMLAB\Domain Admins:(OI)(CI)F" /T /Q
}