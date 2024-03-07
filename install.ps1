$scriptDir = join-path -Path ${env:ProgramData} -childpath 'EmbroideryOrganize' 
function Get-FileTo($file) {
    $url = "https://raw.githubusercontent.com/D-Jeffrey/Embroidery-File-Organize/main/" + $file
    $downloadFile = Join-Path -Path $scriptDir -ChildPath $file
    write-host "Downloading ..`n  $url to`n  $downloadFile" -ForegroundColor Green
    $downloadFromGitHub = Invoke-WebRequest -Uri "$url" 
    if ($downloadFromGitHub.Content.Length -gt 0) {
        Set-Content -Path $downloadFile -Value $downloadFromGitHub.Content -Force
    }
}

New-Item -ItemType Directory -Path $scriptDir -ErrorAction SilentlyContinue| out-null

$pushscript = 'EmbroideryCollection-Cleanup.ps1'

Get-FileTo -file $pushscript
$scriptname = Join-Path -Path $scriptDir -ChildPath $pushscript
Get-FileTo -file 'EmbroideryManager.ico'
# for next upgrade
Get-FileTo -file 'install.ps1'

Powershell -NoLogo -ExecutionPolicy Bypass -File $scriptname -setup