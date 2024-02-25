$scriptDir = join-path -Path ${env:ProgramData} -childpath 'EmbroideryOrganize' 
function Get-FileTo($file) {
    $url = "https://raw.githubusercontent.com/D-Jeffrey/Embroidery-File-Organize/main/$file"
    $downloadFromGitHub = Invoke-WebRequest -Uri "$url"
    $downloadFile = Join-Path -Path $scriptDir -ChildPath $file
    Set-Content -Path $downloadFile -Value $downloadFromGitHub.Content
}

New-Item -ItemType Directory -Path $scriptDir | out-null

$pushscript = 'EmbroideryCollection-Cleanup.ps1'
$pushscript = 'testcall.ps1'
Get-FileTo -file $pushscript
$scriptname = Join-Path -Path $scriptDir -ChildPath $pushscript
Get-FileTo -file 'EmbroideryManager.ico'

Powershell -NoLogo -ExecutionPolicy Bypass -File $scriptname -setup