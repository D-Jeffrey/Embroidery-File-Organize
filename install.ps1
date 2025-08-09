param (
    [Parameter(Mandatory = $false)]
    [String]$OldVersion
)
if ($isLinux -or $isMacOS) {
    $scriptDir = join-path -Path ${env:HOME} -childpath 'EmbroideryOrganize' 
} else {
    $scriptDir = join-path -Path ${env:ProgramData} -childpath 'EmbroideryOrganize' 
}
$sourceurl = "https://raw.githubusercontent.com/D-Jeffrey/Embroidery-File-Organize/main/" 

function Get-FileTo($file) {
    $url = $sourceurl + $file
    $downloadFile = Join-Path -Path $scriptDir -ChildPath $file
    write-host "Downloading ..`n  $url to`n  $downloadFile" -ForegroundColor Green
    $downloadFromGitHub = Invoke-WebRequest -Uri "$url"
    $content =  $downloadFromGitHub.Content
    if ($downloadFromGitHub.Content.Length -gt 0) {
        if ($downloadFromGitHub.Content.gettype().Name -eq 'String') {
            if ((Compare-Object -ReferenceObject ([System.Text.Encoding]::Unicode.GetBytes(($downloadFromGitHub.Content).substring(0,1))) -DifferenceObject @(255,254)).count -eq 0) {
                $content = $downloadFromGitHub.Content.Substring(1)
            }
        }
        Set-Content -Path $downloadFile -Value $content -Force
    }
}
function FetchImageFile ([string]$file) {
    $source = $sourceurl + $file
    $downloadFile = Join-Path -Path $scriptDir -ChildPath $file
    write-host "Downloading ..`n  $url to`n  $downloadFile" -ForegroundColor Green
    try {
        (New-Object System.Net.WebClient).DownloadFile($source, $downloadFile)
    } catch {
        Write-Host "`t[!] Failed to download '$source'"
        
    }
}
function Test-ExistsOnPath {
    param (
        [string]$FileName
    )
    
    $env:PATH.Split([System.IO.Path]::PathSeparator) | where-object {$_ -ne ""} | ForEach-Object {
        $fullPath = Join-Path $_ $FileName
        if (Test-Path $fullPath -PathType Leaf) {
            return $true
        }
    }

    return $false
}



New-Item -ItemType Directory -Path $scriptDir -ErrorAction SilentlyContinue| out-null

$pushscript = 'EmbroideryCollection-Cleanup.ps1'

Get-FileTo -file $pushscript
$scriptname = Join-Path -Path $scriptDir -ChildPath $pushscript
FetchImageFile -file 'EmbroideryManager.ico'
$params = @{}
if ($OldVersion) {
    $params.OldVersion = $OldVersion
}
# for next upgrade
$execName = "pwsh.exe" 
if ($IsLinux -or $IsMacOS) {
    $execName = "pwsh" 
} 
Get-FileTo -file 'install.ps1'
if (Test-ExistsOnPath $execName) {
    pwsh -NoLogo -ExecutionPolicy Bypass -File $scriptname -setup @params
} else {
    Powershell -NoLogo -ExecutionPolicy Bypass -File $scriptname -setup @params
    }