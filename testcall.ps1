param (
    [Parameter(Mandatory = $false)]
    [Switch]$setup
)
    write-host "Called with the following"

foreach ($key in $MyInvocation.BoundParameters.keys)
{
    $value = (get-variable $key).Value 
    write-host "$key -> $value"
}

write-host 'second'
(Get-Command -Name $PSCommandPath).Parameters | Format-Table -AutoSize @{ Label = "Key"; Expression={$_.Key}; }, @{ Label = "Value"; Expression={(Get-Variable -Name $_.Key -EA SilentlyContinue).Value}; }
write-host '------------------------------------'
write-host "PSScriptRoot -> $PSScriptRoot"

write-host "PSCommandPath -> $PSCommandPath"
write-host "Desktop : " $([Environment]::GetFolderPath("Desktop"))
write-host "{env:ProgramData} ${env:ProgramData})"
write-host "Install directory " $(join-path -Path ${env:ProgramData} -childpath 'EmbroideryOrganize') 
start-sleep -Seconds 60 