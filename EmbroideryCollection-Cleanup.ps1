#
# EmbroideryCollection-Cleanup.ps1
#
# Deal with the many different types of embroidery files, put the right format types in mySewingnet Cloud
# We are looking to keep the ??? files types only from the zip files.
#
# Orginal Author: Darren Jeffrey Dec 2021
#                           Last Feb 2024
#

param
(
  [Parameter(Mandatory = $false)]
  [int32]$DownloadDaysOld = 7,                      # How many days old show it scan for Zip files in Download
  [int32]$SetSize = 10,
  [Switch]$KeepAllTypes,                            # Keep all the different types of a file (duplicate name but different extensions)
  [Switch]$CleanCollection,                         # Cleanup the Collection Folder to only EmbroidDir files
  [Switch]$CleanCollectionIgnoreDir,                # Cleanup the Collection Folder to only EmbroidDir files and look for duplicates ignoring directories meanings
# TODO
  [Switch]$CleanCloud,                              # Remove extra instructions and files from the cloud
# BUG in NoDirectory
#  [Switch]$NoDirectory,                             # Do not create directory structure in the upload from space
  [Switch]$OneDirectory,                              #Limit the folders to one directly deep only 
  [string]$EmbroidDir = "Embroidery",               # You may want to change this directory name inside of the 'Collection ' Directory wihtin your 'Documents' directory
                                                    # If it is just a name, then it is assumed to be within the defeault Documents directory, otherwise it will be taken as a full path
#   [string]$InstructDir = "Embroidery\Instructions",# This is a Directory name inside of "Documents" where instructions are saved
                                                    # If it is just a name, then it is assumed to be within the defeault Documents directory, otherwise it will be taken as a full path
  [string]$USBDrive,                                # Write the output to a USB Drives
  [Switch]$HardDelete,                              # Delete the files rather than sending to recycle bin
  [switch]$KeepEmptyDirectory,                      # If you don't want this to remove extra empty directories from Collection folders'
  [Switch]$Testing,                                 # Run it and see what happens
  [Switch]$Setup,                                   # Setup the Shortcut on the desktop to this application
  [Switch]$DragUpload,                              # Use the web page instead of the plug in to drag and drop
  [Switch]$ShowExample,                             # Show the example GIF
  [string]$ConfigFile = "EmbroideryCollection.cfg", # This is the file in the same directory as this script otherwise it is a full path
  [Switch]$PromptPassword,                          # Always Prompt for a password for MySewnet
  [Switch]$ConfigDefault,                           # Got back to default settings
  [Switch]$SwitchDefault,                           # Clear all the preview Switch enabled Values
  [Switch]$FirstRun,                                # Scan all the ZIP files
  [Switch]$Sync,                                    # Sync MySewnet to local folders
  [Switch]$APIBeta                                  # use MySewNet cloud API
  
  )

# $VerbosePreference =  "Continue"
# $InformationPreference =  "Continue"

# ******** CONFIGURATION 
$preferredSewType = 'vp3', 'vip', 'pcs','dst', 'pes', 'hus'
$alltypes = 'hus', 'dst', 'exp', 'jef', 'pes', 'vip', 'vp3', 'xxx', 'sew', 'vp4', 'pcs', 'vf3', 'csd', 'zsk', 'emd', 'ese', 'phc', 'art', 'ofm', 'pxf', 'svg', 'dxf'
$foldupDir = 'images', 'sewing helps', 'Designs', 'Design Files'

$goodInstructionTypes = ('pdf','doc', 'docx', 'txt','rtf', 'mp4', 'ppt', 'pptx', 'gif', 'jpg', 'png', 'bmp','mov', 'wmv', 'avi','mpg', 'm4v' )
$TandCs = @('TERMS-OF-USAGE.*', 'planetappliquetermsandconditions.*','READ ME FIRST.rtf','*copyright.*','*copyright Statement.*','*copyrights.*',
    'copyrightStatement.*','License agreement.*', 'License.*','termsofuse.*')
$opencloudpage = "https://www.mysewnet.com/en-us/my-account/#/cloud/"
# List of paramstring to check
$paramstring =  [ordered]@{
 "DownloadDaysOld" = "Age of files in Download directory";
 "SetSize" = "Keep collections of files together if there are at least this many";
 "EmbroidDir" = "Embriodary Files directory";
 "USBDrive"="USB drive letter (example E: or H:)";
 "MySewNetuserid"= "User ID to Login to MySewNet";
 "MySewNetpw"="" 
}


$parambool = [ordered]@{
'KeepAllTypes'= 'Keep all variations of files types' ; 
'KeepEmptyDirectory'= 'When cleaning up keep empty folders'; 
'DragUpload'= 'Open the mysewnet Cloud browser interface for drag and drop';
'ShowExample'= 'Show how to upload to mySewnet';
'NoDirectory'= 'Do not use Directories from Zip files which creating collection';
'OneDirectory'= 'Keep files a maximum of one directory deep ';
'APIBeta'= 'Use MySewnet Cloud';
'PromptPassword'= 'Always prompt for password to login to MySewnet Cloud'}
$paramarray = [ordered]@{
'preferredSewType' = 'The preferred types of Embriodary file types';
'alltypes' = 'All the possible types of files which are an Embriodary file'; 
'foldupDir' = 'Remove/fold folders of this name'; 
'goodInstructionTypes' = 'Instructions file types which should be saved with files' 
}
$paramswitch =[ordered]@{
    'CleanCollection' = 'Clean the Collection folder';
    'CleanCollectionIgnoreDir' = 'Clean the Collection folder and Ignore Directory structure';
    'APIBeta' = "Using BETA API for mySewNet";
    'Sync' = 'Syncronize computer folders to Cloud'
}



# ----------------------------------------------------------------------
#                 $alltypes
# this is a list of all the different types of embrodiary files that are considered.  
# The '$preferredSewType' should be from the list below based on what is best for your 
# machine and in the order that you prefer.  If there are more than one copy of a file 
# type it will select your first one
# ----------------------------------------------------------------------
#                 $TandCs              
# Term and Conditions added by various store that add up space with the same document type over and over, using up your MySewing Cloud space
# This is a file name pattern so TC.* will match TC.doc or TC.pdf
# ----------------------------------------------------------------------
#                  $foldupDir
# What directories should be flattened to bring the Embroidery files higher up so they are not nested instead of sub-folders.  
# The names are for Directories you want to remove the sub-folder and moved the contents up
# ----------------------------------------------------------------------
#
#

$ECCVERSION = "0.5.3"
write-host " ".padright(15) "Embroidery Collection Cleanup version: $ECCVERSION".padright(70) -ForegroundColor White -BackgroundColor Blue


if ($PSVersionTable['PSVersion'].major -lt 3 ) {
    write-Error "This will NOT work on your version of Powershell"
    write-host $PSVersionTable['PSVersion'].major
    $PSVersionTable
    return
}
$filecnt = 0
$sizecnt = 0
$Script:dircnt = 0
$Script:savecnt = 0
$Script:addsizecnt = 0
$Script:p = 0
$padder = 45
$filesToRemove = @()
$MySewNetuserid = ""
$MySewNetpw = ""

$shell = New-Object -ComObject 'Shell.Application'
$downloaddir = (New-Object -ComObject Shell.Application).NameSpace('shell:Downloads').Self.Path
$docsdir =[environment]::getfolderpath("mydocuments")
$docsdirwOne = $docsdir
if ($docsdir.tolower().contains('\onedrive')) {
    $docsdir = ${env:HOMEDRIVE} + ${env:HOMEPATH}
}
$tmpdir = ${env:temp} + "\cleansew.tmp"

$missingSewnetAddin = ((get-itemproperty -path "Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Classes\.vp3\ShellEx\{e357fccd-a995-4576-b01f-234630154e96}").'(default)' -ne "{370F9E36-A651-4BB3-89A9-A6DB957C63CC}") 

#
# Read and process the Config File
#
# Check if the ConfigFile path is not absolute
if (!$ConfigFile.Contains('\')) {
    # If it's a file only, join it with the script root path
    $ConfigFile = Join-Path -Path $PSScriptRoot -ChildPath $ConfigFile
}

# If ConfigDefault and ConfigFile exists, remove the ConfigFile
if ($ConfigDefault -and (Test-Path -Path $ConfigFile)) {
    Remove-Item -Path $ConfigFile
}
# If ConfigFile exists, read its content and convert from JSON
if (Test-Path -Path $ConfigFile) {
    
    # Get the content of the configuration file and convert it from JSON
    $SavedParam = Get-Content -Path $ConfigFile -Encoding Utf8 -ErrorAction SilentlyContinue | ConvertFrom-Json

    # Iterate over each parameter
    if ($null -eq $SwitchDefault) {
        $paralist = ($paramstring.Keys )
    } else {
        $paralist = ($paramstring.Keys + $parambool.Keys)
    }
    foreach ($param in ($paralist)) {
        # Use variable indirection to check if the parameter is present
        $isParamPresent = $PSBoundParameters.ContainsKey($param)

        # If the parameter is not present and it exists in the saved parameters, assign the saved value
        if (-not $isParamPresent -and $null -ne $SavedParam.$param) {
            Set-Variable -Name $param -Value $SavedParam.$param
            # write-host "$param = " $SavedParam.$param
        }
    }

    foreach ($param in ($paramarray.Keys)) {
        if ($null -ne $SavedParam.$param) {
            Set-Variable -Name $param -Value $SavedParam.$param.value
            
        }
    }
} else {
    $FirstRun = $true
}
    

function SaveAllParams
{
    # Save the state of the variables and settings
    $SavedParam = New-Object PSObject
    # Iterate over each variable name
    foreach ($param in ($parambool.Keys)) {
        $val = [bool](Get-Variable -Name $param -ValueOnly)
        # Add each variable to the object as a NoteProperty
        $SavedParam | Add-Member -MemberType NoteProperty -Name $param -Value $val
    }
        
    foreach ($param in ($paramstring.Keys + $paramarray.keys)) {
        $val = (Get-Variable -Name $param -ValueOnly)
        $SavedParam | Add-Member -MemberType NoteProperty -Name $param -Value $val
    }

    # Convert the object to a JSON string and save it to a file
    $SavedParam | ConvertTo-Json | Set-Content -Path $ConfigFile -Encoding Utf8
}



if ($missingSewnetAddin) {
    $DragUpload = $true
}
if ($APIBeta) {
    $DragUpload = $false
    $ShowExample = $false
}

$doit = !$Testing



# This is for development testing and debugging

if ($env:COMPUTERNAME -eq "DESKTOP-R3PSDBU_") { # -and $Testing) {
    $docsdir = "d:\Users\kjeff\"
    $downloaddir = "d:\Users\kjeff\downloads"
    $docsdirwOne = "d:\Users\kjeff\OneDrive\Documents"
    $doit = $true
    }


    
#=============================================================================================

function LogAction($File, $Action = "++Added", [Boolean]$isInstructions = $false) {
    $now = Get-Date -Format "yyyy/MMM/dd HH:mm "
    $extra = (&{if ($isInstructions) { " Instructions"} else { "" } })
    write-verbose "$Action $File typeof $extra"
    Add-Content -Path $LogFile -Value ("$now$Action $File $extra")
}

Function ShowProgress ([string]$Area, [string]$stat = $null)
{
    $Script:p++
    if ($stat) {
        Write-Progress -PercentComplete ($Script:p % 100 ) $Area -Status $stat
    } else {
        Write-Progress -PercentComplete ($Script:p % 100 ) $Area 
    }
}

Function RecycleFile ($file) {
    $file = (&{if ($file.GetType().Name -ne "String") { $file.FullName } else { $file} })
    if ($HardDelete) {
        Remove-Item -Path $file        # Handled by WhatIf
    } elseif ($doit) {
        $shell.NameSpace(0).ParseName($file).InvokeVerb('delete')
    }
}

function MyPause {
    param (
        [string]$Message,
        [bool]$Choice = $false,
        [string]$BoxMsg,
        [int]$Timeout = 0,
        [bool]$useGUI = $false,
        [bool]$ChoiceDefault = $true
    )

    $yes = $true
    # Check if running Powershell ISE
    if ($psISE -or $useGUI) {
        Add-Type -AssemblyName System.Windows.Forms
        $BoxMsg = if ($BoxMsg -eq "" -or $null -eq $BoxMsg) { $Message } else { $BoxMsg }
        $x = if ($Choice) {
            [System.Windows.Forms.MessageBox]::Show($BoxMsg, 'Cleanup Collection Folders', 'YesNo', 'Question')
        } else {
            [System.Windows.Forms.MessageBox]::Show($Message)
        }
        $yes = ($x -eq 'Yes')
    } else {
        $secondsRunning = 0
        if ($Timeout -gt 0) { 
            Start-Sleep -Milliseconds 100
            $host.ui.RawUI.FlushInputBuffer()
        }
        if ($choice) {
            if (!$Message.contains('?')) {
                $Message +=  "?"
            }
            $yesno = (&{if ($ChoiceDefault) { " (Y/n) " } else { " (y/N) " }})
        } else {
            $yesno = ""
        }

        Write-Host ($Message +  $yesno ) -ForegroundColor Yellow -NoNewline
        while (-not $Host.UI.RawUI.KeyAvailable -and $secondsRunning++ -lt $Timeout) {
            [Threading.Thread]::Sleep(1000)
        }
        if ($Host.UI.RawUI.KeyAvailable -or $Timeout -eq 0) {
            $needakey = $true
            while ($needakey) {
                $keystroke = $host.ui.RawUI.ReadKey("NoEcho,IncludeKeyDown")
                $yes = 'Yy'.Contains($keystroke.Character)
                # Look for yes or no keystrokes
                if ($choice) {
                    $needakey = -not ('YyNn'.Contains($keystroke.Character))
                } else {
                    $needakey = $false
                }
                # Use the Default selection if the ENTER key is pressed
                if ($keystroke.VirtualKeyCode -eq 13) {
                    $needakey = $false
                    $yes = $ChoiceDefault
                }
            }
        }
        if ($choice) {
            $selectchoice = (&{if ($yes) { "Yes"} else { "No" } })
            write-host $selectchoice
        } else {
            Write-Host " "
        }
    }
    return $yes
}

function GetKeystroke ($choices) {
    Start-Sleep -Milliseconds 100
    $host.ui.RawUI.FlushInputBuffer()
    Write-Host $Message " ($choies)" -ForegroundColor White
    $key 
    do {
        
            $getkey = $host.ui.RawUI.ReadKey("NoEcho,IncludeKeyDown")
            $getkey = $getkey.Character.tostring().tolower()
        }
        while ($choices.tolower().notcontains($getkey)) 
    return  $getkey
}

# Define a function that takes a directory path and an optional list of preferred extensions as parameters
function Get-DuplicateFiles($Path) {
    # Initialize an empty list to store the file objects
    $FileList = @()
    $sp = 0
    # Get all the files in the directory and sub-directories recursively
    $Files = Get-ChildItem -Path $Path -Recurse -File
    # Group the files by their name and extension
    $FileGroups = $Files | Group-Object -Property Name
    # Loop through each group of files
    foreach ($FileGroup in $FileGroups) {
        if (($sp++ % 20) -eq 0) {
            ShowProgress   "Checking for duplicate files in different directories"
            }
        # If the group has more than one file, it means there are duplicates
        if ($FileGroup.Count -gt 1) {
            # Sort the files by their directory depth, ascending
            $SortedFiles = $FileGroup.Group | Sort-Object -Property @{Expression = {$_.FullName.Split('\').Count}}
            # Loop through the rest of the files in the group, starting from the second one
            foreach ($File in $SortedFiles[1..($SortedFiles.Count - 1)]) {
                # Compare the file hashes of the first file and the current file
                $FirstFileHash = Get-FileHash -Path $SortedFiles[0].FullName
                $FileHash = Get-FileHash -Path $File.FullName
                # If the hashes are equal, it means the files have the same content
                if ($FirstFileHash.Hash -eq $FileHash.Hash) {
                    # Add the current file's System.IO.FileInfo object to the list of duplicates
                    $FileList += $File 
                }
            }
        }
    }
 
    
    # Return the list of duplicate files
    return $FileList
}

function Get-DuplicateFileNames($Path, $PreferredExtensions = @()) {
    # Initialize an empty list to store the file objects
    $FileList = @()
    $sp = 0
    # Get all the files in the directory and sub-directories recursively
    $Files = Get-ChildItem -Path $Path -Recurse -File
    
    # If the preferred extensions list is not empty, check for duplicate names with different extensions
    if ($PreferredExtensions.Count -gt 0) {
        # Group the files by their base name (without extension)
        $NameGroups = $Files | Group-Object -Property BaseName
        # Loop through each group of files
        foreach ($NameGroup in $NameGroups) {
            if (($sp++ % 20) -eq 0) {
                ShowProgress   "Checking For unneeded formats"
                }
    
            # If the group has more than one file, it means there are duplicates
            if ($NameGroup.Count -gt 1) {
                # Sort the files by their extension, using the preferred extensions list as the order
                
                $SortedFiles = $NameGroup.Group | Sort-Object -Property @{Expression = {(&{if ($PreferredExtensions.IndexOf($_.Extension) -ne -1) { $PreferredExtensions.IndexOf($_.Extension) } else {100} })}; Descending = $false}
                # Loop through the rest of the files in the group, starting from the second one
                foreach ($File in $SortedFiles[1..($SortedFiles.Count - 1)]) {
                    # Add the current file's System.IO.FileInfo object to the list of duplicates
                    $FileList += $File 
                }
            }
        }
    }
    # Return the list of duplicate named files with different types
    return $FileList
}

function CheckAndRemove($RemoveFiles) {

    $fcr = $RemoveFiles.length
    if ($fcr  -gt 0) {
        write-host "Found ${fcr} files that should be removed to clean up extras" -ForegroundColor Yellow
        # $RemoveFiles|select Name, DirectoryName, Extension | Out-GridView -Title "Files that will be removed if you press Space (Close this Windows to continue to see prompt)" 
        $cont = MyPause 'Remove those files? (No to keep them)'  -Choice $true -BoxMsg 'Click Yes to remove them' -ChoiceDefault $false

        if ($cont) {
            if (!$HardDelete -and $fcr  -gt 100) {
                $cont = (MyPause 'This is going to take a while as it moves the files to recycle, you will not able able to use your computer.  Would you like to Delete the file without being able to recover them?'  $true 'Click Yes to for a quick delete with NO Recyle!') 
                if ($cont) {
                    $HardDelete = $true
                    Write-Host "Switching to Fast quick delete without recycle" -ForegroundColor Yellow
                    }
                }
            $howDeleted = 'Recycling '
            if ($HardDelete) {
                $howDeleted = 'Deleting '
                }
            $fcs = 0
            ForEach ($f in $RemoveFiles) {
                RecycleFile ($f.FullName)
                LogAction -File $f.Name -Action "--Remove file"
                ShowProgress  ($howDeleted  + "extra files from cache") "$fcs of $fcr - $($f.Name)"
                $fcs++
                }
            }
        }
    }


# Define a recursive function to traverse directories
Function TailRecursion {
    param (
        [string]$Path,
        [int]$Depth = 0
    )

    $IsFound = $false

    # Recursively call the function for each child directory
    Get-ChildItem -Force -LiteralPath $Path -Directory | ForEach-Object {
        TailRecursion -Path $_.FullName -Depth ($Depth + 1) | Out-Null
    }

    # Check if the current directory is empty
    $IsEmpty = -not (Get-ChildItem -Force -LiteralPath $Path)

    # If the directory is empty and it's not the top directory, remove it
    if ($IsEmpty -and $Depth -gt 0) {
        Write-Verbose "Removing empty folder: '$Path'"
        RecycleFile $Path
        $IsFound = $true

        $Script:DirCount++
        ShowProgress "Removing Directory" $Path
        LogAction -File $Path -Action "--Remove Directory"
    }

    return $IsFound
}

#-----------------------------------------------------------------------
# Only clear/reset the New files Directory once and only if there are new files with this run
# Let's becareful that we are not clearing the wrong directory
function ChecktoClearNewFilesDirectory {
    if ($Script:clearNewFiles) {
        if ((get-volume -filePath  $NewFilesDir).DriveType -eq "Fixed") {
            Get-ChildItem -Path $NewFilesDir -Recurse | Remove-Item -Force -Recurse
            write-verbose "CLEARED Copy File Space"
        }
        $Script:clearNewFiles = $false
    }
}

function FetchImageFile ([string]$source, [string]$destination) {
    if (-not (Test-Path $source)) {
        Write-Host "[+] Downloading file from '$source'"
        try {
            (New-Object System.Net.WebClient).DownloadFile($source, $destination)
        } catch {
            Write-Host "`t[!] Failed to download '$source'"
            Write-Host "`t[!] $_"
        }
    } else {
        Write-Verbose "[+] Using existing file."
    }
}


function makeAES {
    # create aes key - keep this secure at all times
    $aesKey = @()
    $primer = ($Env:UserName + $Env:ComputerName + $env:USERPROFILE + $env:OS + $env:Path)

    for ($i = 0; $i -lt 16; $i++) {
        $a = [byte]$primer[$i*2]
        $b = [byte]$primer[$i*2+1]
        $aeskey += [byte](($a-32) % 16 + (($b-32)*16) % 256)
    }
    return $aesKey
}

function EncryptStr ($plaintext) {
    $aesKey = makeAES
    $s1 = ConvertTo-SecureString -String $plaintext -AsPlainText -Force
    return Convertfrom-SecureString -SecureString $s1 -Key $aesKey 
}

function DecryptStr ($secureObject) {
    if ($secureObject) {
        $aesKey = makeAES
        $s2 = Convertto-SecureString -String $secureObject -Key $aesKey 
        $decrypted = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($s2)
        return [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($decrypted)
    }
    return ""
}



#=======================================================================
#  Cloud API MySewnet interface reversed by Darren Jeffrey
#=======================================================================
$Script:tokens = $null
$Script:authorize = $null

<#
.SYNOPSIS
    This function logs into the SewnetCloud.

.DESCRIPTION
    The Login-SewnetCloud function sends a POST request to the SewnetCloud API to authenticate a user. 
    It takes a username and password as parameters, and if the authentication is successful, it stores the session token in a global variable.

.PARAMETER username
    The username of the SewnetCloud account.

.PARAMETER pass
    The password of the SewnetCloud account.

.EXAMPLE
    Login-SewnetCloud -username "user@example.com" -pass "password"

.OUTPUTS
    Boolean. Returns $true if the login is successful, $false otherwise.

.NOTES
    The function sets the global variable $Script:authorize to the session token if the login is successful.
#>
function Login-SewnetCloud($username, $pass)
{
    $requestUri = "https://api.mysewnet.com/api/v2/accounts/login"
    $headers = @{
        "Origin"="https://www.mysewnet.com"
        "Referer"="https://www.mysewnet.com"
        "Cookie"="epslanguage=en; country=US; svpCulture=en-US;"
        "Accept-Encoding"="gzip, deflate, br"
        "Accept-Language"="en-US,en;q=0.9"
        "Content-Type"="application/json"
        "Sec-Fetch-Dest"="empty"
        "Sec-Fetch-Mode"="cors"
        "Sec-Fetch-Site"="same-site"
        "User-Agent"="EmbroideryCollection Manager Cleanup $ECCVERSION"
    }
    
    $body = @{
        "Email"=$username
        "Password"=$pass
    } | ConvertTo-Json
    ShowProgress "Logginning onto MySewNet"
    try {
        $Script:tokens = Invoke-RestMethod -Uri $requestUri -Method POST -Headers $headers -Body $body 
    } catch {
        # Note that value__ is not a typo.
        # Write-Host "StatusCode:" $_.Exception.Response.StatusCode.value__ 
        # Write-Host "StatusDescription:" $_.Exception.Response.StatusDescription
        # write-host "Error: " $_ -ForegroundColor Red
        $errorPack = $_ | convertfrom-json
        write-host "Error: " $errorPack.message -ForegroundColor Red
    } 
    # A Valid tokens contains
    # $tokens.session_token
    # $tokens.encrypted_user_id
    # $token.refresh_token
    # $token.expires_in
    # $token.refresh_expires_in 
    # $token.profile


    if ($null -eq $Script:tokens.session_token) {
        Write-error "Authentication Failed" 
        return $false
    } 

    Set-Content "Token.txt" $($Script:tokens | ConvertTo-Json)
    $Script:authorize = $Script:tokens.session_token

    return $true
}
 
Function refreshToken() 
{
$refreshTokenParams = @{
    client_id=$clientId;
    client_secret=$clientSecret;
    refresh_token=$refreshToken;
    grant_type="refresh_token"; # Fixed value
  }
  
  $Script:tokens = Invoke-RestMethod -Uri $requestUri -Method POST -Body $refreshTokenParams

}

<#
.SYNOPSIS
    This function generates authorization header values.

.DESCRIPTION
    The authHeaderValues function returns a hashtable of HTTP headers for authorization. 
    It uses the global variable $Script:authorize to set the "Authorization" header.

.EXAMPLE
    $headers = authHeaderValues

.OUTPUTS
    Hashtable. Returns a hashtable of HTTP headers for authorization.

.NOTES
    The function uses the global variable $Script:authorize to set the "Authorization" header.
#>

function authHeaderValues ()
{
    return @{ 
        "Authorization" = "Bearer " + $Script:authorize
        "Accept"="application/json, text/plain, */*"
        "Accept-Encoding"="gzip, deflate, br"
        "Accept-Language"="en-US,en;q=0.9"
        "Origin"="https://www.mysewnet.com"
        "Referer"="https://www.mysewnet.com"
        "Pragma"= "no-cache"
        "Cache-Control"="no-cache"
        "User-Agent"="EmbroideryCollection Manager Cleanup $ECCVERSION"
    }
}
<#
.SYNOPSIS
    This function reads metadata from the SewnetCloud.

.DESCRIPTION
    The ReadCloudMeta function sends a GET request to the SewnetCloud API to retrieve metadata. 
    It uses the authHeaderValues function to get the authorization headers. 
    If the request fails, it retries up to 3 times before giving up.

.EXAMPLE
    $metadata = ReadCloudMeta

.OUTPUTS
    Object. Returns the metadata object if the request is successful, $null otherwise (Error occured).

.NOTES
    The function sets the global variable $global:webresultpath to the result if the request is successful.
#>
Function ReadCloudMeta() 
{
    
    $requestUri = "https://api.mysewnet.com/api/v2/cloud/metadata"

    $authHeader = authHeaderValues
    $result = $null
    $tries = 0 
    while ($null -eq $result -and $tries -lt 3) {
        ShowProgress "Reading file list from MySewNet (attempt: $tries)"
        try {
            $result = Invoke-RestMethod -Headers $authHeader -Uri $requestUri -Method GET -ContentType 'application/json'
        } catch {
            Start-Sleep -seconds 2
        }
        $tries++
        
    }
    
    if ($null -eq $result){
        write-host "MySewNet error, please try again later" -ForegroundColor Red
    } else {
        # Recurse the structure and add a Path attribute to the collection with the directory names
        CloudMetaAddPath "" $result
        $global:webresultpath = $result
    }
    return $result
}

<#
.SYNOPSIS
    This function adds a path to each folder in the given metafolder.

.DESCRIPTION
    The CloudMetaAddPath function iterates over each folder in the given metafolder and adds a 'path' property to it. 
    The value of the 'path' property is the concatenation of the given path and the name of the folder. 
    If the given path is empty, the value of the 'path' property is the name of the folder. 
    The function is recursive, calling itself for each folder in the metafolder.

.PARAMETER path
    The path to be added to each folder in the metafolder.

.PARAMETER metafolder
    The metafolder to which the path is to be added. The default value is $webcollection.

.EXAMPLE
    CloudMetaAddPath -path "\Users\username\Documents" -metafolder $webcollection

.OUTPUTS
    None. This function does not return a value; it modifies the given metafolder in-place.

.NOTES
    The function uses the global variable $webcollection as the default value for the metafolder parameter.
#>
function CloudMetaAddPath ($path, $metafolder = $webcollection) {
    foreach ($fid in $metafolder.folders) {
        $pathHere = if ($path -eq "") { "\" + $fid.name } else { Join-Path -Path $path -ChildPath $fid.name }
        Add-Member -InputObject $fid -MemberType NoteProperty -Name 'path' -Value $pathHere
        CloudMetaAddPath -path $pathHere -metafolder $fid
    }
}


Function GetFileIDfromCloud ($fileNameExt, $metafolder = $webcollection)
{
    if ($metafolder) {
        if ($fileNameExt ) {
            if ($metafolder.files | where-object { $_.name -like $fileNameExt}) {
                return $metafolder.files | where-object { $_.name -like $fileNameExt}
            } else {
                foreach($fid in $metafolder.folders) {
                    $retdir = GetFileIDfromCloud $fileNameExt  $fid 
                    if ($null -ne $retdir) {
                        return $retdir
                    }
                }
            }
        }
    }
    return $null
}

function findMetaDirectory($folderid, $metafolder = $webcollection) {
    if ($folderid) {
        if ($metafolder) {
            if ($metafolder.folders | where-object { $_.id -like $folderid}) {
                return $metafolder.folders | where-object { $_.id -like $folderid}
            } else {
                foreach($fid in $metafolder.folders) {
                    $retdir = findMetaDirectory -folderid $folderid -metafolder $fid 
                    if ($null -ne $retdir) {
                        return $retdir
                    }
                }
            }
        }
        return $null
    } else {
        return $webcollection
    }
    
}

function GetCloudDirectoryNameFromId($folderid) {
    if ($folderid) {
        $nameCollect = findMetaDirectory -folderid $folderid
        $parent = GetCloudDirectoryNameFromId -folderid $nameCollect.parentFolderId
        return $parent + "\" + $nameCollect.name   
    }
    return ""
}

<#
Return a CloudRef which contains id for a path value of foldername  (example \folder1\f2\f3 )
#>
function FindCloudidfromPath($foldername, $metafolder = $webcollection) {
    if ($metafolder) {
        if ($metafolder.folders | where-object { $_.path -like $foldername}) {
            # write-host "Meta _Path" ($metafolder.folders | where-object { $_.path -like $foldername})
            return $metafolder.folders | where-object { $_.path -like $foldername}
        } else {
            foreach($fid in $metafolder.folders) {
                $retdir = FindCloudidfromPath $foldername  $fid 
                if ($retdir) {
                    # write-host "Ret $retdir"
                    return $retdir
                }
            }
        }
    }
    return $null
    
}
function GetCloudDirectoryid($foldername) {
    # Give a folder name find the folder id #
    $pack = FindCloudidfromPath -foldername $foldername
    if ($pack) {
        # $path | fl 
        return $pack.id
    } else {
        return ""
    }
}

Function DeleteCloudFolder ($id)
{
    $authHeader = authHeaderValues
    $myError = $null    
    $requestUri = 'https://api.mysewnet.com/api/v2/cloud/folders/'+ $id;
    write-host "$requestUri"
    try {        
	        $result = Invoke-RestMethod -Uri $requestUri -Method "DELETE" -Headers $authHeader -verbose
	} catch {
		# Note that value__ is not a typo.
		write-host "StatusCode:"  $_.Exception.Response.StatusCode.value__
		write-warning ("StatusDescription:" + $_.Exception.Response.StatusDescription)
		$result = ""
		$myError = $_
	} 


	if ($myError) {
		$eDetails = ($myError.errorDetails.Message|convertfrom-json )
		write-warning ("Error Deleting folder id: '$id'  [ " + $eDetails.message + " ]")
	}
    else {
        # TODO Manage the copy of the shadow structure for Cloud files
    }
    return $result
}
Function DeleteCloudFile ($id)
{
    $authHeader = authHeaderValues
    $myError = $null   
    $requestUri = 'https://api.mysewnet.com/api/v2/cloud/files/'+ $id;
    try {        
	        $result = Invoke-RestMethod -Uri $requestUri -Method "DELETE" -Headers $authHeader 
	} catch {
		# Note that value__ is not a typo.
		write-host "StatusCode:"  $_.Exception.Response.StatusCode.value__
		write-warning ("StatusDescription:" + $_.Exception.Response.StatusDescription)
		$result = ""
		$myError = $_
	} 

	if ($myError) {
		$eDetails = ($myError.errorDetails.Message|convertfrom-json )
		write-warning ("Error Deleting file id: '$id'  [ "  + $eDetails.message + " ]" )
	}
    else {
        # TODO Manage shadow copy structure for Cloud folders
        $filefind = $webcollection | Select-Object -ExpandProperty Folders | select-Object -ExpandProperty Files| Where-Object {$_.id -eq $id} 
        $parentref = findMetaDirectory -folderid $filefind.folderId
        if ($parentref) {
            $parentref.files = $parentref.files | where-Object { $_.id -ne $id}
        }
    }
    
}

Function MoveCloudFile ($fileid, $toFolderid)
{
    $authHeader = authHeaderValues
    $myError = $null   
    $requestUri = "https://api.mysewnet.com/api/v2/cloud/files/$fileid/move"
    if ($toFolderid) {
        $bodyLines = @{
            "newFolderId" = "$toFolderid"
            } | ConvertTo-Json     
    } else {
        # Special case for root - future Bug?
        $bodyLines = "{}"
    }
    try {        
            $result = Invoke-RestMethod -Uri $requestUri -Method "PUT" -Headers $authHeader -ContentType "application/json" -Body $bodyLines
    } catch {
            # Note that value__ is not a typo.
            write-host "StatusCode:"  $_.Exception.Response.StatusCode.value__
            write-warning ("StatusDescription:" + $_.Exception.Response.StatusDescription)
            $result = ""
            $myError = $_
        } 

        if ($myError) {
            $eDetails = $myError
            write-warning ("Error Moving file id: '$fileid'  [ "  + $eDetails + " ]" )
        }
        else {
            # Update Cache
        }
    
}


<#
Create a Folder $name within Folder ID $inFolderID
#>
Function CreateCloudFolder($name, $inFolderID)
{
    if ($name.substring(0,1) -eq '\') {
        $name = $name.substring(1)
    }
    $CheckName = findMetaDirectory -folderid $inFolderID
    if ($null -ne $CheckName) {
        if ($CheckName.Folders.name -like $name) {
            $fld = $CheckName.Folders | where-object { $_.name -like $name }
            write-host "Folder " $fld.Name " ($name) exists as " $fld.id " for in " $fld.parentfolderid " ($inFoldID) "
            return $fld.id
        }
    }
    $authHeader = authHeaderValues
    $requestUri = 'https://api.mysewnet.com/api/v2/cloud/folders';
    $bodyLines = @{ 
        "folderName" = $name
        "parentFolderId" = $inFolderID
        } | ConvertTo-Json

	try {        
	        $result = Invoke-RestMethod -Uri $requestUri -Method POST -Headers $authHeader -ContentType "application/json" -Body $bodyLines
	} catch {
		# Note that value__ is not a typo.
		write-host "StatusCode:"  $_.Exception.Response.StatusCode.value__
		write-warning ("StatusDescription:" + $_.Exception.Response.StatusDescription)
		$result = ""
		$myError = $_
	} 

	if ("" -eq $result) {
		$eDetails = ($myError.errorDetails.Message|convertfrom-json )
		write-warning ("Error Pushing the folder: '$name'  [ " + $eDetails.message + " ]" )
	} else {
		# write-host "Folder '$name' saved as '$result'"
        $parentref = findMetaDirectory -folderid $inFolderID
        $now = get-date -Format "yyyy-MM-ddTHH:mm:ss.fffffffzzz"
        $parentref.folders += [PSCustomObject]@{ 
            id = $result
            name = $name
            lastChanged = $now
            files = @()
            folders = @()
            parentFolderId = $inFolderID
            path = $parentref.path + "\" + $name
            }

	}
    
    return $result

}

Function MakeCloudPathID($path){
    $folderid = FindCloudidfromPath $path
    if ($folderid -or ($path -eq "")) {
        return $folderid.id
    }
    $paths = $path.split("\")
    
    $buildpath = ""
    foreach ($nextpath in $paths) {
        if ($nextpath) {
            $buildpath += "\" + $nextpath
            # This is a BUG as FindCloud should not recursive search
            $pathid = FindCloudidfromPath $buildpath
            if ($pathid) {
                $lastpathid = $pathid.id
            } else {
                $lastpathid = CreateCloudFolder -name $nextpath -inFolderID $lastpathid
                # Create a Delay (Deal with BUG later TODO)
                # start-sleep -Seconds 1
                # TODO ERROR / BUG on their side
                if ($lastpathid -eq "") {
                    Write-Warning "Problem creating folders $buildpath"
                    return ""
                }
            }
        } else {
            $lastpathid = ""
        }
    }
    return $lastpathid
}
function PushCloudFileToDirectory($filepath, $folderpath ) 
{
    $filename = Split-Path -Path $filepath -leaf
    $folderid = MakeCloudPathID -path $folderpath
    return (PushCloudFile -name $filename -inFolderID $folderid -filepath $filepath)
}
Function PushCloudFile($name, $inFolderID, $filepath)
{
    if (test-Path -Path $filepath) {
        $CheckFolder = findMetaDirectory -folderid $inFolderID
        if ($null -ne $CheckFolder) {
            $fld = $CheckFolder.Files | where-object { $_.name -like $name }
            if ($fld) {
                Write-Verbose "File $($fld.Name) ($name) exists as $($fld.id) in $($fld.parentfolderid) ($inFoldID) "
                return $fld.id
            }
        }
        $authHeader = authHeaderValues
        $requestUri = 'https://api.mysewnet.com/api/v2/cloud/files';
        
        $fileBytes = [System.IO.File]::ReadAllBytes($filepath);
        $fileEnc = [System.Text.Encoding]::GetEncoding('UTF-8').GetString($fileBytes);
        $boundary = "------" + [System.Guid]::NewGuid().ToString(); 
        $LF = "`r`n";
        
        $bodyLines = ( 
            "--$boundary",
            "Content-Disposition: form-data; name=`"File`"; filename=`"blob`"",
            "Content-Type: application/octet-stream$LF",
            $fileEnc,
            "--$boundary",
            "Content-Disposition: form-data; name=`"FileName`"$LF",
	        $name,
            "--$boundary",
            "Content-Disposition: form-data; name=`"FolderId`"$LF",
             $inFolderID,
	        "--$boundary--"
            ) -join $LF

	try {        
	        $result = Invoke-RestMethod -Uri $requestUri -Method POST -Headers $authHeader -ContentType "multipart/form-data; boundary=`"$boundary`"" -Body $bodyLines
	} catch {
		# Note that value__ is not a typo.
		write-host "StatusCode:"  $_.Exception.Response.StatusCode.value__
		write-warning ("StatusDescription:" + $_.Exception.Response.StatusDescription)
		$result = ""
		$myError = $_
	} 

	if ("" -eq $result) {
		$eDetails = ($myError.errorDetails.Message|convertfrom-json )
		write-warning ("Error Pushing the file: '$name'  [ "  + $eDetails.message +" ]") 
	} else {
		# write-host "File '$name' saved as '$result'"
        $parentref = findMetaDirectory -folderid $inFolderID
        $now = get-date -Format "yyyy-MM-ddTHH:mm:ss.fffffffzzz"
        $size = get-item $filepath
        $size = $size.Length
        $parentref.files += [PSCustomObject]@{ 
            id = $result
            folderId = $inFolderID
            size = $size
            name = $name
            lastChanged = $now
            }

	}
        return $result
    } else {
        write-Warning "File Not found for upload : $filepath"
    }

}
#----------------------------------------------
#
# API Interface to query files from the cloud
#
#----------------------------------------------
function doWebAPI($userid, $pw) {

    if (([string]$userid -eq "") -or ([string]$pw -eq "") -or $PromptPassword) {
        
        $userid = read-host -Prompt "What is your User id for MySewnet: "
        if ($userid -eq "") {
            write-host "user id is required - stopping" -ForegroundColor Yellow
            return ("", "")
        }
        $dpw = read-host -Prompt "password for MySewnet: "
        $pw = EncryptStr $dpw
        
        if ($dpw -eq "") {
            write-host "MySewNet password is required - stopping" -ForegroundColor Yellow
            return ("", "")
        }
        write-host "Testing Authenication" -ForegroundColor Blue
        $dpw = DecryptStr $pw
        if (!(Login-SewnetCloud -username $userid -pass $dpw )) {
            write-host "--- stopping ---" -ForegroundColor Yellow
            return ("", "")
        }
    }

    if (!$Script:authorize) {
        write-host "Using username: '$userid' to access MySewnet" -ForegroundColor Green 
        $dpw = DecryptStr $MySewNetpw
        if (!(Login-SewnetCloud -username $userid -pass $dpw )) {
            write-host "If you continue to experience login issues, update the password in $ConfigFile"
            write-host "--- stopping ---" -ForegroundColor Yellow
            return ("", "")
        }
    }
    if ($PromptPassword) {
        $pw = ""
    }
    return ($userid, $pw)
    
}

#-----------------------------------------------------------------------
# Move From Directory to either Embrodery or Instrctions directory
function MoveFromDir ( 
        [string] $fromPath, 
        [boolean]$isEmbrodery = $false,        # is this to the Embrodery directory (true) Or to the Instruction Directory (false)
        [string]$whichfiles = "*.*",            # File Names
        [string[]]$files = $null                # File extension types as an array
        
    ) 
{
    $newFileCount = 0
    ShowProgress "Copying" "Added ${Global:savecnt} files"
    if ($isEmbrodery) { 
        $dtype = 'Embroidery' 
        $objs = Get-ChildItem -Path $fromPath -include $SewTypeStar -File -Recurse -filter $whichfiles
        $targetdir = $EmbroidDir
    } else { 
        # Move anything that is not a Embrodery type file (alltypes)
        $dtype = 'Instructions'
        $objs = Get-ChildItem -Path $fromPath -exclude ($excludetypes +$TandCs)  -File -Recurse -filter $whichfiles
        $targetdir = $InstructDir
        }

    $objs | ForEach-Object {
        if (($null -eq $files) -or ($_.Name -in $files)) {                
            $newdir = (Split-Path(($_.FullName).substring(($fromPath.Length), 
                                ($_.FullName).Length - ($fromPath.Length) )))
            
            $newfile = $_.Name
        
            # take off the directory name if it is one of the rollup names
            do { 
                $folding = $false
                foreach ($r in $foldupDirs) {
                    if ($newdir.ToLower().EndsWith("\"+$r)) {
                        #strip off the directory name and perserve the case of the directory and files
                        $newdir = $newdir.substring(0,($newdir.tolower().Replace("\"+$r,'')).length)
                        $folding = $true
                        }
                    }
                } while ($folding)
            
            $newpath = join-path -path $targetdir -childpath $newdir
            if (!(Test-Path -Path $newpath -PathType Container)) {
                New-Item -Path $newpath -ItemType Directory | Out-Null
                }
            $npath = (Join-Path -Path $newpath -ChildPath $newfile )
            if (test-path $npath) {  # See if the file already exists
                
                $newHash = Get-FileHash($_.FullName)
                $orgHash = Get-FileHash($npath)
                if ($orgHash.Hash -eq $newHash.Hash) {
                    Remove-Item -Path $_
                    Write-Verbose "Removed Duplicate ${dtype} file :'$_'" 
                    }
                
                Write-Verbose "Skipping ${dtype}:'$_' to ${newdir}" 
                }
            if (test-path $_) {
                if (test-path $npath) { Remove-item -path $npath -force -ErrorAction  SilentlyContinue }
                    ChecktoClearNewFilesDirectory
                    if ($NoDirectory) {
                        Copy-Item -Path $_ -Destination (Join-Path -Path $NewFilesDir  -ChildPath $newfile)
                    } else {
                        $newpath = join-path -path $NewFilesDir -childpath $newdir
                        if (!(test-path ($newpath))) {
                            New-Item -Path ($newpath) -ItemType Directory  | Out-Null
                            }
                        Copy-Item -Path $_ -Destination (Join-Path -Path $newpath -ChildPath $newfile)
                    }
                    LogAction $newfile
                    Move-Item $_ -Destination $npath  # -ErrorAction SilentlyContinue
                    $newFileCount += 1
                    Write-Information "+++ Saving ${dtype}:'$_' to ${newdir} & ${newfiledir}"
                    if ($isEmbrodery) { 
                        $Script:addsizecnt = $Script:addsizecnt + (Get-Item -Path $npath).Length 
                        $Script:savecnt = $Script:savecnt + 1
                        }
                    }
        } else {
                Write-Verbose "Skipping ${_.Name}" 
            }
    
    ShowProgress "Copying" -Status "Added ${Global:savecnt} files"
    
    }
    return $newFileCount
}

#-----------------------------------------------------------------------
# Format a Size string in xB
#
function NiceSize ($size) {
    $units = " B", " KB", " MB", " GB", " TB"
    $unitIndex = 0

    while ($size -gt 1024 -and $unitIndex -lt $units.Length) {
        $size /= 1024
        $unitIndex++
    }

    return "{0:N1}{1}" -f $size, $units[$unitIndex]
}
#======================================================================================
<#
    $mysewingfiles = Array of Object: 
                        C:\Dir\File.txt
        NameIndexed                 : either N or Base used for matching with or without extension match
        N               File.txt    : filename with extension
        Base            File        : filename without extension
        Ext             txt         : Extension
        DirectoryName   C:\Dir\     : full path Directory name only
        Hash                        : hash value of the file calculated when we need it
        DirWBase        C:\Dir\File : Directory name with Base 
        Priority                    : sorted type based on preference
        CloudRef                    : See object below
        Push                        : Path if new file else null

    CloudRef = 
        folders: Array 
            id              08dc0292-2c14-4944-8384-42ff54f32e85; 
            name            folder1; 
            lastChanged     2023-12-22T02:03:20.779584+00:00;
            parentFolderId  08dc02ad-d255-49b8-8837-184824b72c48  (or "")
            files           = CloudRef
            folders         = CloudRef
            path            \folder1

        files: Array
            id              0865a538-3f84-4fc6-9c5c-4c7d2d967c25; 
            folderId        08dc02ad-d255-49b8-8837-184824b72c48    (or "")
            size            35061; 
            name            X13347A.VP3;
            lastChanged     2023-10-18T03:45:07.099194+00:00

#>
Function LoadSewfiles  {
    
    $thelist = (Get-ChildItem -Path $EmbroidDir  -Recurse -file -include $SewTypeStar)| ForEach-Object { 
        if ($keepAllTypes) {
            $n = $_.Name} 
        else {
            $n = $_.BaseName
        }
        [PSCustomObject]@{                          # C:\Dir\File.txt
            NameIndexed = $n                               
            N = $_.Name                                         # File.txt
            Base = $_.BaseName                                  # File
            Ext = $_.Extension                                  # txt
            DirectoryName = $_.DirectoryName                    # C:\Dir\
            Hash = $null                                        # hash value of the file calculated when we need it
            DirWBase = $_.DirectoryName + "\" + $_.BaseName     # C:\Dir\File
            Priority = $preferredSewType.Indexof($_.Extension.substring(1,$_.Extension.Length-1).tolower())
            RelPath = '?????'
            CloudRef = $null                                    #
            Push = $null
            } 
        }
    if ($null -eq $thelist) {
        $thelist =    @([PSCustomObject]@{ 
                NameIndexed ="mysewingfiles.placeholder"
                N = "mysewingfiles.placeholder"
                Base  = "mysewingfiles"
                Ext  = "placeholder"
                DirectoryName = "DirectoryName"
                DirWBase = "DirWBase"
                Priority = 100
                RelPath = '?????'
                Hash = $null # We will calculate it when we need it
                CloudRef = $null  
                Push = $null
                })
        }
    return $theList
}
#======================================================================================

function ShowPreferences ($showall = $false)
{
    if ($showall) {
        foreach ($paramselect in ($paramstring.Keys)) {
            if ($($paramstring[$paramselect])) {
                $val = Get-Variable -Name $paramselect -ValueOnly
                Write-host $($paramstring[$paramselect]).padright($padder+20) ": " $val
            }
        }
        foreach ($paramselect in ($parambool.Keys)) {
            if ($($parambool[$paramselect])) {
                $val = Get-Variable -Name $paramselect -ValueOnly
                Write-host $($parambool[$paramselect]).padright($padder+20) ": " $val
            }
        }
        foreach ($paramselect in ($paramarray.Keys)) {
            if ($($paramarray[$paramselect])) {
                $val = (Get-Variable -Name $paramselect -ValueOnly ) -join '", "'
                Write-host $($paramarray[$paramselect]).padright($padder+20) ": " """$val"""
            }
        }
    } else {

    Write-Host    "Download source directory".padright($padder)  ": $downloaddir" 
    Write-host    $($paramstring['EmbroidDir']).padright($padder)  ": $EmbroidDir" 
    Write-host    $($paramarray['preferredSewType']).padright($padder)  ": $preferredSewType"
    if ($FirstRun) {
    Write-host    "Processing all files in Download directory".padright($padder) 
    } else {
    Write-host    $($paramstring['DownloadDaysOld']).padright($padder)  ": $DownloadDaysOld"
    }
    Write-host    $($parambool['keepAllTypes']).padright($padder)  ": $keepAllTypes"
    if ($APIBeta) {
        Write-host    $($paramswitch['APIBeta']).padright($padder)   -ForegroundColor yellow
    }
    }
}

function  CheckUSBDrive ($USBPath) {

    $failed = $true
    $driveletter = $USBPath | split-path -Qualifier -ErrorAction SilentlyContinue
    # Check if the drive letter is valid
    if ($driveletter -match "^[A-Z]:" -or ($driveletter.legth -eq 1 -and $driveletter -match "[A-Z]")) {
      $driveletter = $driveletter[0]
      # Get the volume object for the drive letter
      $volume = Get-Volume -DriveLetter $driveletter -ErrorAction SilentlyContinue

      # Check if the volume exists
      if ($volume) {
        # Check if the volume is a USB device
        if ($volume.DriveType -eq "Removable") {
           $failed = $False
        }
        else {
          Write-Host "The drive ${driveletter}: is not a removable device." -ForegroundColor Red
        }
      }
      else {
        Write-Host "The drive ${driveletter}: does not exist." -ForegroundColor Red
      }
      
    } elseif ($USBPath.tolower().contains("off")) {
        $USBDrive = ""
        $failed = $false
        $driveletter = ""
    } else {
      Write-Host "The drive letter is invalid ('$USBPath')." -ForegroundColor Red
    }
    if ($failed) {
        return ""
    } else {
        return $driveletter
    }
    
}

#======================================================================================
#
# Building out all the directory structures and File lists
#
#======================================================================================
$SewTypeStar = $preferredSewType | ForEach-Object { "*.$_" }
$SewTypeMatch = $preferredSewType -join '|'
$foldupDirs = $foldupDir + $preferredSewType | ForEach-Object { $_.ToLower() }

$excludetypes = $alltypes | ForEach-Object {
    $thistype = $_
    if ($preferredSewType -notcontains $thistype ) {
        "*.$thistype "
    }
}



if (-not $doit){
    $PSDefaultParameterValues = @{
  "Copy-Item:WhatIf"=$True
  "Move-Item:WhatIf"=$True
  "Remove-Item:WhatIf"=$True
}
}


if (-not $EmbroidDir.contains("\")) {
    $EmbroidDir = join-path -path $docsdir -childpath $EmbroidDir 
}
$InstructDir = $EmbroidDir
# if (!( $InstructDir.contains("\"))) {
#    $InstructDir = join-path -path $docsdirwOne -childpath $InstructDir 
#}

# TODO test path for exists

$LogFile = join-path $PSScriptRoot "EmbroideryCollection-Cleanup.Log"
if (!(test-path $LogFile)) {
    "EmbroideryCollection-Cleanup\n" | Set-Content -Path $LogFile 
}



if ($setup) {
    write-host "   ".padright(70) -BackgroundColor Yellow -ForegroundColor Black
    $Desktop = [Environment]::GetFolderPath("Desktop")
    if (!(test-path ($Desktop + "\Embroidery Organizator.lnk"))) {
        write-host "  Creating shortcut on the Desktop".padright(70) -BackgroundColor Yellow -ForegroundColor Black
        $WshShell = New-Object -comObject WScript.Shell
        $Desktop = $Desktop + "\Embroidery Organizator.lnk"
        write-Debug "Link: $Desktop"
        $Shortcut = $WshShell.CreateShortcut($Desktop)

        $Shortcut.TargetPath = "$pshome\Powershell.exe"
    
        
        $Shortcut.IconLocation = (join-path -Path $PSScriptRoot -childpath "embroiderymanager.ico")
        $Shortcut.Arguments = "-NoLogo -ExecutionPolicy Bypass -File ""$PSCommandPath"""
        $Shortcut.Description = "Run EmbroideryCollection-Cleanup.ps1 to extract the patterns from the download directory"
        $Shortcut.Save()
        LogAction -File "Desktop Shortcut" -Action "Created"
        }
    ShowPreferences
#TODO NEED TO FIX THIS BLOCK
    # Load the System.Windows.Forms assembly
    Add-Type -AssemblyName System.Windows.Forms

    # Instantiate a FolderBrowserDialog object
    $DirectoryBrowser = New-Object System.Windows.Forms.FolderBrowserDialog -Property @{ 
        SelectedPath = $EmbroidDir
        Description = "Select the Directory for Embroidery Files"
        ShowNewFolderButton = $true
        }

    do {
        write-host "In order to Setup you will be asked questions to configure this script".padright(70)   -BackgroundColor Blue -ForegroundColor White
        write-host $($paramstring['EmbroidDir']) "?" -NoNewline
        # Show the dialog box and store the selected folder path
        if ($DirectoryBrowser.ShowDialog() -eq "OK") {
            $EmbroidDir = $DirectoryBrowser.SelectedPath
        } else {
            $continuesetup = $false
        }
        write-host $EmbroidDir
        write-host "How do you want to transfer your files to your machine (USB, Mysewnet or neither)"
        if (myPause -Message "Are you using a USB Drive?" -Choice $true -ChoiceDefault ($USBDrive -ne "")) {
            do {
                $USBDrive = Read-Host -Prompt "Which Drive is the USB stick connected too?"
                if ($USBDrive -eq "") {
                    $notvalid = myPause -Message "Do you still want to use a USB Drive?" -Choice $true
                } else {
                    $udrive = CheckUSBDrive -USBPath $USBDrive
                    $notvalid = ($udrive -eq "")
                    }
            } while ($notvalid)
            $APIBeta = $false
        } else {
            $USBDrive = ""
        }
        if ($USBDrive -eq "") {
            if (myPause -Message "Are you using MySewnet Cloud" -Choice $true -ChoiceDefault $APIBeta) {
                $APIBeta = $true
                $USBDrive = ""
                $meid = Read-host -Prompt "What is your user name for MySewnet? ($MySewNetuserid)"
                if ($meid) { $MySewNetuserid = $meid }
                write-host "You will be asked for your password the first time you run the script again" -ForegroundColor Yellow
                $savep = myPause -Message "Are you okay with saving an encrypted copy of that password for future logins" -Choice $true -ChoiceDefault (!$PromptPassword)
                $PromptPassword = -not $savep
            } else {
                $APIBeta = $false
                $USBDrive = ""
            }
        }
        $dd = Read-Host "How many days back do you want to always look when checking the Download folder? (currently $DownloadDaysOld)"
        if ($dd -gt 0) {
            $DownloadDaysOld = $dd
        }         
        $KeepAllTypes = myPause $parambool['KeepAllTypes'] -Choice $true -ChoiceDefault $KeepAllTypes
        $KeepEmptyDirectory = myPause $parambool['KeepEmptyDirectory'] -Choice $true -ChoiceDefault $KeepEmptyDirectory
        if ($APIBeta) {
            $DragUpload = myPause $parambool['DragUpload'] -Choice $true -ChoiceDefault $DragUpload
            $ShowExample =myPause $parambool['ShowExample'] -Choice $true -ChoiceDefault $ShowExample
            }
        
#        $NoDirectory = myPause $parambool['NoDirectory'] -Choice $true -ChoiceDefault $NoDirectory
#        $OneDirectory = myPause $parambool['OneDirectory'] -Choice $true -ChoiceDefault $OneDirectory
        $val = $alltypes  -join ', '
        Write-host "All the different Embroidery file types: `n$val" 
        
        
        Write-host "What are the preferred types of files for your machine in order of preference"
        write-host "Current list is: $preferredSewType"
        do {
            $ptype = Read-Host "Files types (seperated by comma)?" 
            if ($ptype) {
                $ptype = ($ptype.split(',')).trim()
                $ptype = $ptype |Where-Object {$_.length -gt 0}
                $problemext = $ptype |Where-Object {($_.contains(' ')) -or ($_.Length -gt 4)}
                if ($problemext) {
                    write-host "Problem with the extension of: $problemext" -ForegroundColor Red
                } else {
                    $preferredSewType = $ptype
                }
            }
        } while ($problemext)
        write-host "  All Settings ".padright(70) -BackgroundColor Blue -ForegroundColor White
        ShowPreferences -showall $true
        $savep = mypause -Message "Do you want to save these settings?" -Choice $true
        $continuesetup = -not $savep
    } while ($continuesetup)
    SaveAllParams

    
##########    
    if (-not ((Test-Path($EmbroidDir)) -and (Test-Path($instructDir)))) {
        write-host "  Creating Directory in Documents for Embroidery and Instructions  ".padright(70) -BackgroundColor Yellow -ForegroundColor Black
        
        if (!(Test-Path($EmbroidDir))) {
            New-Item -ItemType Directory -Path $EmbroidDir | out-null
            write-host "Creating Directory '$EmbroidDir' for Embroidery files" -BackgroundColor Green -ForegroundColor Black
            LogAction -File $EmbroidDir -Action "Created Directory"
        }
        if (!(Test-Path($InstructDir))) {
            New-Item -ItemType Directory -Path $instructDir | out-null
            write-host "Creating Directory '$instructDir' for Instructions files" -BackgroundColor Green -ForegroundColor Black
            LogAction -File $instructDir -Action "Created Directory"
        }
    }
    write-host "All Done " -BackgroundColor Yellow -ForegroundColor Black
    MyPause 'Press any key to Close' | out-null
    Return
}
Write-Host " ".padright(5) "Let's begin managing the Embroidery files".padright(80) -ForegroundColor white -BackgroundColor blue
if ($FirstRun) {
    Write-Host " ".padright(15) $("Checking ALL Zip files ".padright(70)) -ForegroundColor white -BackgroundColor blue
}

# Clean out the old tmp working space
Get-ChildItem -Path  ($tmpdir ) -Recurse | Remove-Item -force -Recurse
if (-not (Test-Path -Path $tmpdir )) { New-Item -ItemType Directory -Path ($tmpdir )}

$failed = $false

# TODO change for USBDrive
# We will put all the new files in here for now
 $UsingUSBDrive = $False
 
if ("" -ne $USBDrive) {
    
    $driveletter = CheckUSBDrive $USBDrive
    if ("" -ne $driveletter) {
        Write-Warning  $driveletter
        $NewFilesDir = $driveletter + ":\"
        # Don't wipe someone's USB drive
        $Script:clearNewFiles = $False
        $DragUpload = $False
        $ShowExample = $false
        $UsingUSBDrive = $True
    }

    
} else {
    $Script:clearNewFiles = $true
    $NewFilesDir = ${env:temp} + "\cleansew.new"
    if (-not (Test-Path -Path $NewFilesDir )) { New-Item -ItemType Directory -Path ($NewFilesDir )}

}

$flatdir = -not $noDirectory
ShowPreferences 

if ($DragUpload) {
    write-host "Using Web Drag & Drop option" -ForegroundColor Green
}

if ($UsingUSBDrive) {
    Write-host    "Saving to USB drive".padright($padder)  ": $NewFilesDir" -ForegroundColor Yellow
}
if ($Testing) {
    Write-Host "Testing Mode".padright($padder) + ": $Testing" -ForegroundColor Yellow
    }
Write-Verbose ("Rollup match pattern".padright($padder-8) + ": $foldupDirs")
Write-Verbose ("Ignore Terms Conditions files".padright($padder-8) + ": $TandCs")
Write-Verbose ("Excludetypes".padright($padder-8) + ": $excludetypes")

if ($APIBeta) {
    ($MySewNetuserid, $MySewNetpw) = doWebAPI $MySewNetuserid $MySewNetpw
    if (!$Script:authorize) {
        return
    }
}

SaveAllParams
if (( $EmbroidDir.tolower().contains("\onedrive") )) {
    Write-Host "The Embroidery files directory '$EmbroidDir' is within OneDrive ---- Warning" -ForegroundColor Yellow
    
    }
    
if (!( test-path -Path $EmbroidDir -PathType Container)) {
    Write-Host "Can not find the main directory $EmbroidDir.  ---- Stopping" -ForegroundColor Red
    Write-Host "Usually create '$EmbroidDir' in the with '$docsdir' directory.  Create the directory if this is your first time." 
    $failed = $true
    }
if (!( test-path -Path $InstructDir -PathType Container)) {
    Write-Host "Can not find the 'Instruction' Directory '$InstructDir'.  ---- Stopping" -ForegroundColor Red
    Write-Host "Usually created in the Documents directory ($docsdir).  Create the directory if this is your first time"  -ForegroundColor Yellow
    $failed = $true
    } 
if ($missingSewnetAddin) {
    write-host "** Install the Explorer Plug-in from https://download.mysewnet.com/MSW/ so the pattern images appear in Windows Explorer" -ForegroundColor Yellow
    }

if ($failed) {
    Write-host " The problem above needs to be corrected".padright(80) -BackgroundColor DarkRed -ForegroundColor White
    write-host "See instructions at https://github.com/D-Jeffrey/Embroidery-File-Organize"
    $failed = MyPause 'Press any key to Close' 

    break
    }


$cont = (MyPause 'Press Start to continue, any other key to stop (Auto starting in 3 seconds)'  $true 'Click Yes to start' 3) 

if (!$cont) { 
    Break
    }
Add-Type -assembly "system.io.compression.filesystem"

ShowProgress  "Calculating size"
$librarySizeBefore = 0
Get-ChildItem -Path ($EmbroidDir)  -Recurse -file  | ForEach-Object { $librarySizeBefore = $librarySizeBefore + $_.Length}

ShowProgress  "Loading file list"
$mysewingfiles = $null
# Get a list of all the existing files in mySewnet

$mysewingfiles = LoadSewfiles

if ($FirstRun) {
    $DownloadDaysOld = 20*365
}
# $mysewingfiles | ft
$isNewInstruct = $false
Get-ChildItem -Path $downloaddir  -file -filter "*.zip" | Where-Object { ($_.CreationTime -gt (Get-Date).AddDays(- $DownloadDaysOld ))  -and ($_.gettype().Name -eq 'FileInfo')} |
  
    ForEach-Object {
        $thisfileBase = $_.BaseName
        ShowProgress  "Checking Zips - Currently looking at $($_.Name)"  "Added $Script:savecnt files"
        $zips = $_.FullName
        Write-Verbose "Checking ZIP '$zips'" 
        $filelist = [io.compression.zipfile]::OpenRead($zips).Entries

        $madeDir = ""
        $isNewInstruct = $false
        $isnew = $false
        $numnew = 0
        foreach ($t in $preferredSewType) {
            $filesOfThisType = @()
            $ts = "*."+ $t
            if ($filelist.Name -like $ts) {
                $isnew = $false
                $myfl = $filelist | where-object {$_.Name -like $ts} 
                foreach ($filesinZip in $myfl) {
                    $filenameInZip = $filesinZip.Name
                    if ($filenameInZip -match "\.") {
                        $fs = $filenameInZip.split('\.')[0] 
                    } else {
                        $fs = $filenameInZip
                    }
                    if (-not $keepAllTypes) {
                        $filenameInZip = $fs
                        }
                    if ($filenameInZip -in ($mysewingfiles.NameIndexed)) {
                        # Check for duplicate FileHash
                        # Find the instances then get the full name then get hash to compare
                        $fileInstance = $mysewingfiles | where-object {$_.NameIndexed -eq $filenameInZip}
                        $fiName = $fileInstance.DirWBase +  $fileInstance.Ext
                        
                        Write-verbose "Duplicate file '${filenameInZip}' to ${fiName}"
                    } else {
                        Write-verbose "New file '${filenameInZip}'"
                        $isnew  = $true
                        if ($keepAllTypes) {
                            $filesOfThisType += $filenameInZip
                        } else {
                            $filesOfThisType += ($filenameInZip + "." + $t)
                        }
                        if ($keepAllTypes) {
                           $n = $filesinZip.Name 
                        } else {
                            $n = $fs
                            }

            # TODO BUG This does not work # -and -not $NoDirectory 
            # without changing MoveFromDir  and normalizing NoDirectory everywhere
                        if ($filesinZip.FullName.LastIndexOf('/') -ge 0 ) {
                            $relativepath = $filesinZip.FullName.Substring(0,$filesinZip.FullName.LastIndexOf('/'))
                        } else {
                            $relativepath = ""
                        }
                        $dirn = (join-path -Path $EmbroidDir -ChildPath $relativepath).trim('\')
                        
                        $mysewingfiles +=  
                            [PSCustomObject]@{ 
                                NameIndexed = $n
                                N = $filesinZip.Name
                                Ext = "." + $t
                                Base = $fs
                                DirectoryName = $dirn
                                Hash = $null
                                DirWBase =  join-path -Path $dirn -ChildPath $fs
                                Priority = $preferredSewType.Indexof($t.tolower())
                                RelPath = '?????'
                                CloudRef = $null
                                Push = '\'+ $relativepath
                                } 
                            
                        }
                    }
                
                # we found a new file in the Zip.  If we have not expanded this Zip, then do it now
                if ($isnew) { 
                    if (-not $isNewInstruct) { 
                        if ($filesOfThisType.count -gt $SetSize) {     # This is a set we should keep together
                            $madeDir = "\" + $thisfileBase -replace ' \(([0-9]+)\)'
                            
                            }
                        $resultTmpDir = Join-Path $tmpdir $madeDir
                        $isNewInstruct = $true
                        Expand-Archive -Path $zips -DestinationPath $resultTmpDir -Force
                        }
                  
                    $numnew += $(MoveFromDir $tmpdir $true $ts $filesOfThisType)
                    }
                }               
            }
            # now lets check to see if there was a ZIP file in a Zip file there
            # BUG zip in zip with expand files also does not work
            $tempziplist = @()
            $filelist | ForEach-Object {
                if ($_.FullName -match ".zip" -and ($_.gettype().Name -eq 'FileInfo')) {
                    $NeedToExpandedZip = $true
                    Write-Host "- - Found: nested zip file $($_.FullName) checking"
                    $tempFile = [System.IO.Path]::GetTempFileName() + ".zip"
                    $tempziplist += $tempFile
                    # TODO we might be duplicating the extraction of the zip files if we expanded above 
                    # Need Error trap for 'InvalidDataException' for corrupted files
                    [System.IO.Compression.ZipFileExtensions]::ExtractToFile($_, $tempFile, $true)
                    #BUG Exception calling "OpenRead" with "1" argument(s): "End of Central Directory record could not be found."
                    [io.compression.zipfile]::OpenRead($tempFile).Entries | foreach-object {
                        if ($_.FullName -match $SewTypeMatch) {
                            $thisname = split-path $_.FullName  -leaf
                            
                            #TODO this needs to be a new path (put it into the same directory where of the orginal extract of other files
                            if ($NeedToExpandedZip) {
                                $NeedToExpandedZip = $false
                                Expand-Archive -Path $tempFile -DestinationPath $tmpdir -Force
                                }
                                $numnew += $(MoveFromDir $tmpdir $true $thisname) 
                            
                        }
                    
                    }
                }
            }
                    

        $zf = $zips.tolower().replace($downloaddir.tolower(),'...')
        if ($numnew -gt 0) {
            Write-host $("* New  : '$zf'").padright(65) " $numnew new patterns" 
        } else {
            Write-host $("- Found: '$zf'").padright(65) " nothing new" 
            }
        # we extracted the Zip already and now let's check for instructions
        if ($isNewInstruct) { 
            $numnew += $(MoveFromDir $tmpdir $false)
            Get-ChildItem -Path $tmpdir -Recurse | Remove-Item -force -Recurse
            }
   
        $filelist = $null      # Close Zipfile
        ShowProgress  "Checking Zips"  "Added $Script:savecnt files"

    }
    $tempFile = $null      # Close Zipfile
# $mysewingfiles | ft


# Look for Files which are not part of a ZIP file, just the selected file types that we are looking for that is in the download directory
$DownloadDaysOld = 365*10  # 10 years of downloads (when you download files, it keeps the old data)
$ppp = 0
foreach ($t in $preferredSewType) {

    $ts = "*."+ $t
    write-Information "Working on File type: $ts"
    Get-ChildItem -Path $downloaddir  -file -include $ts -Depth 1 -Recurse| Where-Object { $_.CreationTime -gt (Get-Date).AddDays(- $DownloadDaysOld ) } |
        ForEach-Object {
            $thisfile = $_
            $f = $_.Name
            $fs = $_.BaseName
            $d = $_.DirectoryName
            $fullname = $_.FullName
            
            if ($keepAllTypes) {
                $findfile = $f
            } else {
                $findfile = $fs
                }
            if ($findfile -in $mysewingfiles.NameIndexed) {
                Write-verbose "Duplicate file '${findfile}'"
            } else {
                if (test-path -path $fullname) { 
                    write-Information "checking on $fullname"
                    ChecktoClearNewFilesDirectory
                    $fullname | Copy-Item -Destination $EmbroidDir  -ErrorAction SilentlyContinue
                    if ($NoDirectory) {
                        $fullname | Copy-Item -Destination (join-path $NewFilesDir $f) -ErrorAction SilentlyContinue
                    } else {
                        $fullname | Copy-Item -Destination $NewFilesDir  -ErrorAction SilentlyContinue
                    }
                    LogAction $f
                    
                    Write-host $("* New  : '$f'").padright(65) " 1 new pattern" 
                    Write-Information "+++ Copied from Downloads :'$($_.Name)' to $EmbroidDir"
                    $fd = join-Path -Path $d -ChildPath ($fs + "." + $t)
                    if (test-path -path $fd) {
                        Copy-Item -Path $fd -Destination $InstructDir -ErrorAction SilentlyContinue 
                        Write-Information "+++ Copied instructions from Downloads :'$($_.Name)' to $InstructDir"
                        LogAction -File $($_.Name), -isInstrution $true
                        }
                    $Script:addsizecnt = $Script:addsizecnt + $_.Length 
                    $Script:savecnt = $Script:savecnt + 1 
                    if ($NoDirectory) {
                        $l = ""
                    } else {
                            $l = (split-path -Path $fullname -Parent).Substring($downloaddir.Length).trim('\')
                    }
                    $d = (join-path -path $EmbroidDir -childpath $l).Trim('\')
                    $prefno = $preferredSewType.Indexof($t.tolower())
                    $mysewingfiles +=  [PSCustomObject]@{ 
                                NameIndexed = $f
                                N = $f
                                Base = $fs
                                Ext = '.' + $t
                                DirectoryName = $d
                                DirWBase = join-path -path $d -childpath $fs
                                Priority = $prefno
                                RelPath = '?????'
                                Hash = $null
                                CloudRef = $null
                                Push = '\' + $l
                    }
                }    
            }
        $ppp++
        if ($ppp % 5 -eq 0) {
            ShowProgress  "Copying from Downloads" 
            }
        }    
    }


    # clean up the zip file mess
    foreach ($tz in $tempziplist) {
        remove-item -Path $tz 
    }
    
# $mysewingfiles | ft

# TODO check for BUGS ?? 

if ($CleanCollection -or $CleanCollectionIgnoreDir) {
    write-host "Scanning for files to clean up in cache"

    # this creates a BUG for below as we are removing files that might be synced without removeing them from the list
    $filesToRemove = DuplicateFiles -Path $EmbroidDir  
    CheckAndRemove -RemoveFiles $filesToRemove
    if (!$KeepAllTypes) {
        $pe = $preferredSewType | ForEach-Object { ".$_" }
        $filesToRemove = DuplicateFileNames -Path $EmbroidDir  -PreferredExtensions  $pe
        CheckAndRemove -RemoveFiles $filesToRemove
        }
    
    write-host "Moving Instructions to selected 'Instructions' directory"
    foreach ($g in $goodInstructionTypes) {
        $numnew += $(MoveFromDir $EmbroidDir $false ("*."+ $g))        
        }
}    


#  Clear out empty Directories
if (-not $KeepEmptyDirectory) {
    $tailr = 0    # Loop thru 8 times to remove empty directories, then go back and check to see if you made any more emty
    while ($tailr -le 8 -and (tailRecursion $EmbroidDir) ) {
         $tailr++
    }
}


if ($APIBeta) {
    $webcollection = ReadCloudMeta
    if ($null -eq $webcollection) {
        write-host "Cloud is not working *** STOPPING" -ForegroundColor Red
        MyPause 'Press any key to Close' | out-null
        return
    
    }
    if ($Testing) {
        write-host "Working on " -ForegroundColor Blue
        $mysewingfiles[0]
        $f = $mysewingfiles[0].DirWBase + $mysewingfiles[0].Ext
        PushCloudFile -name $mysewingfiles[0].N -inFolderID "" -filepath $f
        write-host "Working on " - -ForegroundColor Blue
        $mysewingfiles[1]
        $f = $mysewingfiles[1].DirWBase +  $mysewingfiles[1].Ext
        $d = $mysewingfiles[1].DirectoryName.substring($EmbroidDir.Length)
        $did = MakeCloudPathID -path $d 
        PushCloudFile -name $mysewingfiles[1].N -inFolderID $did.id -filepath $f
        MakeCloudPathID -path "\folder1\folder2\folder3"
        MakeCloudPathID -path "\AE 3D Christmas Gift Tags\Anita's Express - 3D Christmas Gift Tags"
        MakeCloudPathID -path "\Inspirograph Quil\ABC"
        MakeCloudPathID -path ""
        MakeCloudPathID -path "\f1\f2\f3\f4\f1"
        write-host "Done " -ForegroundColor Blue
    }
    if (0) {
        $mysewingfiles[0]
        $f = $mysewingfiles[0].DirWBase + $mysewingfiles[0].Ext

        $did = MakeCloudPathID -path "\3D Christmas Gift Tags\Anita's Express - 3D Christmas Gift Tags"
        PushCloudFile -name $mysewingfiles[0].N -inFolderID $did -filepath $f
        write-host "Created Folder '\3D Christmas Gift Tags' and pushed files"
        MyPause "Press Key to continue and delete"
        $did = FindCloudidfromPath -foldername '\3D Christmas Gift Tags'
        write-host "Deleting Folder $did"
        DeleteCloudFolder $did.id

    }

    $cm = 0
    $cf = 0
    $tolist = @()
    $MySewingfiles | ForEach-Object {
        $thisfile = $_
        $cm = $cm + 1
        if ( $cm % 25 -eq 0) {
            ShowProgress ("Matching files to Cloud ($cm) : " + $thisfile.N)
        }
        $thisfile.CloudRef = GetFileIDfromCloud $_.N
        if ($thisfile.DirectoryName.length -lt $EmbroidDir.Length) {
            $thisfile.DirectoryName
        }
        $spl = $thisfile.DirectoryName.substring($EmbroidDir.Length)
        $thisfile.RelPath = $spl
        $samepath = FindCloudidfromPath -foldername $spl

        if ($thisfile.CloudRef)   {
            $cf = $cf +1
            $tolist += [PSCustomObject]@{ 
                Name = $thisfile.N
                Path = $spl
                ShouldBeid = $samepath.id 
                CurrentFolderID = $thisfile.CloudRef.FolderId
                NeedtoMove = (""+$samepath.id -ne ""+$thisfile.CloudRef.FolderId)
            }
        } else {
            $tolist += [PSCustomObject]@{ 
                Name = $thisfile.N
                Path = $spl
                ShouldBeid = $samepath.id 
                CurrentFolderID = ""
                NeedtoMove = ""
            }
            # write "?? API file for " $_.N
            }
        }
        $tmf = $MySewingfiles.count
    write-host "Found $cf file which match in cloud of $tmf"
#    $tolist | Out-GridView



    if ($sync) {
        $webcollection | Select-Object -ExpandProperty Folders | select-Object -ExpandProperty Files| Where-Object {$_.Name -notin ($mysewingfiles.N)} | ForEach-Object  {
            write-host " --Removing " $_.Name
            LogAction -File $_.Name -Action "Deleted-Sync"
            DeleteCloudFile -id $_.id
        }
        $filestopush = ($mysewingfiles | Where-Object { ($_.Push -and $_.Push.contains('\'))  -or ($_.CloudRef -eq $null)}).count

        }
     else {
        $filestopush = ($mysewingfiles | Where-Object { $_.Push -ne $null}).count
    }

    if (-not $KeepEmptyDirectory) {
        $emtpyfoldersList = $webcollection |   Select-Object -ExpandProperty Folders | where-object {$_.folders.count -eq 0 -and $_.files.count -eq 0 } 
        if ($emtpyfoldersList.count -gt 0) {
            write-host "Clearing Empty Cloud folders : $($emtpyfoldersList.count)"
        }
        $emtpyfoldersList | ForEach-Object { 
            ShowProgress "Delete Cloud Folder $($_.Name)"
                DeleteCloudFolder -id $_.id | Out-null
        }
        
    }
    write-host "Beginning push to MySewnet: $filestopush files" -ForegroundColor Green
    $i = 0
    if ($filestopush) {
        $MySewingfiles | ForEach-Object {
            $thisfile = $_
            if ($thisfile.push -and $thisfile.push.contains("\") ) {
                Write-Progress -PercentComplete ($i++*100/$filestopush) "Pushing files to the Cloud : " -Status $thisfile.N
                PushCloudFileToDirectory -filepath ($thisfile.DirWBase + $thisfile.Ext) -folderpath $thisfile.push | Out-Null
                LogAction -File $thisfile.N -Action "Added"
                $thisfile.push = ""
                }
            if ($sync -and ($null -eq $thisfile.CloudRef)) {
                # TODO This Progress should be %
                Write-Progress -PercentComplete ($i++*100/$filestopush) "Syncing files to the Cloud : " -Status $thisfile.N
                $spl = $thisfile.DirectoryName.substring($EmbroidDir.Length)
                PushCloudFileToDirectory -filepath ($thisfile.DirWBase + $thisfile.Ext) -folderpath $spl | Out-null
                LogAction -File $thisfile.N -Action "Added-Sync"
                }
            }
        }
    }
    if ($sync) {
        $MySewingfiles | ForEach-Object {
            $thisfile = $_
            if ($thisfile.DirectoryName.length -lt $EmbroidDir.Length) {
                $thisfile.DirectoryName
            }
            $samepath = FindCloudidfromPath -foldername $thisfile.RelPath
            if ($thisfile.CloudRef)   {
                # Need to move
                if  (""+$samepath.id -ne ""+$thisfile.CloudRef.FolderId) {
                    ShowProgress -stat "Moving Cloud files" -Area $thisfile.N
                    LogAction -File $thisfile.N -Action "Moved CloudFolders"
 # BUG we need to moka directory if samepath errors                   
#                    $folderid = MakeCloudPathID -path $folderpath

                    MoveCloudFile -fileid $thisfile.CloudRef.Id -toFolderid $samepath.id
                }
            }
        }
    }
    


Function OpenForUpload {
    
    Write-Host "-----------------------------------------------------------------------------------------" -ForegroundColor Green
        
    if ($DragUpload) {
        Write-Host "Opening File Explorer & MySewnet Web page" -ForegroundColor Green
        Write-Host " ** on MySewNet web page choose 'Upload' and Select all files in Explorer and " -ForegroundColor Green
        Write-Host "    drag/drop the files a maximum of 5 at a time into the upload box" -ForegroundColor Green
        
    } else {
        if ((Get-WmiObject -class Win32_OperatingSystem).Caption -match "Windows 11") {
                Write-Host "Opening File Explorer (using mysewnet add-in)" -ForegroundColor Green
                Write-Host " ***  Select all files *right-click* and choose 'Show more Options' -> choose 'MySewNet' -> 'Send'" -ForegroundColor Green
        } else {
            # Assume it is Windows 10 with add-in
            Write-Host "Opening File Explorer (using mysewnet add-in)" -ForegroundColor Green
            Write-Host " ***  Select all files *right-click* and choose 'MySewNet' -> 'Send'" -ForegroundColor Green
            }
        
    }
    $firstfile = $(get-childitem -path $NewFilesDir -File -depth 1)
    if ($firstfile.count -gt 0) {
        $firstfile = $firstfile[0].FullName
        $explorercmd = "explorer  '/select,""$firstfile""'"
        } 
    else { 
        Write-Host " There are NO Files to upload" -ForegroundColor Yellow
        $firstfile = $NewFilesDir + "\."
        $explorercmd = "explorer  '$NewFilesDir'"
    }
    Write-Host "-----------------------------------------------------------------------------------------" -ForegroundColor Green
    
    if ($DragUpload) { 
        Start-Process $opencloudpage 
        }
    Invoke-expression  $explorercmd

    if (-not $DragUpload -and $ShowExample) { 
        $file = Join-Path -path $(Split-Path -path $PSCommandPath) -ChildPath 'HowToSend-w10.gif'
        FetchImageFile $file "https://raw.githubusercontent.com/D-Jeffrey/Embroidery-File-Organize/main/HowToSend-w10.gif"
        if (test-path $file) {
            write-host "Opening Example (Close it by clicking on the 'X' in the top right corner)"
            Add-Type -AssemblyName 'System.Windows.Forms'
            $file = (get-item $file)
            $img = [System.Drawing.Image]::Fromfile((get-item $file))

            [System.Windows.Forms.Application]::EnableVisualStyles()
            $form = new-object Windows.Forms.Form   
            $form.Text = "How to Do it"
            $form.Width = $img.Size.Width;
            $form.Height =  $img.Size.Height;
            $pictureBox = new-object Windows.Forms.PictureBox
            $pictureBox.Width =  $img.Size.Width;
            $pictureBox.Height =  $img.Size.Height;

            $pictureBox.Image = $img;
            $form.controls.add($pictureBox)
            $form.Add_Shown( { $form.Activate() } )
            $form.ShowDialog()
            $form.Close()

        } else {
                write-host "Error: Could not find example file : $file" -foregroundcolor red
        }
    }
}

write-host "Calculating size"
ShowProgress  "Calculating size"
$librarySizeAfter = 0
if ($Script:savecnt -gt 0) {
    if ($null -eq $APIBeta) {
        OpenForUpload
    }
    Get-ChildItem -Path ($EmbroidDir  ) -Recurse -file  | ForEach-Object { $librarySizeAfter = $librarySizeAfter + $_.Length}
} else {
    if ($null -eq $APIBeta) {
        if (MyPause "Do you want to open the web page & directory to upload files from last time?" $true) {
            OpenForUpload
        }
    }
}


  
$librarySizeBefore = niceSize $librarySizeBefore
$librarySizeAfter = niceSize $librarySizeAfter
$sizecntB = niceSize  $sizecnt
$addsizecntB = niceSize $Script:addsizecnt
write-progress -PercentComplete  100  "Done"
if ($Script:dircnt -gt 0 -or $filesToRemove.length -gt 0) {
    $filecnt = $filesToRemove.length
    Write-Host "Cleaned up - Directories removed: '$Script:dircnt    Files removed : '$filecnt' ($sizecntB)." -ForegroundColor Green
    }
if ($Script:savecnt -gt 0) {
    write-host "+++ Added files to ${CollectionTypeofStr}: '${Global:savecnt}' ($addsizecntB) " -ForegroundColor Green
    Write-host "   *** Instructions size is now : $librarySizeAfter was $librarySizeBefore ****   "  -ForegroundColor Green 
    }
else {
    Write-host "   *** Instructions size is : $librarySizeBefore ****   "  -ForegroundColor Green 
}
if ($APIBeta) {
    write-Host "Cloud Storage currently is: " -ForegroundColor Green
    write-Host "       Used of Total:".padright(25) (NiceSize $webcollection.storage.usedSize) "/" (NiceSize $webcollection.storage.totalSize)
    write-Host "       Space remaining:".padright(25) (NiceSize $webcollection.storage.availableSize)
}
# $mysewingfiles | out-GridView
# MyPause 'Press any key to Close' | out-null
Write-Host ( "End") -ForegroundColor Green
