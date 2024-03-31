#Requires -Version 5.1
<# 
 EmbroideryCollection-Cleanup.ps1
    GPL-3.0 license

 Deal with the many different types of embroidery files, put the right format types in mySewingnet Cloud or onto a USB drive
 We are looking to keep the ??? files types only from the zip files.

 Orginal Author: Darren Jeffrey, Dec 2021 - Mar 2024
#>

param
(
  [Parameter(Mandatory = $false)]
  [int32]$DownloadDaysOld = 7,                      # How many days old show it scan for Zip files in Download
  [int32]$SetSize = 10,
  [Switch]$KeepAllTypes,                            # Keep all the different types of a file (duplicate name but different extensions)
  [Switch]$CleanCollection,                         # Cleanup the Collection Folder to only EmbroidDir files
# BUG in NoDirectory
#  [Switch]$NoDirectory,                             # Do not create directory structure in the upload from space
#  [Switch]$OneDirectory,                              #Limit the folders to one directly deep only 
  [string]$EmbroidDir = "Embroidery",               # You may want to change this directory name inside of the 'Collection ' Directory wihtin your 'Documents' directory
                                                    # If it is just a name, then it is assumed to be within the defeault Documents directory, otherwise it will be taken as a full path
  [string]$USBDrive,                                # Write the output to a USB Drives
  [Switch]$HardDelete,                              # Delete the files rather than sending to recycle bin
  [switch]$KeepEmptyDirectory,                      # If you don't want this to remove extra empty directories from Collection folders'
  [Switch]$Testing,                                 # Run it and see what happens
  [Switch]$Setup,                                   # Setup the Shortcut on the desktop to this application
  [Switch]$DragUpload,                              # Use the web page instead of the plug in to drag and drop
  [Switch]$ShowExample,                             # Show the example GIF
  [string]$ConfigFile = "EmbroideryCollection.cfg", # This is the file in the same directory as this script otherwise it is a full path (Not recommened to change)
  [Switch]$ConfigDefault,                           # Got back to default settings
  [Switch]$SwitchDefault,                           # Clear all the preview Switch enabled Values
  [Switch]$FirstRun,                                # Scan all the ZIP files
  [Switch]$Sync,                                    # Sync MySewnet to local folders
  [Switch]$CloudAPI                                  # use MySewNet cloud API
  )

$ECCVERSION = "v0.6.7"

# $VerbosePreference =  "Continue"
# $InformationPreference =  "Continue"

<# ********     CONFIGURATION    ********#>
$preferredSewType = 'vp3', 'vip', 'pcs','dst', 'pes', 'hus'
$alltypes = 'hus', 'dst', 'exp', 'jef', 'pes', 'vip', 'vp3', 'xxx', 'sew', 'vp4', 'pcs', 'vf3', 'csd', 'zsk', 'emd' , 'ese', 'phc', 'art', 'ofm', 'shv', 'pxf', 'svg', 'dxf', 'pec', 'pcm', 'pxf', 'dem', 'phc', 'mhv', 'bmc'
$foldupDir = 'images', 'sewing helps', 'Designs', 'Design Files', 'brother-babylock-pes', 'janome-jef', 'singer-xxx', 'husqvarna-viking-hus', 'commercial formats - dst-exp', 'artista-art'

$goodInstructionTypes = ('pdf','doc', 'docx', 'txt','rtf', 'mp4', 'ppt', 'pptx', 'gif', 'jpg', 'png', 'bmp','mov', 'wmv', 'avi','mpg', 'm4v', 'htm', 'html' )
$TandCs = @('TERMS-OF-USAGE.*', 'planetappliquetermsandconditions.*','READ ME FIRST.rtf','*copyright.*','*copyright Statement.*','*copyrights.*',
    'copyrightStatement.*','License agreement.*', 'License.*','termsofuse.*', 'Thumbs.db')

<# ********     END CONFIGURATION  - active settings kept in EmbroideryCollection.cfg  ********#>

# List of paramstring to check
$paramstring =  [ordered]@{
 "EmbroidDir" = "Embriodary Files directory";
 "USBDrive"="USB drive letter (example E: or H:)";
 "LastCheckedGithub"="";
 "DownloadDaysOld" = "Age of files in Download directory";
 "SetSize" = "Keep collections of files together if there are at least this many"
}

$parambool = [ordered]@{
'KeepAllTypes'= 'Keep all variations of files types' ; 
'KeepEmptyDirectory'= 'When cleaning up keep empty folders'; 
'DragUpload'= 'Open the mysewnet Cloud browser interface for drag and drop';
'ShowExample'= 'Show how to upload to mySewnet';
'NoDirectory'= 'Do not use Directories from Zip files which creating collection';
'OneDirectory'= 'Keep files a maximum of one directory deep ';
'CloudAPI'= 'Use MySewnet Cloud'
}
$paramarray = [ordered]@{
'preferredSewType' = 'The preferred types of Embriodary file types';
'alltypes' = 'All the possible types of files which are an Embriodary file'; 
'foldupDir' = 'Remove/fold folders of this name'; 
'goodInstructionTypes' = 'Instructions file types which should be saved with files' 
}
$paramswitch =[ordered]@{
    'CleanCollection' = 'Clean the Collection folder';
    'CloudAPI' = "Using API to update mySewNet Cloud (It is buggy, try again if you get errors/warnings)";
    'Sync' = 'Syncronize computer folders to Cloud'
}



# ----------------------------------------------------------------------
#                 $alltypes
# this is a list of all the different types of embrodiary files that are considered. The '$preferredSewType' should be from the list below based on what is best for your 
# machine and in the order that you prefer.  If there are more than one copy of a file type it will select your first one
# ----------------------------------------------------------------------
#                 $TandCs              
# Term and Conditions added by various store that add up space with the same document type over and over, using up your MySewing Cloud space
# This is a file name pattern so TC.* will match TC.doc or TC.pdf
# ----------------------------------------------------------------------
#                  $foldupDir
# What directories should be flattened to bring the Embroidery files higher up so they are not nested instead of sub-folders.  
# The names are for Directories you want to remove the sub-folder and moved the contents up
# ----------------------------------------------------------------------

write-host " ".padright(15) "Embroidery Collection Cleanup version: $ECCVERSION on PS $($PSVersionTable.PSVersion.major).$($PSVersionTable.PSVersion.minor)".padright(70) -ForegroundColor White -BackgroundColor Blue

$RemovePrefix = ($PSVersionTable.PSVersion.Major -lt 7 ) 
$filecnt = 0
$script:sizecnt = 0
$Script:dircnt = 0
$Script:savecnt = 0
$Script:addsizecnt = 0
$Script:p = 0
$padder = 45
$use7zipsize = 1024*1024*100    # 100 MB switch to 7zip if it is install for zip files over 100 MB, if 7zip is installed
$filesToRemove = @()
$script:lostfiles = @()
$script:CloudStatusGood = $true

$shell = New-Object -ComObject 'Shell.Application'
$downloaddir = (New-Object -ComObject Shell.Application).NameSpace('shell:Downloads').Self.Path
if (!(test-path $downloaddir)) {
    Write-Error "The Download Directory does not work, please correct the script"
    return
}
$docsdir =[environment]::getfolderpath("mydocuments")
if ($docsdir.tolower().contains('\onedrive')) {
    $docsdir = ${env:HOMEDRIVE} + ${env:HOMEPATH}
}
$tmpdir = ${env:temp} + "\cleansew.tmp"

$opencloudpage = "https://www.mysewnet.com/en-us/my-account/#/cloud/"

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
        $paralist = ($paramstring.Keys)
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
        # toggle a Switch Option 
        if ($isParamPresent -and ($parambool.Contains($param))) {
            Set-Variable -Name $param -Value (-not $SavedParam.$param)
        }
    }
    foreach ($param in ($paramarray.Keys)) {
        if ($null -ne $SavedParam.$param) {
            $newvalue = $SavedParam.$param.value
            if ($null -eq $newvalue) {
                $newvalue = $SavedParam.$param
            }
            Set-Variable -Name $param -Value $newvalue
            
        }
    }
} else {
    $FirstRun = $true
}

class CloudFile {
    # Class properties
    [string]        $Cloudid
    [string]        $FileName
    CloudFile() { $this.Init(@{}) }
    # Convenience constructor from hashtable
    CloudFile([hashtable]$Properties) { $this.Init($Properties) }
    # Common constructor for id and FileName
    CloudFile([string]$Cloudid, [string]$FileName) {
        $this.Init(@{Cloudid = $Cloudid; File = $FileName })
    }
    # Shared initializer method
    [void] Init([hashtable]$Properties) {
        foreach ($Property in $Properties.Keys) {
            $this.$Property = $Properties.$Property
        }
    }
    # Method to return a string representation of the file
    [string] ToString() {
        return "$($this.CloudId) by $($this.FileName)"
    }
}
class SewingFile {
    # Class properties
    [String]                $NameIndexed                            # could be file or file.txt
    [String]                $Name                                      # File.txt
    [String]                $Base                                   # File
    [String]                $DirectoryName                          # C:\Dir\
    [String]                $Hash                                   # hash value of the file calculated when we need it
    [String]                $FullName                               # C:\Dir\File.txt
    [System.IO.FileInfo]    $FileInfo
    [System.DateTime]       $LastWriteTime
    [int32]                 $Priority
    [String]                $RelPath
    [String]                     $CloudRef
    [String]                $Push
    [System.IO.FileInfo]    $TmpPath
    [System.IO.FileInfo]    $KeepPath

     # Default constructor
    SewingFile() { $this.Init(@{}) }
    # Convenience constructor from hashtable
    SewingFile([hashtable]$Properties) { $this.Init($Properties) }
    # Common constructor for id and FileName
    SewingFile([string]$Cloudid, [string]$FileName) {
        $this.Init(@{Cloudid = $Cloudid; File = $FileName })
    }
    # Shared initializer method
    [void] Init([hashtable]$Properties) {
        foreach ($Property in $Properties.Keys) {
            $this.$Property = $Properties.$Property
        }
    }
    # Method to return a string representation of the file
    [string] ToString() {
        return "$($this.relpath + $this.Name)"
    }
    [string] MoveFile() {
        if ($this.TmpPath) {
            copy-item -Path $this.TmpPath.FullName -Destination $this.FileInfo.FullName
            $this.TmpPath = $null
        } elseif ($this.KeepPath) {
            copy-item -Path $this.KeepPath.FullName -Destination $this.FileInfo.FullName
            $this.KeepPath = $null
        } else {
            # Nothing to move
            return ""
        }
        return $this.FileInfo.FullName
    }
}

class SewingFileList {
    # Static property to hold the list of SewingFiles
    static [System.Collections.Generic.List[SewingFile]] $SewingFiles
    # Static method to initialize the list of SewingFiles. Called in the other
    # static methods to avoid needing to explicit initialize the value.
    static [void] Initialize()             { [SewingFileList]::Initialize($false) }
    static [bool] Initialize([bool]$force) {
        if ([SewingFileList]::SewingFiles.Count -gt 0 -and -not $force) {
            return $false
        }

        [SewingFileList]::SewingFiles = [System.Collections.Generic.List[SewingFile]]::new()
        return $true
    }
    # Ensure a SewingFile is valid for the list.
    static [void] Validate([SewingFile]$SewingFile) {
        $Prefix = @(
            'SewingFile validation failed: SewingFile must be defined with the NameIndexed,'
            'FullName, and LastWriteTime properties, but'
        ) -join ' '
        if ($null -eq $SewingFile) { throw "$Prefix was null" }
        if ([string]::IsNullOrEmpty($SewingFile.NameIndexed)) {
            throw "$Prefix NameIndexed wasn't defined"
        }
        if ([string]::IsNullOrEmpty($SewingFile.FullName)) {
            throw "$Prefix FullName wasn't defined"
        }
        if ([datetime]::MinValue -eq $SewingFile.LastWriteTime) {
            throw "$Prefix LastWriteTime wasn't defined"
        }
    }
    # Static methods to manage the list of SewingFiles.
    # Add a SewingFile if it's not already in the list.
    static [void] Add([SewingFile]$SewingFile) {
        [SewingFileList]::Initialize()
        [SewingFileList]::Validate($SewingFile)
        if ([SewingFileList]::SewingFiles.Contains($SewingFile)) {
            throw "SewingFile '$SewingFile' already in list"
        }

        $FindPredicate = {
            param([SewingFile]$s)

            $s.NameIndexed -eq $SewingFile.NameIndexed -and
            $s.AutFullNamehor -eq $SewingFile.FullName -and
            $s.LastWriteTime -eq $SewingFile.LastWriteTime
        }.GetNewClosure()
        if ([SewingFileList]::SewingFiles.Find($FindPredicate)) {
            throw "SewingFile '$SewingFile' already in list"
        }

[SewingFileList]::SewingFiles.Add($SewingFile)
    }
    # Clear the list of SewingFiles.
    static [void] Clear() {
      [SewingFileList]::Initialize()
      [SewingFileList]::SewingFiles.Clear()
    }
    # Find a specific SewingFile using a filtering scriptblock.
    static [SewingFile] Find([scriptblock]$Predicate) {
        [SewingFileList]::Initialize()
        return [SewingFileList]::SewingFiles.Find($Predicate)
    }
    # Find every SewingFile matching the filtering scriptblock.
    static [SewingFile[]] FindAll([scriptblock]$Predicate) {
        [SewingFileList]::Initialize()
        return [SewingFileList]::SewingFiles.FindAll($Predicate)
    }
    # Remove a specific SewingFile.
    static [void] Remove([SewingFile]$SewingFile) {
        [SewingFileList]::Initialize()
        [SewingFileList]::SewingFiles.Remove($SewingFile)
    }
    # Remove a SewingFile by property value.
    static [void] RemoveBy([string]$Property, [string]$Value) {
        [SewingFileList]::Initialize()
        $Index = [SewingFileList]::SewingFiles.FindIndex({
            param($b)
            $b.$Property -eq $Value
        }.GetNewClosure())
        if ($Index -ge 0) {
            [SewingFileList]::SewingFiles.RemoveAt($Index)
        }
    }
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
    $SavedParam | ConvertTo-Json  | Set-Content -Path $ConfigFile -Encoding Utf8
}



if ($missingSewnetAddin) {
    $DragUpload = $true
}
$CloudAuthAvailable = $false
if ($CloudAPI) {
    $DragUpload = $false
    $ShowExample = $false
    $CloudAuthAvailable = ((Get-Module -Name PSAuthClient).count -gt 0)
    if (-not $CloudAuthAvailable) {
        Write-Progress -Activity "Checking for Authenication Module"
        
        if ((Get-Module -ListAvailable | Where-Object { $_.Name -eq "PSAuthClient" }).count -gt 0) {
            Import-Module PSAuthClient 
            $CloudAuthAvailable = ((Get-Module -Name PSAuthClient).count -gt 0)
        }
        else {
            write-host "-- Please wait while the script installs missing module --" -foreground Yellow -NoNewline
            Install-Module -name PSAuthClient -scope:CurrentUser
            Import-Module PSAuthClient 
            $CloudAuthAvailable = ((Get-Module -Name PSAuthClient).count -gt 0)
            write-host "Completed" -foreground Yellow
        }
        Write-Progress -Completed $true
    }
    
}

$doit = !$Testing



# This is for development testing and debugging

if ($env:COMPUTERNAME -eq "DESKTOP-R3PSDBU_") { # -and $Testing) {
    $docsdir = "d:\Users\kjeff\"
    $downloaddir = "d:\Users\kjeff\downloads"
    $doit = $true
    }



    
#=============================================================================================

function LogAction($File, $Action, [Boolean]$isInstructions = $false) {
    $now = Get-Date -Format "yyyy/MMM/dd HH:mm "
    $extra = (&{if ($isInstructions) { " Instructions"} else { "Embroidery" } })
    write-verbose "$Action $File type: $extra"
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
# Delete or recycle a file (full path required)
Function RecycleFile {
    param (
    [System.IO.FileInfo[]]$file, 
    [boolean]$purge )

    if (!($file)) {
        write-warning 'Recycle blank name'
        return
    }
    #BUG-FIX? Long File name detection
    $thisfile = get-item $file
    if ($thisfile.attributes.hasflag([IO.FileAttributes]'Readonly')) {
        $thisfile.attributes -= 'Readonly'
    }
    
    try {
        if ($purge) {
            if ($RemovePrefix) {
                Remove-Item -Path "\\?\$($file.FullName)"        # Handled by WhatIf
            } else {
                Remove-Item -Path $file.FullName        # Handled by WhatIf
            }
        } elseif ($doit) {
            $shell.NameSpace(17).ParseName($file).InvokeVerb('delete')
        }
    } catch {
        write-warning "Problem deleting: $file - $($file.Fullname)"
        
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
        write-progress -Completed $true
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

function Test-ExistsOnPath {
    param (
        [string]$FileName
    )
    $found = $false
    $env:Path.Split([System.IO.Path]::PathSeparator) | where-object {$_ -ne ""} | ForEach-Object {
        $fullPath = Join-Path $_ $FileName
        if (Test-Path $fullPath -PathType Leaf) {
            $found = $true
        }
    }

    return $found
}

# return the relative path of existing folders and files relative to a root path with the prefix of  .\
function RelativeDirectory {
    param (
        [string]$Path,
        [string]$rootPath
        )

    Push-Location -path $rootPath
    $relative = Resolve-Path -Path $path -Relative
    $relative = $relative.trim("\.\\")
    Pop-Location
    return $relative
}

# Look within the directory and find files of the same name and return that list
function DuplicateFiles($Path) {
    # Initialize an empty list to store the file objects
    $FileList = @()
    $sp = 0
    $Excludes = ($allTypesStar + $TandCs )
    ShowProgress   "Get file list for duplicate files in different directories ...."    
    # Get all the files in the directory and sub-directories recursively
    $Files = Get-ChildItem -Path $Path -Recurse -include $Excludes -File
    # Group the files by their name and extension
    ShowProgress   "Sorting for duplicate files in different directories .... Please wait"
    $FileGroups = $Files | Group-Object -Property Length | Where-Object count -gt 1 
    ShowProgress   "Checking for duplicate files in different directories by name"
    
    $FileGroups = $FileGroups.Group | Group-Object -Property Name | where-object count -gt 1
    ShowProgress   "Checking for duplicate files in different directories now by hash... this will take time... Please wait"
    $FileGroups = $FileGroups.Group | Group-Object -Property {(get-filehash $_.FullName -Algorithm md5).Hash } | where-object count -gt 1 
    $FileGroups.Group | group-object -property Name | Out-GridView -Title "Additional instances of these files will be removed - as the are exact duplicates" 
    # Loop through each group of files
    foreach ($FileGroup in $FileGroups) {

        # If the group has more than one file, it means there are duplicates
            # Sort the files by their directory depth, ascending
            $SortedFiles = $FileGroup.Group | Sort-Object -Property @{Expression = {$_.FullName.Split('\').Count}}
            # Loop through the rest of the files in the group, starting from the second one
            foreach ($File in $SortedFiles[1..($SortedFiles.Count - 1)]) {
                # Compare the file hashes of the first file and the current file
                # Add the current file's System.IO.FileInfo object to the list of duplicates
                $FileList += $File 
                
            
        }
    }
 
    
    # Return the list of duplicate files
    return $FileList
}

# Look within the directory and find files where the basename is the same as another instance and return that list based on the preference types
function DuplicateFileNames($Path, $ExtensionsOrder = @()) {
    # Initialize an empty list to store the file objects
    $FileList = @()
    $sp = 0
    # Get all the files in the directory and sub-directories recursively
    $Files = Get-ChildItem -Path $Path -Recurse -File
    
    # If the preferred extensions list is not empty, check for duplicate names with different extensions
    if ($ExtensionsOrder.Count -gt 0) {
        # Group the files by their base name (without extension)
        ShowProgress   "Sorting for unneeded formats .... Please wait"
        $NameGroups = $Files | Group-Object -Property BaseName,LastWriteTime.date
        # Loop through each group of files
        foreach ($NameGroup in $NameGroups) {
            if (($sp++ % 20) -eq 0) {
                ShowProgress   "Checking for unneeded formats"
                }
    
            # If the group has more than one file, it means there are duplicates
            if ($NameGroup.Count -gt 1) {
                # Sort the files by their extension, using the preferred extensions list as the order
                
                $SortedFiles = $NameGroup.Group | Sort-Object -Property @{Expression = {(&{if ($ExtensionsOrder.IndexOf($_.Extension) -ne -1) { $ExtensionsOrder.IndexOf($_.Extension) } else {100} })}; Descending = $false}
                # Loop through the rest of the files in the group, starting from the second one
                foreach ($File in $SortedFiles[1..($SortedFiles.Count - 1)]) {
                    # Add the current file's System.IO.FileInfo object to the list of duplicates
                    $FileList += $File 
                }
            }

        }
        $Files  | Sort-Object -Property @{Expression = {(&{if ($ExtensionsOrder.IndexOf($_.Extension) -ne -1) { $ExtensionsOrder.IndexOf($_.Extension) } else {100} })}; Descending = $false}   | Group-Object -Property BaseName | where-object count -gt 1 | Out-GridView -Title "Additional instances of these files will be removed - first instances is kept additional are removed" 

    }
    # Return the list of duplicate named files with different types
    return $FileList
}

<#
.SYNOPSIS
    Checks for files to remove and optionally deletes them.

.DESCRIPTION
    This function checks a collection of files and prompts the user to remove them.
    It can either move the files to the recycle bin or delete them permanently.

.PARAMETER RemoveFiles
    An array of FileInfo objects representing the files to be checked and potentially removed.

.PARAMETER HardDelete
    A switch parameter that, when set, will cause files to be deleted permanently instead of being recycled.

.EXAMPLE
    $filesToRemove = Get-ChildItem -Path "C:\Temp\*" -File
    CheckAndRemove -RemoveFiles $filesToRemove -DeleteWithoutRecycle $true -why "are duplicates"

    This example will check all files in C:\Temp and prompt the user to remove them.

.INPUTS
    System.IO.FileInfo[]

.OUTPUTS
    None

#>
function CheckAndRemove {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        [System.IO.FileInfo[]]$RemoveFiles,
        [boolean]$DeleteWithoutRecycle,
        [string]$why
    )

    $fcr = $RemoveFiles.length
    if ($fcr  -gt 0) {
        write-host "Found ${fcr} files that $why and should be removed" -ForegroundColor Yellow
        # $RemoveFiles|Select-Object Name, FullName, DirectoryName, Extension | Out-GridView -Title "Files that will be removed - $why (Close this Windows to continue)" 
        $cont = MyPause 'Remove those files? (No to keep them)'  -Choice $true -BoxMsg 'Click Yes to remove them' -ChoiceDefault $false

        if ($cont) {
            if (!$DeleteWithoutRecycle -and $fcr  -gt 100) {
                $cont = (MyPause 'This is going to take a while as it moves the files to recycle.  Would you like to Delete the file without being able to recover them?'  $true 'Click Yes to for a quick delete with NO Recyle!') 
                if ($cont) {
                    $DeleteWithoutRecycle = $true
                    Write-Host "Switching to Fast quick delete without recycle" -ForegroundColor Yellow
                    }
                }
            $howDeleted = if ($HardDelete) { 'Deleting ' } else { 'Recycling ' }
            $fcs = 0
            ForEach ($f in $RemoveFiles) {
                RecycleFile -file $f.FullName -purge $DeleteWithoutRecycle
                LogAction -File $f.Name -Action "--Remove-file"
                if ($fcs % 10 -eq 0) {
                    ShowProgress  ($howDeleted  + "extra files from cache") "$fcs of $fcr - $($f.Name)"
                    }
                $fcs++
                }
            
            }
            if ($fcs) {
                write-progress -Activity "Updating Lists of removed files .... Please wait"
                $script:mysewingfiles = $mysewingfiles |where-object {$_.FullName -notin ($RemoveFiles.FullName)}
                write-progress -Completed $true
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
        RecycleFile -file $Path -purge $HardDelete
        $IsFound = $true

        $Script:DirCount++
        ShowProgress "Removing Directory" $Path
        LogAction -File $Path -Action "--Remove-Empty-Directory"
    }

    return $IsFound
}

#-----------------------------------------------------------------------
# Only clear/reset the New files Directory once and only if there are new files with this run
# Let's becareful that we are not clearing the wrong directory
function ChecktoClearNewFilesDirectory {
    if ($Script:clearNewFiles) {
        if ((get-volume -filePath  $NewFilesDir).DriveType -eq "Fixed") {
            Get-ChildItem -Path $("\\?\" + $NewFilesDir) -Recurse | Remove-Item -Force -Recurse
            write-verbose "CLEARED Copy File Space"
        }
        $Script:clearNewFiles = $false
    }
}
<#
 function to convert characters which are not found in ASCII; such as á, é, í, ó, ú; into something acceptable such as a, e, i, o, u. 
 #>
function Remove-Diacritics
{
    Param([string]$Text)
    $chars = $Text.Normalize([System.Text.NormalizationForm]::FormD).GetEnumerator().Where{ 
        [System.Char]::GetUnicodeCategory($_) -ne [System.Globalization.UnicodeCategory]::NonSpacingMark
    }
    (-join $chars).Normalize([System.Text.NormalizationForm]::FormC)
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

function Get-LatestGitHubTag {
    param (
        [string]$RepositoryOwner,
        [string]$RepositoryName
    )
    
    # Construct the GitHub API URL for releases
    $apiUrl = "https://api.github.com/repos/$RepositoryOwner/$RepositoryName/releases"
    Write-Progress -Activity "Checking for script updates..."
    try {
        # Fetch the releases using Invoke-RestMethod
        $releases = Invoke-RestMethod -Uri $apiUrl
        # Filter the releases to get the latest one
        $latestRelease = $releases | Sort-Object -Property created_at -Descending | Select-Object -First 1
        # Extract and return the tag name
        $latestRelease = $latestRelease.tag_name
    }
    catch {
        Write-Verbose "Error fetching releases from GitHub: $_"
        $latestRelease = ""    
    }
    Write-Progress -Completed $true
    return $latestRelease
    
}



Function OpenForUpload {
    
    Write-Host "-----------------------------------------------------------------------------------------" -ForegroundColor Green
        
    if ($DragUpload) {
        Write-Host "Opening File Explorer & MySewnet Web page" -ForegroundColor Green
        Write-Host " ** on MySewNet web page choose 'Upload' and Select all files in Explorer and " -ForegroundColor Green
        Write-Host "    drag/drop the files a maximum of 5 at a time into the upload box" -ForegroundColor Green
        
    } else {
        
        if ((Get-WmiObject -class Win32_OperatingSystem).Caption -match "Windows 11") {
            $wtype = "W11"
            Write-Host "Opening File Explorer (using mysewnet add-in)" -ForegroundColor Green
            Write-Host " ***  Select all files *right-click* and choose 'Show more Options' -> choose 'MySewNet' -> 'Send'" -ForegroundColor Green
        } else {
            # Assume it is Windows 10 with add-in
            $wtype = "W10"
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
        FetchImageFile $file "https://raw.githubusercontent.com/D-Jeffrey/Embroidery-File-Organize/docs/images/main/HowToSend-$wtype.gif"
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


#=======================================================================
#  Cloud API MySewnet interface reversed by Darren Jeffrey
#=======================================================================

<#
.SYNOPSIS
    This function logs into the SewnetCloud.

.DESCRIPTION
    The LoginSewnetCloud function sends a POST request to the SewnetCloud API to authenticate a user. 
    It takes a username and password as parameters, and if the authentication is successful, it stores the session token in a global variable.


.EXAMPLE
    LoginSewnetCloud 

.OUTPUTS
    Boolean. Returns $true if the login is successful, $false otherwise.

.NOTES
    The function sets the global variable $Global:sewAuthorizeToken to the session token if the login is successful.
#>
function LoginSewnetCloud
{
    $authorization_endpoint = "https://auth.singer.com/authorize"
    $token_endpoint = "https://auth.singer.com/oauth/token"
    $idparams = @{
        Client_Id="rEJOLIIM2rpa155BupM4MxantCGqDc7o"
        scope="openid profile email offline_access"
        Redirect_Uri="https://mysewnet.com/my-account/login"
        customParameters = @{ 
            audience="https://api.mysewnet.com/"
            ui_locales="en"
        }
    }
    if ($Global:sewAuthorizeToken -and $Global:expires_on -gt (get-date) ) {
        ShowProgress "Using previous MySewNet Logon"
    } else {
        ShowProgress "Logginning onto MySewNet"
        try {
            $code = Invoke-OAuth2AuthorizationEndpoint -uri $authorization_endpoint @idparams
        } catch {
            write-progress -Completed $true
            return $false
        }
        ShowProgress "Authenicating onto MySewNet"
        $tokens = Invoke-OAuth2TokenEndpoint -uri $token_endpoint  @code  
        ShowProgress "Completed Logon to MySewNet"

        $Global:sewAuthorizeToken = $tokens.access_token

        if ($null -eq $tokens.access_token) {
            Write-error "Authentication Failed" 
            return $false
        } 
        $Global:expires_on = (get-date).AddSeconds($tokens.expires_in)

        Set-Content -Path $(Join-Path -Path $PSScriptRoot -childpath "Token.txt") $(($Global:expires_on.DateTime | ConvertTo-Json) + ($tokens | ConvertTo-Json))
    }

    

    return $true
}
 

<#
.SYNOPSIS
    This function generates authorization header values.

.DESCRIPTION
    The authHeaderValues function returns a hashtable of HTTP headers for authorization. 
    It uses the global variable $Global:sewAuthorizeToken to set the "Authorization" header.

.EXAMPLE
    $headers = authHeaderValues

.OUTPUTS
    Hashtable. Returns a hashtable of HTTP headers for authorization.

.NOTES
    The function uses the global variable $Global:sewAuthorizeToken to set the "Authorization" header.
#>

function authHeaderValues ()
{
    return @{ 
        "Accept"="application/json, text/plain, */*"
        "Accept-Encoding"="gzip, deflate, br, zstd"
        "Accept-Language"="en-US,en;q=0.9"
        "Authorization" = "Bearer " + $Global:sewAuthorizeToken
        "Origin"="https://www.mysewnet.com"
        "Referer"="https://www.mysewnet.com/"
        "Host"="api.mysewnet.com"
        "Pragma"= "no-cache"
        "Cache-Control"="no-cache"
        "User-Agent"="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36 Edg/124.0.0.0 EmbroideryCollection Manager Cleanup $ECCVERSION"
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
    do {
        ShowProgress "Reading file list from MySewNet $(" ".padright($tries,[char]2160))"
        try {
            $result = Invoke-RestMethod -Headers $authHeader -Uri $requestUri -Method GET -ContentType 'application/json'
        } catch {
            Start-Sleep -seconds 2
        }
        $tries++
        
    } while ($null -eq $result -and $tries -lt 3)
    
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
function CloudMetaAddPath {
param (
    [string]$path,
    [object]$metafolder = $webcollection
)    
    foreach ($fid in $metafolder.folders) {
        $pathHere = if ($path -eq "") { "\" + $fid.name } else { Join-Path -Path $path -ChildPath $fid.name }
        Add-Member -InputObject $fid -MemberType NoteProperty -Name 'path' -Value $pathHere
        CloudMetaAddPath -path $pathHere -metafolder $fid
    }
}

# BUG BUG BUG - This works if there are unique files names only

Function GetFileIDfromCloud 
{
param (
    [string]$fileNameExt,
    [object]$metafolder = $webcollection
)
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

function findMetaDirectory {
    param (
        [string]$folderid,
        [object]$metafolder = $webcollection
    )
    if ($folderid) {
        if ($metafolder) {
            if ($metafolder.folders | where-object { $_.id -like $folderid}) {
                return $metafolder.folders | where-object { $_.id -like $folderid}
            } else {
                foreach($fid in $metafolder.folders) {
                    $retdir = findMetaDirectory -folderid $folderid -metafolder $fid 
                    if ($retdir) {
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

function GetCloudDirectoryNameFromId {
    param (
        [string]$folderid
    )
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
function FindCloudidfromPath {
    param (
        [string]$foldername,
        [object]$metafolder = $webcollection
    )
    # return $metafolder | Select-Object -ExpandProperty Folders | Where-Object {$_.path -like $foldername} 
    
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
    if ($null -eq $id) {
        return
    }
    $authHeader = authHeaderValues
    $myError = $null    
    $requestUri = 'https://api.mysewnet.com/api/v2/cloud/folders/'+ $id;
    $webcollection.folderId
    $pool = $webcollection
    # BUGS
    <#
        do {
            $pool | Select-Object -ExpandProperty Folders |  Where-Object {$id -in ($_.folder.Id)} | ForEach-Object { 
                $_.folderId = ($id -notin ($_.FolderId))
                $pool = $null
                break
                }
            $pool = $pool | Select-Object -ExpandProperty Folders 
            } while ($pool.folders.count -gt 0)
#>
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
		write-warning ("Error Deleting folder id: '$id'  [ " + $eDetails.message + " ] -- The cloud API is buggy try again")
        $script:CloudStatusGood = $false

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
        $script:CloudStatusGood = $false

	}
    else {
        # TODO Manage shadow copy structure for Cloud folders

        <#  This does not work
        $filefind = $webcollection | Select-Object -ExpandProperty Folders | select-Object -ExpandProperty Files| Where-Object {$_.id -eq $id} 
        $parentref = findMetaDirectory -folderid $filefind.folderId
        if ($parentref) {
            $parentref.files = $parentref.files | where-Object { $_.id -ne $id}
        }
        #>
    }
    return $result
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
            $script:CloudStatusGood = $false

        }
        else {
            # TODO Update Cache
        }
    return $result
}


<#
Create a Folder $name within Folder ID $inFolderID
#>
Function CreateCloudFolder($name, $inFolderID)
{
    if ($name.substring(0,1) -eq '\') {
        $name = $name.substring(1)
    }
    $dianame = Remove-Diacritics $name
    $CheckName = findMetaDirectory -folderid $inFolderID
    if ($null -ne $CheckName) {
        if ($CheckName.Folders.name -like $name) {
            $fld = $CheckName.Folders | where-object { $_.name -like $dianame }
            # write-host "Folder " $fld.Name " ($name) exists as " $fld.id " for in " $fld.parentfolderid " ($inFoldID) "
            return $fld.id
        }
    }
    $authHeader = authHeaderValues
    $requestUri = 'https://api.mysewnet.com/api/v2/cloud/folders';
    $bodyLines = @{ 
        "folderName" = $dianame
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
        $script:CloudStatusGood = $false
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
    if ($script:CloudStatusGood) {
        return (PushCloudFile -name $filename -inFolderID $folderid -filepath $filepath)
    } else {
        return ""
    }
    
}

<#
.SYNOPSIS
    Uploads a file to a Mysewnet cloud given the Folder ID.

.DESCRIPTION
    The PushCloudFile function uploads a file to a specified directory in the cloud.
    It checks if the file already exists in the target directory and, if not, proceeds to upload it.

.PARAMETER name
    The name of the file to be uploaded.

.PARAMETER inFolderID
    The ID of the folder in the cloud where the file will be uploaded.

.PARAMETER filepath
    The local path to the file that needs to be uploaded.

.EXAMPLE
    PushCloudFile -name "example.txt" -inFolderID "12345" -filepath "C:\path\to\your\file.txt"

    This example uploads the file "example.txt" from the specified local path to the cloud directory with the ID "12345".

.INPUTS
    None

.OUTPUTS
    String - Giud/ID for the file upload
    If the file already exists, the function returns the ID of the existing file.

.NOTES
    Handles file upload through a REST API call using multipart/form-data content type.

.LINK
    https://api.mysewnet.com/documentation

#>

Function PushCloudFile($name, $inFolderID, $filepath)
{
    if (test-Path -Path $filepath) {
        $diaName = Remove-Diacritics $name
        $CheckFolder = findMetaDirectory -folderid $inFolderID
        if ($null -ne $CheckFolder) {
            $fld = $CheckFolder.Files | where-object { $_.name -like $dianame }
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
	        $dianame,
            "--$boundary",
            "Content-Disposition: form-data; name=`"FolderId`"$LF",
             $inFolderID,
	        "--$boundary--"
            ) -join $LF

	try {        
	        $result = Invoke-RestMethod -Uri $requestUri -Method POST -Headers $authHeader -ContentType "multipart/form-data; boundary=`"$boundary`"" -Body $bodyLines
	} catch {
		# Note that value__ is not a typo.
		# write-host "StatusCode:"  $_.Exception.Response.StatusCode.value__
		# write-warning ("StatusDescription:" + $_.Exception.Response.StatusDescription)
		$result = ""
		$myError = $_
	} 

	if ("" -eq $result) {
		$eDetails = ($myError.errorDetails.Message|convertfrom-json )
		write-warning ("Error Pushing the file: '$name'  [ "  + $myError.Exception.Response.StatusDescription + $eDetails.message + " : " + " ]") 
        $script:CloudStatusGood = $false

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
        Get-TraceSource
    }

}
function FoldupDirPath {
    param (
        [string]$directoryPath,
        $minimumlen = 2
    )

    do {
        $folding = $false
        foreach ($r in $foldupDirs) {
            if ($directoryPath.ToLower().EndsWith("\$r")) {
                if ($(split-path $directoryPath -IsAbsolute) -and ($directoryPath.Length -lt $EmbroidDir.Length)) {
                    # Emergency panic exit
                    $directoryPath = $EmbroidDir
                    write-warning "Folding Directories has a problem, please change the 'foldupDir' List"
                    break
                }
                # Strip off the directory name and preserve the case of the directory and files
                # Carefully tested trimming
                # Does fold it too much
                if (($directoryPath.Length - $r.Length - 1)  -gt $minimumlen) {
                    $directoryPath = $directoryPath -ireplace "\\$r$", ""
                    $folding = $true
                    }
            }
        }
    } while ($folding)

    return $directoryPath
}

#-----------------------------------------------------------------------
# Move From Directory to either Embrodery or Instrctions directory
function MoveFromDir ( 
    [string] $fromPath, 
    [boolean]$isEmbrodery = $false,        # is this to the Embrodery directory (true) Or to the Instruction Directory (false)
    [string]$whichfiles = $null,            # File Names
    [string[]]$files = $null,                # File extension types as an array
    [string]$isFromNestedRelative = $null      # Caution zip inside of zip inside of zip
    ) 
{
    $loopy = 0
    $newFileCount = 0
    if ($isEmbrodery) { 
        $dtype = "Embroidery"
        if ($whichfiles) {
            $objs = Get-ChildItem -Path $fromPath -include $whichfiles -File -Recurse  
        } else {
            $objs = Get-ChildItem -Path $fromPath -include $PrefSewTypeStar -File -Recurse  
        }
        $sublen = $fromPath.Length
        if ($files) {
            $objs = $objs | Where-Object {$_.Fullname.Substring($sublen) -iin $files -or $_.Fullname.Substring($sublen+1) -iin $files}
        }
        $targetdir = $EmbroidDir
    } else { 
        # Move anything that is not a Embrodery type file (alltypes)
        $dtype = "Instructions"
        $Excludes = ($allTypesStar + $TandCs )
        # if it is from a nested zip, then keep the zip because we don't yet expand that TODO
        if (!($isFromNestedRelative)) {
            $Excludes += @("*.zip")
        }
        $objs = Get-ChildItem $fromPath -file -Recurse -Exclude $Excludes 
        if ($whichfiles) {
            $objs = $objs | where-Object { $whichfiles -ilike $_.Name } 
        }
        $targetdir = $InstructDir
        }
    if ($objs.count) {
        ShowProgress -Area "Copying $dtype" -stat "Added ${Script:savecnt} files"
        }

    $objs | ForEach-Object {
        # If all files or the name matching a file we should be moving..
#        if (($null -eq $files) -or ($_.Name -in $files)) {                
            # Get the relative path 
            $newdir = (Split-Path(($_.FullName).substring(($fromPath.Length), 
                                ($_.FullName).Length - ($fromPath.Length) )))
            $newfile = $_.Name
            if ($isFromNestedRelative) {
                $newdir = join-path $isFromNestedRelative -ChildPath $newdir
            }
            # take off the directory name if it is one of the rollup names
            $newdir = FoldupDirPath -directoryPath $newdir
            
            $newpath = join-path -path $targetdir -childpath $newdir
            if (Test-Path -Path $newpath -PathType Leaf) {
                # what do if there is already a file with the same name as the folder? Rename the confliciting folder
                $parent = Split-Path -Path $newpath -Parent
                $thisfolder = Split-Path -Path $newpath -Leaf
                $newpath = join-path -path $parent -childpath $($thisfolder.Replace('.','-'))
            }
            if (!(Test-Path -Path $newpath -PathType Container)) {
                New-Item -Path $newpath -ItemType Directory | Out-Null
                }
            $npath = Join-Path -Path $newpath -ChildPath $newfile 
            if (test-path $npath) {  # See if the file already exists
                
                $newHash = get-filehash -Algorithm md5 $_
                $orgHash = get-filehash -Algorithm md5 $npath
                if ($orgHash -eq $newHash) {
                    $attrcheck = get-item $_.FullName
                    if ($attrcheck.attributes.hasflag([IO.FileAttributes]'Readonly')) {
                        $attrcheck.Attributes -= 'Readonly'
                    }
                    if ($RemovePrefix) {
                        Remove-Item -Path "\\?\$($_.FullName)" -ErrorAction SilentlyContinue
                    } else {
                        Remove-Item -Path $_.FullName -ErrorAction SilentlyContinue
                    }
                    Write-Verbose "Removed Duplicate ${dtype} file :'$_'" 
                    
                    }
                
                }
            # Test to see if we purged the file ub the previous step, otherwise we will overwrite the file
            if (test-path $_) {
                # BUG 
                # this can happen with the same file type is nested within other directories which then get folded to the same directory and are duplicate
                # We could rename the files, but then the code above need to find and match the same rename pattern
                if (test-path $npath) { 
                    Write-verbose " WARNING File already exists $npath for $($_.FullName) - maybe renaming - overwriting?"
                    }
                else {
                    ChecktoClearNewFilesDirectory
                    if ($NoDirectory) {
                        Copy-Item -Path $_ -Destination $(Join-Path -Path $NewFilesDir  -ChildPath $newfile)
                    } else {
                        $newpath = join-path -path $NewFilesDir -childpath $newdir
                        if (!(test-path ($newpath))) {
                            New-Item -Path ($newpath) -ItemType Directory  | Out-Null
                            }
                        Copy-Item -Path $_ -Destination $(Join-Path -Path $newpath -ChildPath $newfile)
                        }
                    LogAction $newfile -Action "++Added-MoveFrom"
                    # BUG THIS IS THE LINE THAT ERRORS out wiht Directory Not Found
                    $movedirto = split-path -path $npath -parent
                    if (!(test-path -Path $movedirto -PathType Container)) {
                        New-Item -Path $movedirto -ItemType Directory  | Out-Null
                    }
                    Move-Item $_ -Destination $npath  -force # -ErrorAction SilentlyContinue
                    $newFileCount += 1
                    Write-Information "+++ Saving ${dtype}:'$_' to ${newdir} & ${newfiledir}"
                    <#if ($isEmbrodery) { 
                        $Script:addsizecnt = $Script:addsizecnt + (Get-Item -Path $npath).Length 
                        $Script:savecnt = $Script:savecnt + 1
                        }
                        #>
                    }
                }
#            } else {
#                Write-Verbose "Skipping ${_.Name}" 
#            }
        
        if ($loopy++ % 20 -eq 0) {
            ShowProgress -Area "Copying $($_.Name)" -Stat "Added ${Script:savecnt} files"
        }
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
        FullName        C:\Dir\File.txt : Full path and name
        Filedatetime                    : lastwritetime of file on disk or zip
        Priority                        : sorted type based on preference
        LastWriteTime                   : Last WriteTime from the file or zip
        CloudRef                        : See object below
        RelPath                         : Directory path relative to the root of EmbiodryDir
        Push                            : Path if new file else null

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
    $thelist = (Get-ChildItem -Path $EmbroidDir  -Recurse -file -include $PrefSewTypeStar)| ForEach-Object { 
        if ($keepAllTypes) {
            $n = $_.Name} 
        else {
            $n = $_.BaseName
        }
        
            [PSCustomObject]@{                          # C:\Dir\File.txt
            NameIndexed = $n                               
            N = $_.Name                                         # File.txt
            # Ext = $_.Extension                                  # txt
            Base = $_.BaseName                                  # File
            DirectoryName = $_.DirectoryName                    # C:\Dir\
            Hash = [string]$null                                # hash value of the file calculated when we need it
            FullName = $_.FullName                              # C:\Dir\File.txt
            FileInfo = $_
            LastWriteTime = $_.LastWriteTime
            Priority = $preferredSewType.Indexof($_.Extension.substring(1,$_.Extension.Length-1).tolower())
            RelPath = $_.DirectoryName.Substring($EmbroidDir.Length)
            CloudRef = $null                                    #
            Push = $null
            TmpPath = $null
            } 
        }
    if ($null -eq $thelist) {
        $FileInfo =  New-Object System.IO.FileInfo("C:\placeholder.directoryname\zzzmysewingfiles.placeholder")
        $thelist =    
            [PSCustomObject]@{ 
                NameIndexed = $fileInfo.Name                               
                N = $fileInfo.Name                                         # File.txt
                # Ext = $fileInfo.Extension                                  # txt
                Base = $fileInfo.BaseName                                  # File
                DirectoryName = $fileInfo.DirectoryName                    # C:\Dir\
                Hash = "A100000A"                                          # hash value of the file calculated when we need it
                FullName = $fileInfo.FullName                              # C:\Dir\File.txt
                FileInfo = $fileInfo
                LastWriteTime = $fileInfo.LastWriteTime
                Priority = 100
                RelPath = '?????'
                CloudRef = $null  
                Push = ""
                TmpPath = $null
                } 
                
        }
        
    if ($thelist.getType().Name -ne 'Object[]') {
        # Hack the object to cause it to be an array
        $thelist = @( $thelist, $thelist )
        }
    return $theList
}
function BuildHashofMySewingFiles {
    $setList = @{}
    
    for ($index = 0; $index -lt $MySewingfiles.count; $index++) {
        if ($setList[$MySewingfiles[$index].NameIndexed.tolower()]) {
            $setList[$MySewingfiles[$index].NameIndexed.tolower()] += ($index+1)
            }
        else { 
            $setList.Add($MySewingfiles[$index].NameIndexed.tolower(), @(($index+1)))
            }
    }
#     $setList | out-gridview
    return $setList
}
#======================================================================================

function ShowPreferences ($showall = $false)
{
    if ($showall) {
        Set-Variable -Name $param -Value $SavedParam.$param
        foreach ($paramselect in ($paramstring.Keys  + $parambool.Keys)) {
            $description = $paramstring[$paramselect] + $parambool[$paramselect]
            if ($description) {
                $val = Get-Variable -Name $paramselect -ValueOnly
                Write-host $($description).padright($padder+20) ": " $val
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
    if ($CloudAPI) {
        if ($CloudAuthAvailable)  {
            Write-host    $($paramswitch['CloudAPI']).padright($padder)   -ForegroundColor yellow
        } else {
            Write-host "Using API to update mySewNet Cloud requires PSAuthClient to be installed to use CloudAPI feature" -ForegroundColor Red
            }
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

function ShouldMakeASet {
param ([System.IO.Compression.ZipArchive]$zipfile)

    $settotal = 0
    foreach ($exts in $preferredSewType) {
        $extstar = "*.$exts"
        $settotal += ($zipfile.Entries | where-object {$_.Name -like $extstar}).count
        if ($settotal -ge $SetSize) {
            return $true
        }
    }
    return $false
}

# $mysewingfiles | ft
function AddToSewList {
    Param (
    $NameIndex,
    $Name,
    $Directory,
    $LastWriteTime,
    $RelativePath = $null,
    [System.IO.FileInfo]$TmpPath = $null,
    [System.IO.FileInfo]$keepPath = $null
    )
    if ($nameIndex -eq "" -or $nameIndex -eq $null) {
        
        write-Error "** BLANK NAME - '$NameIndex', '$Name', '$directory', '$lastWriteTime' "
        start-sleep -Milliseconds 100
    }
    $isnewfile = $true
    if ($quickmysewfiles[$NameIndex]) {
        # BUG duplicate filename but different checksum??? 
        # TODO use  LastWriteTime to overcome
        # Check for duplicate FileHash
        # Find the instances then get the full name then get hash to compare
        # If there are multiple versions of the file (with different extensions then this Verbose will show several values)
        foreach ($q in $quickmysewfiles[$NameIndex]) {
            if ($LastWriteTime.date -eq $mysewingfiles[$q-1].LastWriteTime.date) {
                $isnewfile = $false
                break
                }
            }
        }
    
    $hash = $null


   if ($KeepAllTypes) {
        # $tmpfilepath = join-path -Path $tmpdir -ChildPath $RelativePath | join-path -ChildPath $Name
        if ($TmpPath.Exists) {
            $hash = (get-filehash -Algorithm md5 $TmpPath).Hash
            # TODO need to retest the file compare beyond name and date to Hash
            if (!($isnewfile)) {
                $isnewfile = $true
                foreach ($q in $quickmysewfiles[$NameIndex]) {
                    if (!($mysewingfiles[$q-1].Hash)) {
                        $mysewingfiles[$q-1].Hash = (get-filehash -Algorithm md5 $mysewingfiles[$q-1].FullName).Hash
                        }
                    if ($mysewingfiles[$q-1].Hash -eq $hash) {
                        $isnewfile = $false
                        break
                        }
                    }
                }
            }
        }

    if (!$isnewfile) {
        return ""
    }
    $Directory = FoldupDirPath -directoryPath $Directory
    $fullName = join-path -Path $Directory -ChildPath $Name
    $fileinfo = New-Object System.IO.FileInfo($fullname) 
    <# if (!($RelativePath)) {
        
        $RelativePath = (Split-path -Path $fullName -parent)
        if ($EmbroidDir.Length -lt $RelativePath.Length) {
            $RelativePath = $RelativePath.substring($EmbroidDir.Length+1)
        } else {
            $RelativePath = ""
            }

        }
        #>
    $script:mysewingfiles +=  
        [PSCustomObject]@{ 
            NameIndexed = $NameIndex
            N = $Name
            # Ext = $fileinfo.Extension
            Base = $fileinfo.BaseName
            DirectoryName = $Directory
            Hash = $hash
            FullName = $fullName
            FileInfo =  $FileInfo
            LastWriteTime = $LastWriteTime
            Priority = $preferredSewType.Indexof($fileinfo.Extension.tolower())
            RelPath = $relativepath
            CloudRef = $null
            Push = '\'+ $RelativePath
            TmpPath = $TmpPath
            KeepPath = $keepPath
        }
    $currentSewingFile = $mysewingfiles.count
    if ($script:quickmysewfiles[$NameIndex.tolower()]) {
        $script:quickmysewfiles[$NameIndex.tolower()] += $currentSewingFile
        }
    else { 
        $script:quickmysewfiles.Add($NameIndex.tolower(), @($currentSewingFile))
        }
    $Script:addsizecnt = $Script:addsizecnt + $_.Length 
    $Script:savecnt = $Script:savecnt + 1 
    if ($relativepath) {
        return join-path -Path $relativepath -childpath $Name
        } else {
        return $Name
        }
    
}
function ExpandAZip {
    param (
        $zippath,
        $RelativePath
    )
    $resultTmpDir = (Join-Path $tmpdir -childpath $RelativePath).trim("\")
    # Check for long path names inside the zip file
    write-progress "Expanding Zip Archive.. Please wait " -Status (split-path -path $zippath -leaf)
    $bigzip = (get-item $zippath).Length -gt $use7zipsize
    if (($bigzip -or $havewarning) -and (Test-Path "C:\Program Files\7-Zip\7z.exe")) {
        Set-Alias sevenz "C:\Program Files\7-Zip\7z.exe"
        # write-host "`n`n`n`n`n"
        sevenz x $zippath -o"$resultTmpDir" -y
    } else {
        Expand-Archive -Path $zippath -DestinationPath $resultTmpDir -Force
    }
    write-progress -Completed $true
    return $resultTmpDir
}

function ProcessZipContents {
    param (
        $zips,       # FullName
        $Base,
        $isNested = $false,
        $NestedName = ""       # Is it a zip in a zip
    )

    $thisZipBase = $Base -replace ' \(([0-9]+)\)'
    try {
        $zipfilelist = [io.compression.zipfile]::OpenRead($zips)
    } catch {
        write-warning "- - - Problem with zip file, skipping : $zips"
        return 
    }
    
    
    $makeASet = ShouldMakeASet($zipfilelist)
    if ($makeASet) {     # This zip file is big enought to keep the files together
        if ($NestedName) {
            $madeDir = $NestedName + "\" + $thisZipBase
        } else {
            $madeDir = "\" + $thisZipBase
        }
    } else {
        $madeDir = $NestedName
    }
    $resultTmpDir = $null
    $isNewInstruct = $false
    $isnew = $false
    $numnew = 0
    foreach ($thistype in $preferredSewType) {
        $filesInThisList = @()
        $ts = "*."+ $thistype
        if ($zipfilelist.Entries.Name -ilike $ts) {
            $isnew = $false
            # like our extension, but does not start with . 
            $SpecificExtensionFiles = $zipfilelist.Entries | where-object {$_.Name -like $ts -AND $_.Name.substring(0,1) -ne "."} 
            ShowProgress  "Checking Zips - Looking at $($_.Name) - looking at '$($ts.substring(2).ToUpper())' type"  -stat "Added $Script:savecnt files"
            foreach ($fileInZip in $SpecificExtensionFiles) {
                $isnewfile = ""
                $filenameInZip = $fileInZip.Name
                # grab the base file name only (works with or without extension)
                if ($PSVersionTable.PSVersion.Major -ge 6) {
                    $fs = split-path -Path $filenameInZip -LeafBase
                } else {
                    $fs = (Split-Path -Path $filenameInZip  -Leaf) -replace '\.[^.]*$'
                    }
                if ($keepAllTypes) {
                    $n = $fileInZip.Name 
                } else {
                    $filenameInZip = $fs
                    $n = $fs
                    }
                $relativepath = Split-path -Path $fileInZip.FullName -parent
                if ($madedir) {
                $relativepath = join-path -Path $madeDir -ChildPath $relativepath
                }
                    
                $dirn = (join-path -Path $EmbroidDir -ChildPath $relativepath).trim('\')
                if ($isNewInstruct) {
                    if ($madedir) {
                        $buildpath = join-path -path $tmpdir -ChildPath $madeDir 
                    } else {
                        $buildpath = $tmpdir
                        }
                    $tempPath = get-item -path $(join-path -path $buildpath -childpath $fileInZip.FullName)
                } else {
                    $tempPath = $null
                }
                $isnewfile = AddToSewList -Name $fileInZip.Name -Directory $dirn -NameIndex $n -LastWriteTime $fileInZip.LastWriteTime -RelativePath $relativepath -TmpPath $tempPath
                if ($isnewfile) {
                    Write-verbose "New file '${filenameInZip}'"
                    $isnew  = $true
                    $filesInThisList += $isnewfile
                    if (-not $isNewInstruct) { 
                        $isNewInstruct = $true
                        ExpandAZip -zippath $zips -RelativePath $madeDir |out-null
                        # Fix up Hash triggered by the first file
                        foreach ($q in $quickmysewfiles[$n]) {
                            if (!($mysewingfiles[$q-1].Hash) -and !(test-path -Path $mysewingfiles[$q-1].FullName)) {
                                $tmpfile = $(join-path -Path $tmpdir -ChildPath $mysewingfiles[$q-1].RelPath | join-path -ChildPath $mysewingfiles[$q-1].N)
                                if (test-path $tmpfile) {
                                    $Script:mysewingfiles[$q-1].Hash = (get-FileHash -Algorithm md5 -Path $tmpfile).Hash
                                    if ($madedir) {
                                        $buildpath = join-path -path $tmpdir -ChildPath $madeDir 
                                    } else {
                                        $buildpath = $tmpdir
                                        }
                                    $tempPath = get-item -path $(join-path -path $buildpath -ChildPath $fileInZip.FullName)
                                    $Script:mysewingfiles[$q-1].TmpPath = $tempPath
                                }
                            }
                        }
                    }
                } else {
                    if ($VerbosePreference -eq  "Continue") {
                        $fileInstance = $mysewingfiles | where-object {$_.NameIndexed -eq $filenameInZip}
                        $fiName = $fileInstance.FullName
                        Write-verbose "Duplicate zfile '${filenameInZip}' to ${fiName}"
                        }
                    }
                }
            
                # we found a new file in the Zip.  If we have not expanded this Zip, then do it now
                if ($isnew) { 
                    $numnew += $(MoveFromDir -fromPath $tmpdir -isEmbrodery $true -files $filesInThisList -whichfiles $ts)
                }               
            }
        }
        if ($isNewInstruct) {
            $zipfilelist.Entries | where-object Name -imatch "\.zip$" | ForEach-Object { 
                Write-Host "- - Found: nested zip file $($_.FullName) checking" 
                if ($madeDir) {
                    $tmpzip = join-path -path $tmpdir -ChildPath $madeDir    
                } else {
                    $tmpzip = $tmpdir
                }
                $tmpzip = join-path -path $tmpzip -ChildPath $_.FullName
                $nestBase = $($_.FullName -replace '\.[^.]*$' ) -replace '\/','\' 
                
                ProcessZipContents -zips $tmpzip -Base $nestBase -NestedName $madeDir -isNested $true
            }
        }
        
        $zipfilelist.Dispose()      # Close Zipfile
        
        $zf = $zips.tolower().replace($downloaddir.tolower(),'...').replace($tmpdir.tolower(),'...')
        if ($numnew -gt 0) {
            if ($isNested) {
                $indent = "- + "
            } else {
                $indent = "* "
            }
            Write-host $("$indent New  : '$zf'").padright(65) " $numnew new patterns" -foreground green
            ShowProgress  "Checking Zips"  "Added $Script:savecnt files"
        } else {
            Write-host $("- Found: '$zf'").padright(65) " nothing new" 
            }
        # we extracted the Zip already and now let's check for instructions but only do this at the very end of nested files
        if ($isNewInstruct -and (!($isNested))) { 
            $numnew += $(MoveFromDir -fromPath $tmpdir -isEmbrodery $false)
            Get-ChildItem -Path $tmpdir -Recurse | Remove-Item -force -Recurse
            }
   
        
    }   
#======================================================================================
#
# Building out all the directory structures and File lists
#
#======================================================================================
$PrefSewTypeStar = $preferredSewType | ForEach-Object { "*.$_" }
if ($PrefSewTypeStar.count -eq 0 -or $null -eq $PrefSewTypeStar) {
    write-error "Miss configuration of 'preferredSewType', can not continue"
    return
}
$SewTypeMatch = $preferredSewType -join '|'
$foldupDirs = $foldupDir + $preferredSewType | ForEach-Object { $_.ToLower() }
if ($foldupDirs.count -eq 0 -or $null -eq $foldupDir) {
    write-error "Miss configuration of 'foldupDir', can not continue"
    return
}
$allTypesStar = $alltypes | ForEach-Object {"*.$_"}



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

# TODO test path for exists

$LogFile = join-path $PSScriptRoot -childpath "EmbroideryCollection.Log"
if (!(test-path $LogFile)) {
    "$PSCommandPath Powershell action log file\n" | Set-Content -Path $LogFile 
}


if ($null -eq $LastCheckedGithub -or ($(get-date) -gt $(get-date $LastCheckedGithub).adddays(7)))  {
    $latestTag = Get-LatestGitHubTag -RepositoryOwner "D-Jeffrey" -RepositoryName "Embroidery-File-Organize"
    $script:LastCheckedGithub = get-date -format "g"
    if ($latestTag) {
        Write-Verbose "Latest tag in D-Jeffrey/Embroidery-File-Organize $latestTag"
        if ($latestTag -ne $ECCVERSION) {
            Write-host "  *** Newer version ($latestTag) of this script is available" -ForegroundColor Green
            $upgrademe = MyPause -Message "Do you want to upgrade" -Choice $true -Timeout 300 -ChoiceDefault $false
            if ($upgrademe) {
                $upgradescript = Join-Path -Path $PSScriptRoot -ChildPath "install.ps1"
                if (test-path $upgradescript) {
                    powershell -ExecutionPolicy bypass -file $upgradescript
                    return
                } else {
                    Write-Warning "Automatic upgrade script '$upgradescript' can not be found to run, continuing without upgrade"
                }
            }
        }
    }
    else {
        Write-Verbose "Failed to retrieve the latest tag for $owner/$repo from github."
    }
    }



if ($setup) {
    write-host "   ".padright(70) -BackgroundColor Yellow -ForegroundColor Black
    $Desktop = [Environment]::GetFolderPath("Desktop")
    $DesktopLink = $Desktop + "\Embroidery Organizer.lnk"
    $WshShell = New-Object -comObject WScript.Shell
    $Shortcut = $WshShell.CreateShortcut($DesktopLink)
    if (test-path ($Desktop + "\Embroidery Organizer.lnk")) {
        if ($Shortcut.TargetPath.contains("powershell.exe")) {
            if (Test-ExistsOnPath "pwsh.exe") {
                Write-Host "    Upgraded to using PWSH" -BackgroundColor Yellow -ForegroundColor Black
                $Shortcut.TargetPath = "pwsh.exe"
                $Shortcut.Save()
                LogAction -File $Desktop -Action "Updated-Desktop-Shortcut PWSH"
            }
        } 
    }
    else {
        write-host "  Creating shortcut on the Desktop".padright(70) -BackgroundColor Yellow -ForegroundColor Black
        
        if (Test-ExistsOnPath "pwsh.exe") {
            $Shortcut.TargetPath = "pwsh.exe"
        } else {
            $Shortcut.TargetPath = "$pshome\Powershell.exe"
            }
        $icon = join-path -Path $PSScriptRoot -childpath "embroiderymanager.ico"
        if (!(test-path -path $icon )) {
            try {
                (New-Object System.Net.WebClient).DownloadFile("https://raw.githubusercontent.com/D-Jeffrey/Embroidery-File-Organize/main/embroiderymanager.ico", $icon)
            } catch {
                Write-Host "`t[!] Failed to download '$icon'"
                }
            }
        $Shortcut.IconLocation = $icon    
        $Shortcut.Arguments = "-NoLogo -ExecutionPolicy Bypass -File ""$PSCommandPath"""
        $Shortcut.Description = "Run EmbroideryCollection-Cleanup.ps1 to extract the patterns from the download directory"
        $Shortcut.Save()
        LogAction -File $Desktop -Action "Created-Desktop-Shortcut"
        }
    ShowPreferences
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
            $CloudAPI = $false
        } else {
            $USBDrive = ""
        }
        if ($USBDrive -eq "") {
            if (myPause -Message "Are you using MySewnet Cloud" -Choice $true -ChoiceDefault $CloudAPI) {
                $CloudAPI = $true
                $USBDrive = ""
                if ((get-module -name PSAuthClient).count -lt 1) {
                    write-Host "This requires the installation of PSAuthClient"
                    if (myPause -Message "Do you want to install that module now" -Choice $true -ChoiceDefault $false) {
                        if (([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
                            write-host "This may take a few minutes to complete - using Adminstrator access" -ForegroundColor Green
                            install-module -name PSAuthClient -Scope:AllUsers
                            write-host "Completed" -ForegroundColor Green
                        } else {
                            write-host "This may take a few minutes to complete" -ForegroundColor Green
                            install-module -name PSAuthClient -scope:CurrentUser
                            write-host "Completed" -ForegroundColor Green
                        }
                    }
                }
            } else {
                $CloudAPI = $false
                $USBDrive = ""
            }
        }
        $dd = Read-Host "How many days back do you want to always look when checking the Download folder? (currently $DownloadDaysOld)"
        if ($dd -gt 0) {
            $DownloadDaysOld = $dd
        }         
        $KeepAllTypes = myPause $parambool['KeepAllTypes'] -Choice $true -ChoiceDefault $KeepAllTypes
        $KeepEmptyDirectory = myPause $parambool['KeepEmptyDirectory'] -Choice $true -ChoiceDefault $KeepEmptyDirectory
        if ($CloudAPI) {
            $DragUpload = myPause $parambool['DragUpload'] -Choice $true -ChoiceDefault $DragUpload
            $ShowExample =myPause $parambool['ShowExample'] -Choice $true -ChoiceDefault $ShowExample
            }
        
#        $NoDirectory = myPause $parambool['NoDirectory'] -Choice $true -ChoiceDefault $NoDirectory
#        $OneDirectory = myPause $parambool['OneDirectory'] -Choice $true -ChoiceDefault $OneDirectory
        $val = $alltypes  -join ', '
        Write-host "All the different Embroidery file types: `n$val" 
        
        
        Write-host "What are the preferred types of files for your machine in order of preference"
        write-host "Current list is: " $($preferredSewType -join ", " )  -ForegroundColor Yellow
        
        do {
            write-host "Files types (seperated by comma)?"  -NoNewline  -ForegroundColor Yellow
            $ptype = Read-Host 
            if ($ptype) {
                $ptype = ($ptype.split(',')).trim()
                $ptype = $ptype |Where-Object {$_.length -gt 0}
                $problemext = $ptype |Where-Object {($_.contains(' ')) -or ($_.Length -gt 4)}
                if ($problemext) {
                    write-host "Problem with the extension of: $problemext" -ForegroundColor Red
                } else {
                    # BUG must be an Array
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
        write-host "  Creating Directory for Embroidery files cache ".padright(70) -BackgroundColor Yellow -ForegroundColor Black
        
        if (!(Test-Path($EmbroidDir))) {
            New-Item -ItemType Directory -Path $EmbroidDir | out-null
            write-host "Creating Directory '$EmbroidDir' for Embroidery files" -BackgroundColor Green -ForegroundColor Black
            LogAction -File $EmbroidDir -Action "Created-Directory"
        }
        if (!(Test-Path($InstructDir))) {
            New-Item -ItemType Directory -Path $instructDir | out-null
            write-host "Creating Directory '$instructDir' for Instructions files" -BackgroundColor Green -ForegroundColor Black
            LogAction -File $instructDir -Action "Created-Directory"
        }
    }
    if ($CloudAPI -and ((get-module -name PSAuthClient).count -lt 1)) {
        write-host "You will need to install the Powershell module in order to access MySewnet use the following command:" -ForegroundColor Yellow
        write-host "      install-module -name PSAuthClient" -ForegroundColor Blue
        write-host "with completing that you will not be able to use the CloudAPI feature" -ForegroundColor Yellow

    }
    write-host "All Setup " -BackgroundColor Yellow -ForegroundColor Black
    $FirstRun = mypause -Message "What you like the run the script and collect all your Embroidery that you have every downloaded? " -Choice $true
    if (-not $firstRun) {
        Return
    }
    # If we got to this point then $FirstRun is $true
}
Write-Host " ".padright(5) "Let's begin managing the Embroidery files".padright(80) -ForegroundColor white -BackgroundColor blue
if ($FirstRun) {
    Write-Host " ".padright(15) $("** Checking ALL Zip files **".padright(70)) -ForegroundColor white -BackgroundColor blue
}

# Clean out the old tmp working space
write-progress -Activity "Cleaning up temporary work space"
Get-ChildItem -Path  ($tmpdir ) -Recurse | Remove-Item -force -Recurse
if (-not (Test-Path -Path $tmpdir )) { New-Item -ItemType Directory -Path ($tmpdir )}
write-progress -Completed $true

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
# Write-Verbose ("Excludetypes".padright($padder-8) + ": $excludetypes")

if ($CloudAPI -and $CloudAuthAvailable) {
    if (-not (LoginSewnetCloud)) {
        write-host "Login failed or canceled - Stopping" -ForegroundColor Red
        write-progress -Completed $true
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


$cont = (MyPause 'Press Start to continue, any other key to stop (Auto starting in 3 seconds)'  $true 'Click Yes to start' 10) 

if (!$cont) { 
    Break
    }
$beginTimer = Get-Date
Add-Type -assembly "system.io.compression.filesystem"

ShowProgress  "Calculating size"
$librarySizeBefore = 0
$libraryEmbSizeBefore = 0
Get-ChildItem -Path ($EmbroidDir)  -Recurse -file  | ForEach-Object { 
    $librarySizeBefore +=  $_.Length
    if ($_.Extension -and $preferredSewType.Contains($_.Extension.Substring(1)))  {  $libraryEmbSizeBefore += $_.Length }
}
write-host "Starting with All files: $(niceSize $librarySizeBefore) - Embroidery files:  $(niceSize $libraryEmbSizeBefore)"

ShowProgress  "Loading file list"
# Get a list of all the existing files in mySewnet
$mysewingfiles = LoadSewfiles
$quickmysewfiles = BuildHashofMySewingFiles

if ($FirstRun) {
    $DownloadDaysOld = 20*365
}
# $mysewingfiles | ft
$isNewInstruct = $false

$havewarning = $false
$afterdate = (Get-Date).AddDays(- $DownloadDaysOld )
Get-ChildItem -Path $downloaddir  -file -filter "*.zip" | Where-Object { (($_.CreationTime -gt $afterdate) -OR ($_.LastWriteTime -gt $afterdate)) -and ($_.gettype().Name -eq 'FileInfo')} |
  
    ForEach-Object {
        ShowProgress  "Checking Zips - Looking at $($_.Name)"  -stat "Added $Script:savecnt files"
        Write-Verbose "Checking ZIP '$($_.FullName)'" 
        ProcessZipContents -zips $_.FullName -Base $_.BaseName
    }

# Look for Files which are not part of a ZIP file, just the selected file types that we are looking for that is in the download directory
$DownloadDaysOld = 365*10  # 10 years of downloads (when you download files, it keeps the old data)
$ppp = 0
foreach ($thistype in $preferredSewType) {

    
    write-Information "Working on File type: *.$ts"
    Get-ChildItem -Path $downloaddir  -file  -Depth 1 -Recurse| Where-Object { $_.Extension -like $thistype -and $_.CreationTime -gt (Get-Date).AddDays(- $DownloadDaysOld ) } |
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
                        $fullname | Copy-Item -Destination $(join-path -Path $NewFilesDir -ChildPath $f) -ErrorAction SilentlyContinue
                    } else {
                        $fullname | Copy-Item -Destination $NewFilesDir  -ErrorAction SilentlyContinue
                    }
                    LogAction $f -Action "++Added-from-Download"
                    
                    Write-host $("* New  : '$f'").padright(65) " 1 new pattern" -foreground green
                    Write-Information "+++ Copied from Downloads :'$($_.Name)' to $EmbroidDir"
                    $fd = join-Path -Path $d -ChildPath $($fs + "." + $thistype)
                    if (test-path -path $fd) {
                        Copy-Item -Path $fd -Destination $InstructDir -ErrorAction SilentlyContinue 
                        Write-Information "+++ Copied instructions from Downloads :'$($_.Name)' to $InstructDir"
                        LogAction -File $($_.Name) -Action "++Added-from-Download" -isInstrution $true
                        }
                    if ($NoDirectory) {
                        $l = ""
                    } else {
                            $l = (split-path -Path $fullname -Parent).Substring($downloaddir.Length).trim('\')
                    }
                    $d = (join-path -path $EmbroidDir -childpath $l).Trim('\')
                    AddToSewList -NameIndex $findfile -Name $f -Directory $d -LastWriteTime $thisfile.LastWriteTime -keepPath $thisfile -RelativePath $l
                    
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

if ($CleanCollection) {
    write-host "Scanning for files to clean up in cache"

    # this creates a BUG for below as we are removing files that might be synced without removeing them from the list
    $filesToRemove = DuplicateFiles -Path $EmbroidDir 
    if ($filesToRemove) { 
        CheckAndRemove -RemoveFiles $filesToRemove -DeleteWithoutRecycle $HardDelete -why "you have multiple copies of the same file"
        }
    if (!$KeepAllTypes) {
        write-host "Scanning for folders to clean up in cache"
        $pe = $preferredSewType | ForEach-Object { ".$_" }
        $filesToRemove = DuplicateFileNames -Path $EmbroidDir -ExtensionsOrder $pe
        if ($filesToRemove) { 
            CheckAndRemove -RemoveFiles $filesToRemove -DeleteWithoutRecycle $HardDelete -why "you have multiple files of different embroidery types"
            }
        }
        $Movelist = @()
        # Look for lone directories
        $maxtries = 15
        do {
            $foundone = $false
            Get-ChildItem -Path $EmbroidDir -Recurse -Directory | Where-Object { $_.GetFiles().count -eq 0 -and $_.GetDirectories().count -eq 1} | ForEach-Object {
                # because we moved the directory in a previous iteration, this maybe null
                if (test-path -path $_) {
                    $thisdir = get-item $_
                } else {
                    $thisdir = $null
                }
                if ($thisdir -and $thisdir.GetDirectories().count) {
                    ShowProgress -Area $_ -stat "Moving Lone Directories"
                    $subdir = $thisdir.GetDirectories()
                    $subitem = get-item -path $subdir
                    if ($subitem.GetFileSystemInfos().count -eq 0) {
                        if (-not $KeepEmptyDirectory) {
                            remove-item $subdir
                        }
                    } else {
                        $Movelist += [PSCustomObject]@{
                            From = $subdir.FullName
                            To = $thisdir.FullName
                        }
                        move-item -Path $($subdir.FullName + "\*") -Destination $thisdir.FullName -force
                        # write-host "move from $subdir to $thisdir"
                        $foundone = $true
                    }
                }
            }
            # remove the items where we move it From X To Y To Z ... remove 'From X'
            $MoveList = $moveList | where-object { $_.From -notin $moveList.To }
            #then we need to fix the mysewing files list
            $mysewingfiles | Where-Object {$_.DirectoryName -in $MoveList.From } | ForEach-Object {
                $MoveTo = $_.DirectoryName
                foreach ($MoveLoc in ($MoveList)) {
                    if ($MoveLoc.From -like $_.DirectoryName) {
                        $MoveTo = $moveLoc.To
                        break
                    }
                }
                $_.DirectoryName = $MoveTo
                $_.FullName =  join-path -path $MoveTo -ChildPath $_.Name
                
            } #>

            $maxtries--
        } while ($foundone -and $maxtries)
        

}    


#  Clear out empty Directories
if (-not $KeepEmptyDirectory) {
    $tailr = 0    # Loop thru 8 times to remove empty directories, then go back and check to see if you made any more emty
    while ($tailr -le 8 -and (tailRecursion $EmbroidDir) ) {
         $tailr++
    }
}

$script:lostfiles | Out-GridView -Title "Lost Files" 

if ($CloudAPI -and $CloudAuthAvailable) {
    $webcollection = ReadCloudMeta
    if ($null -eq $webcollection) {
        write-host "Cloud is not working *** STOPPING (Try logging onto MySewnet before retrying)" -ForegroundColor Red
        MyPause 'Press any key to Close' | out-null
        return
    
    }
    # TODO Check the Cloud to see if we have space to load the new files.

    $cm = 0
    $cf = 0
#     $tolist = @()

    $MySewingfiles | ForEach-Object {
        $thisfile = $_
        $cm = $cm + 1
        if ( $cm % 50 -eq 0) {
            write-Progress -Activity "Matching files to Cloud :  $($thisfile.N)" -PercentComplete ($cm * 100 / $mysewingfiles.count) -Status "$cm of $($mysewingfiles.count)"
        }
        $thisfile.CloudRef = GetFileIDfromCloud $_.N
        if ($EmbroidDir.Length -lt $thisfile.DirectoryName.length) {
            $spl = $thisfile.DirectoryName.substring($EmbroidDir.Length)
        } else {
            $spl = ""
        }
        $thisfile.RelPath = $spl
        
        if ($thisfile.CloudRef)   {
            $cf = $cf +1
            }
        }
    write-Progress -Activity "Matched files to Cloud " -complete $true
    write-host "Found $cf cloud file which match local cache of $($MySewingfiles.count) files"
    # $MySewingfiles | Out-GridView -Title "MySewingfiles List"



    if ($sync) {
        $pool = $webcollection
        $cloudfileremove = @()
        do {
        $cloudfileremove  += $pool | Select-Object -ExpandProperty Folders | select-Object -ExpandProperty Files| Where-Object {$_.Name -notin ($mysewingfiles.N)} 
        $pool = $pool | Select-Object -ExpandProperty Folders 
        } while ($pool.files.count + $pool.folders.count -gt 0)
        if ($cloudfileremove.count -gt 10) {
            write-host "Removing $($cloudfileremove.count) files from Mysewnet Cloud, this is going to take some time... " -ForegroundColor Yellow
        }
        $i = 0
        if ($cloudfileremove) {
            $cloudfileremove | ForEach-Object  {
                # write-host " --Removing from the cloud" $_.Name
                Write-Progress -PercentComplete (($i++)*100/$cloudfileremove.count) "Removing files from cloud :" -Status $_.Name
                LogAction -File $_.Name -Action "--Deleted-Sync"
                DeleteCloudFile -id $_.id | Out-Null
            }
        }
        $filestopush = ($mysewingfiles | Where-Object { ($_.Push -and $_.Push.contains('\'))  -or ($_.CloudRef -eq $null)}).count

        }
     else {
        $filestopush = ($mysewingfiles | Where-Object { ($_.Push -and $_.Push.contains('\')) }).count
        }

    if (-not $KeepEmptyDirectory) {
        $totalfoldersremoved = 0
        # BUG need CloudDeleteFolder to manage deletions
        #do {
            $emtpyfoldersList = @()
        
            $pool = $webcollection
            do {
                $emtpyfoldersList  += $pool | Select-Object -ExpandProperty Folders |  Where-Object {$_.files.count -eq 0 -and $_.folders.count -eq 0} 
                $pool = $pool | Select-Object -ExpandProperty Folders 
                } while ($pool -and $pool.folders.count -gt 0)

            if ($emtpyfoldersList.count -gt 0) {
                $totalfoldersremoved += $emtpyfoldersList.count
                write-host "Clearing Empty Cloud folders : $($emtpyfoldersList.count) / $totalfoldersremoved"
                $emtpyfoldersList | ForEach-Object { 
                    ShowProgress "Delete Cloud Folder $($_.Name)"
                    DeleteCloudFolder -id $_.id | Out-null
                    }
            }
         #   } while ($emtpyfoldersList.count -gt 0)
        }
    if ($script:CloudStatusGood) {
        write-host "Beginning push to MySewnet: $filestopush files" -ForegroundColor Green
        $i = 0
        if ($filestopush -or $sync) {
            $MySewingfiles | ForEach-Object {
                $thisfile = $_
                if ($thisfile.push -and $thisfile.push.contains("\") ) {
                    Write-Progress -PercentComplete (($i++)*100/$filestopush) "Pushing files to the Cloud : " -Status $thisfile.N
                    PushCloudFileToDirectory -filepath ($thisfile.FullName) -folderpath $thisfile.push | Out-Null
                    LogAction -File ($thisfile.push + "\" + $thisfile.N) -Action "^^Cloud-Added"
                    $thisfile.push = ""
                    }
                if ($sync) {

                    if ($thisfile.CloudRef) {
                                        
                        $samePath = FindCloudidfromPath -foldername $thisfile.RelPath
                        # Check to see if the proper path can not be found
                        if ($samePath -eq $null) {
                            $samePathid = ""
                            if ($thisfile.RelPath -ne '') {
                                write-verbose "Making New Path : $($thisfile.RelPath) " 
                                $samePathid = MakeCloudPathID -path $thisfile.RelPath
                            }    
                        } else {
                            $samePathid = $samePath.id
                        }
                        
                        
                        if ($thisfile.CloudRef.FolderId -ne $samePathid -and $isOkToMove) {
                            write-verbose ")) Relocated $($thisfile.N) from $($thisfile.CloudRef.FolderId) to $($samePathid)"
                            MoveCloudFile -fileid $thisfile.CloudRef.Id -toFolderid $samePathid
                            LogAction -File ($thisfile.RelPath + "\" + $thisfile.N) -Action "^^Cloud-Move"
                        }

                    } else {
                        Write-Progress -PercentComplete (($i++)*100/$filestopush) "Syncing files to the Cloud : " -Status $thisfile.N
                        $spl = $thisfile.DirectoryName.substring($EmbroidDir.Length)
                        PushCloudFileToDirectory -filepath ($thisfile.FullName) -folderpath $spl | Out-null
                        LogAction -File ($spl + "\" + $thisfile.N) -Action "^^Cloud-Added-Sync"
                        }
                    }
                    if (!($script:CloudStatusGood)) {
                        break
                    }
                    
                }
                Write-Progress -completed $true "Pushed files to Cloud "
            }
        }
    }

    
write-host "Calculating size"


$librarySizeAfter = 0
$libraryEmbSizeAfter = 0
$byExt = @{}
if ($Script:savecnt -gt 0) {
    if (-not $CloudAPI) {
        OpenForUpload
    }
    Get-ChildItem -Path ($EmbroidDir)  -Recurse -file  | ForEach-Object { 
        $librarySizeAfter +=  $_.Length
        if ($_.Extension -and $preferredSewType.Contains($_.Extension.Substring(1)))  {  $libraryEmbSizeAfter += $_.Length }
    }
    Get-ChildItem -Path ($EmbroidDir  ) -Recurse -file  | ForEach-Object { 
        $librarySizeAfter = $librarySizeAfter + $_.Length
        $byExt[$_.Extension] +=  $_.Length
    }
} else {
    if (-not $CloudAPI) {
        if (MyPause "Do you want to open the web page & directory to upload files from last time?" $true) {
            OpenForUpload
        }
    }
}


  
$addsizecntB = niceSize $Script:addsizecnt
write-progress -PercentComplete  100  "Done"
if ($Script:dircnt -gt 0 -or $filesToRemove.length -gt 0) {
    $filecnt = $filesToRemove.length
    Write-Host "Cleaned up - Directories removed: '$Script:dircnt    Files removed : '$filecnt' ($sizecntB)." -ForegroundColor Green
    }
if ($Script:savecnt -gt 0) {
    write-host "+++ Added files to Embriodery Collection: '${Script:savecnt}' files $(niceSize $Script:addsizecnt) " -ForegroundColor Green
    write-host "File size after All: $(niceSize $librarySizeAfter) - Embroidery files: $(niceSize $libraryEmbSizeAfter)"

    }
else {
    Write-host "   *** Instructions size is   : $( ) ****   "  -ForegroundColor Green 
    write-host "   *** Embroidary file size is: $(niceSize $libraryEmbSizeBefore) ****" -ForegroundColor Green 
}
if ($CloudAPI -and $CloudAuthAvailable) {
    if ($script:CloudStatusGood) {
        $updatedMeta = ReadCloudMeta
        write-Host "Cloud Storage currently is: " -ForegroundColor Green
        write-Host "       Used of Total:".padright(25) (NiceSize $updatedMeta.storage.usedSize) "/" (NiceSize $updatedMeta.storage.totalSize)
        write-Host "       Space remaining:".padright(25) (NiceSize $updatedMeta.storage.availableSize)
    } else {
        write-Host "Cloud unavailable or errored during processing - try again"
    }
}
 # $mysewingfiles | out-GridView
 # $byExt | Out-GridView
# Capture the end time
$endTimer = Get-Date

# Calculate the difference
$timeSpan = $endTimer - $beginTimer

# Display the difference in minutes and seconds
$minutes = [math]::Floor($timeSpan.TotalMinutes)
$seconds = [math]::Round($timeSpan.Seconds, 2)
Write-Host "Runtime: $minutes minutes and $seconds seconds" -ForegroundColor Blue

write-progress -Completed $true
Write-Host ( "End") -ForegroundColor Green
MyPause 'Press any key to Close' | out-null
