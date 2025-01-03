﻿#Requires -Version 5.1
<# 
 EmbroideryCollection-Cleanup.ps1
    GPL-3.0 license

 Deal with the many different types of embroidery files, put the right format types in mySewingnet Cloud or onto a USB drive
 We are looking to keep the ??? files types only from the zip files.

 Orginal Author: Darren Jeffrey, Dec 2021 - Oct 2024
#>

param
(
  [Parameter(Mandatory = $false)]
  [int32]$DownloadDaysOld,                      # How many days old show it scan for Zip files in Download
  [int32]$SetSize = 5,
  [Switch]$KeepAllTypes,                            # Keep all the different types of a file (duplicate name but different extensions)
  [Switch]$CleanCollection,                         # Cleanup the Collection Folder to only EmbroidDir files
# incomplete NoDirectory
#  [Switch]$NoDirectory,                             # Do not create directory structure in the upload from space
#  [Switch]$OneDirectory,                              #Limit the folders to one directly deep only 
  [string]$EmbroidDir = "Embroidery",               # You may want to change this directory name inside of the 'Collection ' Directory wihtin your 'Documents' directory
                                                    # If it is just a name, then it is assumed to be within the defeault Documents directory, otherwise it will be taken as a full path
  [string]$USBDrive,                                # Write the output to a USB Drives
  [switch]$USBEject,                                # Write the output to a USB Drives
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
  [Switch]$CloudAPI,                                # use MySewNet cloud API
  [String]$OldVersion                               # Used during upgrade only
  )

$ECCVERSION = "v0.8.2"
$GitOwner = "D-Jeffrey" 
$GitName = "Embroidery-File-Organize"
# $VerbosePreference =  "Continue"
# $InformationPreference =  "Continue"

<# ********     CONFIGURATION    ********#>
$preferredSewType = 'vp4', 'vp3',  'pes', 'pcs','hus','dst'
$alltypes = 'art', 'art50', 'art60', 'art70', 'art80', 'asd', 'bmc', 'cnd', 'csd', 'dem', 'dst', 'dxf', 'emb', 'emd', 'ese', 'exp', 'gnc', 'hus', 'jef', 'jef+', 'jpx', 'mhv', 'oef', 'ofm', 'pcd', 'pcm', 'pcq', 'pcs', 'pec', 'pel', 'pem', 'pes', 'phb', 'phc', 'pxf', 'sew', 'shv', 'svg', 'tap', 'vf3', 'vip', 'vp3', 'vp4', 'xxx', 'zsk'
$defaultfoldupDir = 'images', 'sewing helps', 'Designs', 'Design Files', 'brother-babylock-pes', 'janome-jef', 'singer-xxx', 'husqvarna-viking-hus', 'commercial formats - dst-exp', 'artista-art','vp'

$defaultgoodInstructionTypes = 'avi', 'bmp', 'doc', 'docx', 'emf', 'gif', 'htm', 'html', 'jpg', 'm4v', 'mov', 'mp4', 'mpg', 'pcx', 'pdf', 'png', 'ppt', 'pptx', 'rtf', 'tif', 'txt', 'wmf', 'wmv'
$defaultTandCs = 'TERMS-OF-USAGE.*', 'planetappliquetermsandconditions.*','READ ME FIRST.rtf','*copyright.*','*copyright Statement.*','*copyrights.*',
    'copyrightStatement.*','License agreement.*', 'License.*','termsofuse.*', 'Thumbs.db'

<# ********     END CONFIGURATION  - active settings kept in EmbroideryCollection.cfg  ********#>

# List of paramstring to check
$paramstring =  [ordered]@{
 'EmbroidDir' = 'Embriodary Files directory';
 'USBDrive'='USB drive letter (example E: or H:)';
 'LastCheckedGithub'=''; 
 'DownloadDaysOld' = 'Age of files in Download directory';
 'SetSize' = 'Keep collections of files together if there are at least this many'
}

$parambool = [ordered]@{
'KeepAllTypes'= 'Keep all variations of files types' ; 
'KeepEmptyDirectory'= 'When cleaning up keep empty folders'; 
'DragUpload'= 'Open the mysewnet Cloud browser interface for drag and drop';
'ShowExample'= 'Show how to upload to mySewnet';
'NoDirectory'= 'Do not use Directories from Zip files which creating collection';
'OneDirectory'= 'Keep files a maximum of one directory deep ';
'CloudAPI'= 'Use MySewnet Cloud';
'USBEject'='Safely Eject the USB drive when complete'
}
$paramarray = [ordered]@{
'preferredSewType' = 'The preferred types of Embriodary file types';
'alltypes' = '* All the possible types of files which are an Embriodary file'; 
'foldupDir' = '* Remove/fold folders of this name'; 
'goodInstructionTypes' = '* Instructions file types which should be saved with files';
'TandCs' = '* Readme, Copyright files and T&C' 
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
write-progress "Starting Embroidery Collection Cleanup version: $ECCVERSION on PS $($PSVersionTable.PSVersion.major).$($PSVersionTable.PSVersion.minor) ... Please wait" 

$RemovePrefix = ($PSVersionTable.PSVersion.Major -lt 7 ) 
if ($RemovePrefix) {
    write-host " ".padright(15) "    (Runs better with a new version of Powershell)".padright(70) -ForegroundColor Yellow -BackgroundColor Blue
    }
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
$script:markdrive = "eCollection.txt"

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

$missingSewnetAddin = (get-itemproperty -ErrorAction SilentlyContinue -path "Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Classes\CLSID\{E2F80510-8B1C-4A23-B657-D173D4F6E418}")."(default)" -notlike "VsmSoftware*"
$script:protime = $(get-date)
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


####################### CLASS ###################



####################### Functions ###################
#-#-#-#-#-

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

Function PrepareCloud() {
    $script:CloudAuthAvailable = $false
    # $Script:DragUpload = $false
    $Script:ShowExample = $false
    if ($CloudAPI) {
        AdvanceProgress "Preparing Authenication Module" -BigStep
        $script:CloudAuthAvailable = ((Get-Module -Name PSAuthClient).count -gt 0)
        if (-not $script:CloudAuthAvailable) {
            AdvanceProgress "Checking for Authenication Module" -BigStep
            
            if ((Get-Module -ListAvailable | Where-Object { $_.Name -eq "PSAuthClient" }).count -gt 0) {
                Import-Module PSAuthClient 
                $script:CloudAuthAvailable = ((Get-Module -Name PSAuthClient).count -gt 0)
            }
            else {
                AdvanceProgress "-- Please wait while the script installs missing module --" -BigStep
                Install-Module -name PSAuthClient -scope:CurrentUser
                Import-Module PSAuthClient 
                $script:CloudAuthAvailable = ((Get-Module -Name PSAuthClient).count -gt 0)
                write-host "Completed" -foreground Yellow
            }
            Complete-Progress
        }
    }
    
}

function Show-Progress { 
    param ( 
        [Parameter(Mandatory = $false)]
        [string]$Activity, 
        [string]$Status, 
        [int]$PercentComplete, 
        [Switch]$Completed
        ) 
    $params = @{}
    if ($(get-date).AddSeconds(-1) -gt $script:protime) {
        if ($Activity) { $params.Activity = $Activity }
        if ($Status) { $params.Status = $Status }
        if ($PercentComplete) { $params.PercentComplete = [Math]::max(1,[Math]::min(100,$PercentComplete))}
        if ($Completed) { $params.Completed = $Completed}
        Write-Progress @params
        $script:protime = $(get-date)
        }
}
function Complete-Progress {
    $script:protime = (get-date).AddMinutes(-1)
    Show-Progress -Completed $true
}


if ($missingSewnetAddin) {
    $DragUpload = $true
}

    
#=============================================================================================

function LogAction($File, $Action, [Boolean]$isInstructions = $false) {
    $now = Get-Date -Format "yyyy/MMM/dd HH:mm "
    $extra = (&{if ($isInstructions) { " Instructions"} else { "Embroidery" } })
    write-verbose "$Action $File type: $extra"
    Add-Content -Path $LogFile -Value ("$now$Action $File $extra")
}

Function AdvanceProgress {
    param ( 
        [Parameter(Mandatory = $false)]
        [string]$Area, [string]$stat = $null, [switch]$BigStep)

    $Script:p++
    if ($BigStep) {
        $Script:p = $Script:p + 9 ; 
        $script:protime = (get-date).AddMinutes(-1)
    }
    if ($stat) {
        Show-Progress -PercentComplete ($Script:p % 100 ) $Area -Status $stat
    } else {
        Show-Progress -PercentComplete ($Script:p % 100 ) $Area 
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
                Remove-Item -Path "\\?\$($file.FullName)"    # Handled by WhatIf
            } else {
                Remove-Item -Path $file.FullName             # Handled by WhatIf
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
        [bool]$ChoiceDefault = $true,
        # If this is included, then the first key will be used for it and this value will be returned instead of a true/false
        [string]$extraChoice = ""           
        
        # [bool]$YesNoCancel = $false
    )

    $bKeys = ""
    $yes = $true
    # Check if running Powershell ISE
    if ($psISE -or $useGUI) {
        Add-Type -AssemblyName System.Windows.Forms
        $BoxMsg = if ($BoxMsg -eq "" -or $null -eq $BoxMsg) { $Message } else { $BoxMsg }

        if ($Choice) {
            $x = [System.Windows.Forms.MessageBox]::Show($BoxMsg, 'Embroidery Collection Cleanup', 'YesNo', 'Question')
        } else {
            [System.Windows.Forms.MessageBox]::Show($BoxMsg, 'Embroidery Collection Cleanup')
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
            if ($extraChoice) {
                    $extraKeys = "/" + $extraChoice.ToLower()
                    $bKeys = $extraChoice.substring(0,1).ToUpper() +  $extraChoice.substring(0,1).ToLower()
            } else { $extraKeys = ""   }
            $yesno = (&{if ($ChoiceDefault) { " (Y/n$extraKeys) " } else { " (y/N$extraKeys) " }})
            
        } else {
            $yesno = ""
        }
        Complete-Progress
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
                    $needakey = -not ("YyNn$bKeys".Contains($keystroke.Character))
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
            if ($extraChoice -and $bKeys.Contains($keystroke.Character) ) {
                $selectchoice = $extraChoice
                $yes = $extraChoice
            }
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

function doWinForm() {
    Add-Type -AssemblyName System.Windows.Forms
    $ico = "C:\ProgramData\EmbroideryOrganize\EmbroideryManager.ico"
    
    # Create the form
    $form = New-Object System.Windows.Forms.Form -Property @{
        Text = "Embroidery Collection Cleanup"
        Size = New-Object System.Drawing.Size(800, 500)
        StartPosition = "CenterScreen"
        BackColor = "Window"
        Font = "Calibri, 9pt"  
    }
    if (test-path -path $ico) {
        $form.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($ico)
    }
    
    $lbl_title = New-Object System.Windows.Forms.Label -Property @{
        BackColor="Blue"
        Font="Arial Black, 16pt"
        ForeColor="White"
        Size= New-Object System.Drawing.Size(800, 32)
        #TextAlign="TopCenter" 
        Text = "Embroidery Collection Cleanup version: $ECCVERSION (on PS $($PSVersionTable.PSVersion.major).$($PSVersionTable.PSVersion.minor))"
    }
    $form.Controls.Add($lbl_title)
    
    # Create a panel to hold the parameters
    $panel = New-Object System.Windows.Forms.Panel -Property @{
        Location = New-Object System.Drawing.Point(125, 34)
        Size=  New-Object System.Drawing.Size(668, 450)
        Font="Tahoma,9pt"
    }
    $lbl_EmbrodRootDirtop = New-Object System.Windows.Forms.Label -Property @{ 
        Location = New-Object System.Drawing.Point(4, 42)
        Size=  New-Object System.Drawing.Size(120, 21)
        Text="Embroidery base dir" 
    }
    $lbl_DownloadDaysOld = New-Object System.Windows.Forms.Label -Property @{ 
        Location = New-Object System.Drawing.Point(4, 66)
        Size=  New-Object System.Drawing.Size(120, 21)
        Text="Download days old"
    }
    $lbl_KeepAlltypes= New-Object System.Windows.Forms.Label -Property @{ 
        Location = New-Object System.Drawing.Point(4, 90)
        Size=  New-Object System.Drawing.Size(120, 21)
        Text="Keep all types"
    }
    $lbl_preferredSewType= New-Object System.Windows.Forms.Label -Property @{ 
        Location = New-Object System.Drawing.Point(4, 114)
        Size=  New-Object System.Drawing.Size(120, 42)
        Text="Preferred sew file`nsort by preference"
    }
    $lbl_otherSewType= New-Object System.Windows.Forms.Label -Property @{ 
        Location = New-Object System.Drawing.Point(260, 98)
        Size=  New-Object System.Drawing.Size(120, 21)
        Text="Other sew file types"
    }
    $lbl_location= New-Object System.Windows.Forms.Label -Property @{ 
        Location = New-Object System.Drawing.Point(4, 234)
        Size=  New-Object System.Drawing.Size(120,21)
        Text="Output to"     
    }
    $lbl_Info= New-Object System.Windows.Forms.Label -Property @{ 
        Location = New-Object System.Drawing.Point(4, 344)
        Size=  New-Object System.Drawing.Size(600,61)
        Text=""
        Padding="5, 5, 5, 5"
        ForeColor="InfoText"
        BackColor="Info"
    }
    $script:Info_Extra = "`n"

    $panel.Controls.Add($lbl_EmbrodRootDirtop)
    $panel.Controls.Add($lbl_DownloadDaysOld)
    $panel.Controls.Add($lbl_KeepAlltypes)
    $panel.Controls.Add($lbl_preferredSewType)
    $panel.Controls.Add($lbl_otherSewType)
    $panel.Controls.Add($lbl_location)
    $panel.Controls.Add($lbl_Info)

    $tbx_EmbrodRootDirtop = New-Object System.Windows.Forms.TextBox -Property @{ 
        Location = New-Object System.Drawing.Point(130, 42)
        Size = New-Object System.Drawing.Size(380, 20)
        Text = $script:EmbroidDir
        }
    $panel.Controls.Add($tbx_EmbrodRootDirtop)
    
    $DirectoryBrowser = New-Object System.Windows.Forms.FolderBrowserDialog -Property @{ 
        SelectedPath = $script:EmbroidDir
        Description = "Select the Directory for Embroidery Files.  This will be used as a reference location for you to look at the current collection of files"
        ShowNewFolderButton = $true
        # https://learn.microsoft.com/en-us/dotnet/api/system.environment.specialfolder
        rootFolder = 40 # CommonDocuments
        }
    $btn_selectDir = New-Object System.Windows.Forms.Button -Property @{ 
        Text = "Select Directory"
        BackColor="LightSkyBlue"
        UseVisualStyleBackColor=$False
        Location = New-Object System.Drawing.Point(511, 39)
        Size = New-Object System.Drawing.Size(99, 23)
        }
    $btn_selectDir.Add_Click({
        $DirectoryBrowser.SelectedPath = $script:EmbroidDir
        # Update the variables with the current values in the controls
        if ($DirectoryBrowser.ShowDialog() -eq "OK") {
            $script:EmbroidDir = $DirectoryBrowser.SelectedPath
            $tbx_EmbrodRootDirtop.Text = $script:EmbroidDir
        }
    })
    $panel.Controls.Add($btn_selectDir)
    $btn_open = New-Object System.Windows.Forms.Button -Property @{ 
        Text = "Open"
        Location = New-Object System.Drawing.Point(613,39)
        Size=  New-Object System.Drawing.Size(43, 23)
        BackColor="LightSkyBlue"   
        }
    $btn_open.Add_Click({
            $odir = $tbx_EmbrodRootDirtop.Text
            Invoke-expression "explorer ""$odir"""
            
        })
    $panel.Controls.Add($btn_open)    
    if ($script:DownloadDaysOld -lt 1) { $script:DownloadDaysOld = 7 }
    $nud_DownloadDaysOld = New-Object System.Windows.Forms.NumericUpDown -Property @{ 
        Location = New-Object System.Drawing.Point(130, 66)
        Size = New-Object System.Drawing.Size(67, 20)
        Value = $script:DownloadDaysOld
        Minimum = 1
        Maximum = 9000
        }
    $panel.Controls.Add($nud_DownloadDaysOld)
        
    $cbx_KeepAllTypes = New-Object System.Windows.Forms.CheckBox -Property @{ 
        Location = New-Object System.Drawing.Point(130, 90)
        Checked = ($true -eq $script:KeepAllTypes)
        }
    $panel.Controls.Add($cbx_KeepAllTypes)
    
    $lbx_preferredSewType = New-Object System.Windows.Forms.ListBox -Property @{ 
        Location = New-Object System.Drawing.Point(130, 120)
        Size = New-Object System.Drawing.Size(68, 108) 
        SelectionMode = [System.Windows.Forms.SelectionMode]::One
        }     
    $panel.Controls.Add($lbx_preferredSewType)
        
    $lbx_otherTypes = New-Object System.Windows.Forms.ListBox -Property @{ 
        Location = New-Object System.Drawing.Point(280, 120)
        Size = New-Object System.Drawing.Size(68, 108) 
        SelectionMode = [System.Windows.Forms.SelectionMode]::One
        
        }     
    $panel.Controls.Add($lbx_otherTypes)
    
    # Populate the ListBox with sample items
    
    $lbx_preferredSewType.Items.AddRange($preferredSewType)
    
    $lbx_otherTypes.Items.AddRange($($alltypes |where-object {$_ -notin $preferredSewType}))
    # Create the Move Up button
    $lbbx_moveUpButton = New-Object System.Windows.Forms.Button -Property @{ 
        Text = "↑"
        Location = New-Object System.Drawing.Point(200, 115)
        Size = New-Object System.Drawing.Size(20, 20)
    }
    $lbbx_moveUpButton.Add_Click({
        $selectedIndex = $lbx_preferredSewType.SelectedIndex
        if ($selectedIndex -gt 0) {
            $temp = $lbx_preferredSewType.Items[$selectedIndex - 1]
            $lbx_preferredSewType.Items[$selectedIndex - 1] = $lbx_preferredSewType.Items[$selectedIndex]
            $lbx_preferredSewType.Items[$selectedIndex] = $temp
            $lbx_preferredSewType.SelectedIndex = $selectedIndex - 1
        }
    })
    $panel.Controls.Add($lbbx_moveUpButton)
    $lbbx_moveAddButton = New-Object System.Windows.Forms.Button -Property @{ 
        Text = "←"
        Location = New-Object System.Drawing.Point(220, 140)
        Size = New-Object System.Drawing.Size(30, 20)
    }
    $lbbx_moveAddButton.Add_Click({
        $selectedIndex = $lbx_otherTypes.SelectedIndex
        if ($selectedIndex -ge 0) {
            $lbx_preferredSewType.Items.Add($lbx_otherTypes.Items[$selectedIndex])
            $lbx_otherTypes.Items.RemoveAt($selectedIndex)
        }
    })
    $panel.Controls.Add($lbbx_moveAddButton)
    $lbbx_moveRmButton = New-Object System.Windows.Forms.Button -Property @{ 
        Text = "→"
        Location = New-Object System.Drawing.Point(220, 165)
        Size = New-Object System.Drawing.Size(30, 20)
    }
    $lbbx_moveRmButton.Add_Click({
        $selectedIndex = $lbx_preferredSewType.SelectedIndex
        if ($selectedIndex -ge 0) {
            $lbx_otherTypes.Items.Add($lbx_preferredSewType.Items[$selectedIndex])
            $lbx_preferredSewType.Items.RemoveAt($selectedIndex)
        }
    })
    $panel.Controls.Add($lbbx_moveRmButton)
    
    # Create the Move Down button
    $lbbx_moveDnButton = New-Object System.Windows.Forms.Button -Property @{ 
        Text = "↓"
        Location = New-Object System.Drawing.Point(200, 190)
        Size = New-Object System.Drawing.Size(20, 20)
    }
    $lbbx_moveDnButton.Add_Click({
        $selectedIndex = $lbx_preferredSewType.SelectedIndex
        if ($selectedIndex -ge 0 -and $selectedIndex -lt $lbx_preferredSewType.Items.Count - 1) {
            $temp = $lbx_preferredSewType.Items[$selectedIndex + 1]
            $lbx_preferredSewType.Items[$selectedIndex + 1] = $lbx_preferredSewType.Items[$selectedIndex]
            $lbx_preferredSewType.Items[$selectedIndex] = $temp
            $lbx_preferredSewType.SelectedIndex = $selectedIndex + 1
        }
    })
    $panel.Controls.Add($lbbx_moveDnButton)
        
        
    if ($script:CloudAPI) {
        $script:USBDrive = "MySewnet"
    } 
    $script:InfoLabel = ""
    ################# Local Functions ##############
    # Query Win32_LogicalDisk class to get information about drives
    function refeshDriveList() {
        # Filter and list removable drives
        $removableDrives = Get-WmiObject -Class Win32_LogicalDisk | Where-Object { $_.DriveType -eq 2 }
        return $removableDrives
    }
    function isDriveReady($drv) {
        return (Get-WmiObject -Class Win32_LogicalDisk | Where-Object { $_.DriveType -eq 2 -and $_.FreeSpace -gt 0 -and $_.DeviceID -like $drv })
    }
    function updateDriveList() {
        $items = @("None", "mySewNet")
        $removableDrives = refeshDriveList
        $items = @()
    
        $items += "None"
        $items += "mySewNet"
        foreach ($drv in $removableDrives.DeviceID) {
            $items += $drv
        }
        # Clear the ComboBox and add the new items
        
        $cmb_location.Items.Clear() 
        $cmb_location.Items.AddRange($items)
        $i = 0
        $driveletter = $script:USBDrive | split-path -Qualifier -ErrorAction SilentlyContinue
        if ($null -eq $driveletter) {
            $driveletter = $script:USBDrive
        }
        $btn_open2.Enabled = $false
        foreach ($sel in $items ) {
            if ($sel -like ($driveletter)) {
                $cmb_location.SelectedIndex = $i
                $btn_open2.Enabled = $true
            }
            $i = $i + 1
        }
        if ("" -eq $USBDrive ) {
            $cmb_location.SelectedIndex = 0
            $btn_open2.Enabled = $false
        }
    } 
    $cmb_location = New-Object System.Windows.Forms.ComboBox -Property @{ 
        Location = New-Object System.Drawing.Point(130, 234)
        DropDownStyle  = "DropDownList"
        }
    function driveComboChange () 
        {   
            if ($btn_go) {
                $btn_go.Enabled = $true
                $btn_sync.Enabled = $true
                $btn_clean.Enabled = $true
                $lbl_Info.Text =  ""
                }
            if ($cbx_usbeject) {
                $d = $cmb_location.Items[$cmb_location.SelectedIndex]
                $lbl_Info.Text = ""
                $cbx_usbeject.Enabled = ($cmb_location.SelectedIndex -gt 1)
                $lbl_usbeject.Enabled = $cbx_usbeject.Enabled
                $btn_open2.Enabled = $cbx_usbeject.Enabled
                $cbx_openmySewnet.Enabled = $cmb_location.SelectedIndex -in (0,1)
                if (($cmb_location.SelectedIndex -eq 1)) {
                    $lbl_Info.Text =  "Note: 'Sync'ing to MySewnet will remove files, in addition to adding them`n"
                }
                elseif (($cmb_location.SelectedIndex -gt 1)) {
                    if (isDriveReady -drv $d) {
                        if (-not $(test-path -path $(join-path -Path $d -ChildPath $markdrive))) {
                            $lbl_Info.Text =  "USB Drive $d has not been used with this for Embrodary files.`nUse the Sync button copy your Embroidery collection to a new stick`n"
                            }
                    } else {    
                        $lbl_Info.Text = "USB Drive $d is not ready`nInsert USB drive and click refresh`n"
                        if ($btn_go) {
                            $btn_go.Enabled = $false
                            $btn_sync.Enabled = $false
                            $btn_clean.Enabled = $false
                        }
                    }
                } 
            }
            if ($cmb_location.SelectedIndex -eq -1 -and $btn_go) {
                    $btn_go.Enabled = $false
                    $btn_sync.Enabled = $false
                    $btn_clean.Enabled = $false
                    $lbl_Info.Text =  "Select an different output (or insert your USB drive and click refresh)`n"
            }
            if ($btn_sync) {
                $btn_sync.Enabled = ($cmb_location.SelectedIndex -ne 0 -and $btn_go.Enabled) 
            }
        }
    
    $cmb_location.Add_SelectionChangeCommitted( {driveComboChange})
    
    $btn_refreshDir = New-Object System.Windows.Forms.Button -Property @{ 
        Text = "Refresh"
        BackColor="LightSkyBlue"
        UseVisualStyleBackColor=$False
        Location = New-Object System.Drawing.Point(260, 234)
        Size = New-Object System.Drawing.Size(59, 23)
        }
    $btn_refreshDir.Add_Click({
        $sav = $cmb_location.SelectedIndex
        if ($cmb_location.SelectedIndex -ge 0) {
            $cmb_location.Items[$cmb_location.SelectedIndex]
            $script:USBDrive = $cmb_location.Items[$cmb_location.SelectedIndex]
        }
        updateDriveList
        driveComboChange
        if (-1 -eq $cmb_location.SelectedIndex) {
            $cmb_location.SelectedIndex =$sav
        } 
    })
    
    $panel.Controls.Add($btn_refreshDir)
    $panel.Controls.Add($cmb_location)
    $btn_open2 = New-Object System.Windows.Forms.Button -Property @{ 
        Text = "Open"
        Location = New-Object System.Drawing.Point(329,234)
        Size=  New-Object System.Drawing.Size(43, 23)
        BackColor="LightSkyBlue"   
        Enabled = $false
        }
    $btn_open2.Add_Click({
            $odir = $cmb_location.Items[$cmb_location.SelectedIndex]
            Invoke-expression "explorer ""$odir"""
            
        })
    $panel.Controls.Add($btn_open2)    
    
    $lbl_usbeject= New-Object System.Windows.Forms.Label -Property @{ 
        Location = New-Object System.Drawing.Point(24, 264)
        Size=  New-Object System.Drawing.Size(100,31)
        Text="Safely Eject USB when done"
    }
    $panel.Controls.Add($lbl_usbeject)
    $cbx_usbeject = New-Object System.Windows.Forms.CheckBox -Property @{ 
        Location = New-Object System.Drawing.Point(130, 264)
        Size=  New-Object System.Drawing.Size(15,15)
        Checked = $USBEject
        Enabled = $script:USBEject
        }
    $panel.Controls.Add($cbx_usbeject)
    $lbl_openmySewnet = New-Object System.Windows.Forms.Label -Property @{ 
        Location = New-Object System.Drawing.Point(184, 264)
        Size=  New-Object System.Drawing.Size(100,31)
        Text="Open MySewnet Cloud when done"
    }
    $panel.Controls.Add($lbl_openmySewnet)
    $cbx_openmySewnet = New-Object System.Windows.Forms.CheckBox -Property @{ 
        Location = New-Object System.Drawing.Point(290, 264)
        Size=  New-Object System.Drawing.Size(15,15)
        Checked = $script:DragUpload
        }
    $panel.Controls.Add($cbx_openmySewnet)

    if (test-path -path $(${env:temp} + "\cleansew.new")) {
    
        $btn_open3 = New-Object System.Windows.Forms.Button -Property @{ 
            Text = "Last New`nFiles Folder"
            Location = New-Object System.Drawing.Point(340, 264)
            Size = New-Object System.Drawing.Size(73, 33)
            BackColor="LightSkyBlue"   
            Font = "8pt"
        }
        $btn_open3.Add_Click({
            
            Invoke-expression $("explorer """ + ${env:temp} + "\cleansew.new""")
        })
        $panel.Controls.Add($btn_open3)
    }
    
    # Add the panel to the form
    $form.Controls.Add($panel)
    
    # Create a panel to hold control buttons
    $panelB = New-Object System.Windows.Forms.Panel -Property @{
        Location = New-Object System.Drawing.Point(1, 53)
        Size=  New-Object System.Drawing.Size(120, 420)
        Font="Calibri, 15pt"
    }
      
    
    $exitmode = $Null
    # Create the Edit button
    $btn_go = New-Object System.Windows.Forms.Button -Property @{ 
        Text = "Go"
        Location = New-Object System.Drawing.Point(4,4)
        Size=  New-Object System.Drawing.Size(100, 30)
        ForeColor="White"
        BackColor="RoyalBlue"   
        UseVisualStyleBackColor="False"
        }
    $btn_go.Add_Click({
            $script:exitmode = "Go"
            $form.Close()
        })
    $panelB.Controls.Add($btn_go)
    
    # Create the Sync button
    $btn_sync = New-Object System.Windows.Forms.Button -Property @{ 
        Text = "Sync"
        Location = New-Object System.Drawing.Point(4,87)
        Size=  New-Object System.Drawing.Size(100, 20)
        ForeColor="White"
        BackColor="RoyalBlue"   
        UseVisualStyleBackColor="False"
        Font="8pt"
        }
    $btn_sync.Add_Click({
            $script:exitmode = "Sync"
            $Script:Sync = $true
            $form.Close()    
        })
    $panelB.Controls.Add($btn_sync)
    $btn_help = New-Object System.Windows.Forms.Button -Property @{ 
        Text = "Help"
        Location = New-Object System.Drawing.Point(4,130)
        Size=  New-Object System.Drawing.Size(100, 30)
        ForeColor="White"
        BackColor="RoyalBlue"   
        UseVisualStyleBackColor="False"
        
        }
    $btn_help.Add_Click({
            $script:exitmode = "Help"
            Start-Process "https://github.com/D-Jeffrey/Embroidery-File-Organize/blob/main/help.md" 
        })
    $panelB.Controls.Add($btn_help)
    
    $btn_clean = New-Object System.Windows.Forms.Button -Property @{ 
        Text = "Clean up"
        Location = New-Object System.Drawing.Point(4,268)
        Size=  New-Object System.Drawing.Size(100, 20)
        ForeColor="White"
        BackColor="Darkred"   
        UseVisualStyleBackColor="False"
        Font="8pt"
        
        }
    $btn_clean.Add_Click({
        $script:exitmode = "Clean"
        $script:CleanCollection = $true
        $script:Sync = $true
        $form.Close()
        
    })
    $panelB.Controls.Add($btn_clean)

    $btn_cfg = New-Object System.Windows.Forms.Button -Property @{ 
        Text = "Config"
        Location = New-Object System.Drawing.Point(4,310)
        Size=  New-Object System.Drawing.Size(48, 20)
        ForeColor="White"
        BackColor="Darkred"   
        UseVisualStyleBackColor="False"
        Font="8pt"
        }
    $btn_cfg.Add_Click({
            $script:exitmode = "Config"
            Start-Process "notepad" -ArgumentList $ConfigFile 
        })
    $panelB.Controls.Add($btn_cfg)
    $btn_log = New-Object System.Windows.Forms.Button -Property @{ 
            Text = "Log"
            Location = New-Object System.Drawing.Point(56,310)
            Size=  New-Object System.Drawing.Size(48, 20)
            ForeColor="White"
            BackColor="RoyalBlue"   
            UseVisualStyleBackColor="False"
            Font="8pt"
            }
    $btn_log.Add_Click({
                $script:exitmode = "Config"
                Start-Process "notepad" -ArgumentList $LogFile 
            })
    $panelB.Controls.Add($btn_log)

    # Create the Exit button
    $btn_exit = New-Object System.Windows.Forms.Button -Property @{ 
        Text = "Exit"
        BackColor="RoyalBlue" 
        ForeColor="White" 
        Location = New-Object System.Drawing.Point(4,360)
        Size=  New-Object System.Drawing.Size(100, 30)
        UseVisualStyleBackColor="False"
        }
    $btn_exit.Add_Click({
        $script:exitmode = "Exit"
        $form.Close()    
        })
    $panelB.Controls.Add($btn_exit)
    $form.Controls.Add($panelB)
    updateDriveList
    driveComboChange


    # Hide Console Window
    # Check if the type already exists
    if (-not ([System.Management.Automation.PSTypeName]'Console.Window').Type) {
        Add-Type -Name Window -Namespace Console -MemberDefinition @'
            [DllImport("Kernel32.dll")]
            public static extern IntPtr GetConsoleWindow();

            [DllImport("user32.dll")]
            public static extern bool ShowWindow(IntPtr hWnd, Int32 nCmdShow);
'@
    }

# Now you can use the added type without errors
    $winstate = [Console.Window]::ShowWindow([Console.Window]::GetConsoleWindow(), 0)
    $script:exitmode = "Exit"
    # Show the form
    $form.ShowDialog()
    [Console.Window]::ShowWindow([Console.Window]::GetConsoleWindow(), $winstate)
    $script:EmbroidDir = $tbx_EmbrodRootDirtop.Text
    
    # $script:USBDrive
    if ($cmb_location.SelectedIndex -ge 0) {
        $script:USBDrive = $cmb_location.Items[$cmb_location.SelectedIndex]
        }
    if ($USBDrive -eq "MySewnet") {
        $script:CloudAPI = $true
        $script:USBDrive = ""
        $script:UsingUSBDrive = $false
    } else { 
        $script:CloudAPI = $false
        if ($USBDrive -eq "None") { $script:USBDrive = "" }
        $script:UsingUSBDrive = ($USBDrive -ne "")
    }
    $script:SetExiting = ("Exit" -eq $script:exitmode)
    if (-not $script:SetExiting) {
        SetNewFilesDir
    }
    $script:DownloadDaysOld = $nud_DownloadDaysOld.Value
    $script:KeepAllTypes = $cbx_KeepAllTypes.Checked
    $script:preferredSewType = $lbx_preferredSewType.Items
    $script:USBEject = $cbx_usbeject.Checked
    $script:DragUpload = $cbx_openmySewnet.Checked -and $cbx_openmySewnet.Enabled
    # $script:exitmode
    
    BuildTypeLists
    return $SetExiting
}
function EjectAndCheckUSB {

    # Eject the USB drive
    $ejectResult = (New-Object -ComObject Shell.Application).Namespace(17).ParseName($usbDrive).InvokeVerb("Eject")
    # Check if the USB drive has been ejected successfully
    if ($null -eq $ejectResult) {
        # Verify that all files are flushed
        $fileSystem = Get-WmiObject -Query "SELECT * FROM Win32_LogicalDisk WHERE DeviceID='$usbDrive'"
        if ($null -eq $fileSystem.FreeSpace) {
            Write-host "All files have been written and the USB drive is ready to be safely removed."  -ForegroundColor Green
            return $true
        }
    } else {
        Write-host "Failed to eject USB drive $usbDrive." -ForegroundColor Red
    }
    return $false
}
function Test-ExistsOnPath {
    param (
        [string]$FileName
    )
    
    $env:Path.Split([System.IO.Path]::PathSeparator) | where-object {$_ -ne ""} | ForEach-Object {
        $fullPath = Join-Path $_ $FileName
        if (Test-Path $fullPath -PathType Leaf) {
            return $true
        }
    }

    return $false
}

# return the relative path of existing folders and files relative to a root path with the prefix of  .\
function RelativeDirectory {
    param (
        [string]$Path,
        [string]$rootPath
        )
    $rootUri = [System.Uri]::New($rootPath)
    $directoryUri = [System.Uri]::New($Path)
    $relativeUri = $rootUri.MakeRelativeUri($directoryUri)
    $relativePath = $relativeUri.ToString().Replace('/', '\')

    return ".\" + $relativePath
}

# Look within the directory and find files of the same name and return that list
function DuplicateFiles($Path) {
    # Initialize an empty list to store the file objects
    $FileList = @()
    $Excludes = ($allTypesStar + $TandCs )
    AdvanceProgress   "Get file list for duplicate files in different directories ...."   -BigStep 
    # Get all the files in the directory and sub-directories recursively
    $Files = Get-ChildItem -Path $Path -Recurse -include $Excludes -File
    # Group the files by their name and extension
    AdvanceProgress   "Sorting for duplicate files in different directories .... Please wait" -BigStep
    $FileGroups = $Files | Group-Object -Property Length | Where-Object count -gt 1 
    AdvanceProgress   "Checking for duplicate files in different directories by name" -BigStep
    
    $FileGroups = $FileGroups.Group | Group-Object -Property Name | where-object count -gt 1
    AdvanceProgress   "Checking for duplicate files in different directories now by hash... this will take time... Please wait" -BigStep
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
    # Filter it down to just the known files types.  Then sort by the preferred types
    $Files = Get-ChildItem -Path $Path -Recurse -File |
                    where-object {$_.Extension.replace(".","") -in ($alltypes + $preferredSewType)}
    
    
    # If the preferred extensions list is not empty, check for duplicate names with different extensions
    if ($ExtensionsOrder.Count -gt 0) {
        # Group the files by their base name (without extension)
        AdvanceProgress   "Sorting for extra unneeded files (based on your preferred file types).... Please wait" -BigStep
        $NameGroups = $Files | Group-Object -Property Directory,BaseName # ,LastWriteTime.date
        # Loop through each group of files
        $dupcnt = $($NameGroups | where-object count -gt 1).count
        $cntstep = 0
        foreach ($NameGroup in $($NameGroups | where-object count -gt 1)) {
            $cntstep++
            if (($sp++ % 20) -eq 0) {
                Show-Progress -Activity "Checking for unneeded formats" -PercentComplete $($cntstep*100/$dupcnt) -Status "$cntstep of $dupcnt"
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
        Complete-Progress
        $FileList  | Sort-Object -Property @{Expression = {(&{if ($ExtensionsOrder.IndexOf($_.Extension) -ne -1) { $ExtensionsOrder.IndexOf($_.Extension) } else {100} })}; Descending = $false}   | Group-Object -Property BaseName | where-object count -gt 1 | 
            Out-GridView -Title "Additional instances of these files will be removed - first instances is kept additional are removed" 

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
    $ttl = $RemoveFiles.count
    if ($ttl  -gt 0) {
        write-host "Found $fcr files that $why and should be removed" -ForegroundColor Yellow
        # $RemoveFiles|Select-Object Name, FullName, DirectoryName, Extension | Out-GridView -Title "Files that will be removed - $why (Close this Windows to continue)" 
        $cont = MyPause 'Remove those files? (No to keep them)'  -Choice $true -BoxMsg 'Click Yes to remove them' -ChoiceDefault $false

        if ($cont) {
            if (!$DeleteWithoutRecycle -and $fcr  -gt 100) {
                $cont = (MyPause 'This is going to take a while as it moves the files to recycle. Would you like to DELETE the files without being able to recover them?'  $true 'Click Yes to for a quick delete with NO Recyle!') 
                if ($cont) {
                    $DeleteWithoutRecycle = $true
                    Write-Host "Switching to fast delete without recycle" -ForegroundColor Yellow
                    }
                }
            $howDeleted = if ($HardDelete -or $DeleteWithoutRecycle) { 'Deleting ' } else { 'Recycling ' }
            $fcs = 0
            ForEach ($f in $RemoveFiles) {
                RecycleFile -file $f.FullName -purge $DeleteWithoutRecycle
                LogAction -File $f.Name -Action "--Remove-file"
                Show-Progress  -Activity $($howDeleted  + "extra files from collection") -Status "$fcs of $ttl - $($f.Name)" -PercentComplete $($fcs/$ttl*100)
                $fcs++
                }
            
            }
            if ($ttl) {
                AdvanceProgress  "Updating Lists of removed files .... Please wait" -BigStep
                $script:mysewingfiles = $mysewingfiles |where-object {$_.FullName -notin ($RemoveFiles.FullName)}
                Complete-Progress
            }


        }

    }


# Define a recursive function to traverse directories and remove empty directories
Function TailRecursion {
    param (
        [string]$Path,
        [int]$Depth = 0,
        [bool]$purge = $HardDelete
    )

    $IsFound = $false
    if ($depth -eq 0) { 
        AdvanceProgress "Looking at Directories" -BigStep
    } elseif ($depth -in (1,2)) { 
        AdvanceProgress "Looking at Directories"
    }

    # Recursively call the function for each child directory
    Get-ChildItem -Force -LiteralPath $Path -Directory | ForEach-Object {
        TailRecursion -Path $_.FullName -Depth ($Depth + 1) -Purge $purge| Out-Null
    }

    # Check if the current directory is empty
    $IsEmpty = -not (Get-ChildItem -Force -LiteralPath $Path)

    # If the directory is empty and it's not the top directory, remove it
    if ($IsEmpty -and $Depth -gt 0) {
        Write-Verbose "Removing empty folder: '$Path'"
        RecycleFile -file $Path -purge $purge
        $IsFound = $true

        $Script:DirCount++
        AdvanceProgress "Removing Directory" $Path
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
            AdvanceProgress "Preparing working directory" -BigStep
            Get-ChildItem -Path $("\\?\" + $NewFilesDir) -Recurse | Remove-Item -Force -Recurse
            write-verbose "CLEARED working copy filespace"
            Complete-Progress
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


# retrieve an image file (destination) from a web link (Source)
function FetchImageFile ([string]$URL, [string]$localfile) {
    if (-not (Test-Path $localfile)) {
        Write-Host "[+] Downloading file from '$URL'"
        try {
            (New-Object System.Net.WebClient).DownloadFile($URL, $localfile)
        } catch {
            Write-Host "`t[!] Failed to download '$URL'"
            Write-Host "`t[!] $_"
        }
    } else {
        Write-Verbose "[+] Using existing file."
    }
}

# Check GitHub to see if this has been updated
function Get-LatestGitHubTag {
    param (
        [string]$RepositoryOwner,
        [string]$RepositoryName
    )
    
    # Construct the GitHub API URL for releases
    $apiUrl = "https://api.github.com/repos/$RepositoryOwner/$RepositoryName/releases"
    AdvanceProgress "Checking for script updates..." -BigStep
    try {
        # Fetch the releases using Invoke-RestMethod
        $releases = Invoke-RestMethod -Uri $apiUrl
        # Filter the releases to get the latest one
        $latestRelease = ($releases | Sort-Object -Property created_at -Descending | Select-Object -First 1).tag_name
        # Extract and return the tag name
        $newfeatures = $releases | Sort-Object -Property created_at -Descending  | where-object { $_.tag_name -gt $ECCVERSION -and $_.prerelease -eq $false -and $_.target_commitish -eq "main" }
        $newfeatures = $newfeatures | foreach-object {$_.name + ": " + $_.body + "`n"}
    }
    catch {
        Write-Verbose "Error fetching releases from GitHub: $_"
        $latestRelease = ""    
        $newfeatures = ""    
    }
    Complete-Progress
    return $latestRelease, $newfeatures
}



Function OpenForUpload {
    
    if ($UsingUSBDrive) {
        return
    }
    Write-Host "-----------------------------------------------------------------------------------------" -ForegroundColor Green
        
    if ($DragUpload) {
        Write-Host "Opening File Explorer & MySewnet Web page" -ForegroundColor Green
        Write-Host " ** on MySewNet web page choose 'Upload' and Select all files in Explorer and " -ForegroundColor Green
        Write-Host "    drag/drop the files a maximum of 5 at a time into the upload box" -ForegroundColor Green
        
    } else {
        if ($CloudAPI) {
            return
        }
            
        if ((Get-WmiObject -class Win32_OperatingSystem).Caption -match "Windows 11") {
            $wtype = "w11"
            Write-Host "Opening File Explorer (using mysewnet add-in)" -ForegroundColor Green
            Write-Host " ***  Select all files *right-click* and choose 'Show more Options' -> choose 'MySewNet' -> 'Send'" -ForegroundColor Green
        } else {
            # Assume it is Windows 10 with add-in
            $wtype = "w10"
            Write-Host "Opening File Explorer (using mysewnet add-in)" -ForegroundColor Green
            Write-Host " ***  Select all files *right-click* and choose 'MySewNet' -> 'Send'" -ForegroundColor Green
            }
        
    }
    $firstfile = $(get-childitem -path $NewFilesDir -File -depth 1)
    if ($firstfile.count -gt 0) {
        $firstfile = $firstfile[0].FullName
        $explorercmd = "explorer  ""/select,$firstfile"""
        } 
    else { 
        Write-Host " There are NO Files to upload" -ForegroundColor Yellow
        $firstfile = $NewFilesDir + "\."
        $explorercmd = "explorer  ""$NewFilesDir"""
    }
    Write-Host "-----------------------------------------------------------------------------------------" -ForegroundColor Green
    
    if ($DragUpload) { 
        Start-Process $opencloudpage 
        }
    Invoke-expression  $explorercmd

    if (-not $DragUpload -and $ShowExample) { 
        $file = Join-Path -path $(Split-Path -path $PSCommandPath) -ChildPath "HowToSend-$wtype.gif"
        
        FetchImageFile  -URL "https://raw.githubusercontent.com/D-Jeffrey/Embroidery-File-Organize/main/docs/images/HowToSend-$wtype.gif" -localfile $file
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
        AdvanceProgress "Using previous MySewNet Logon" -BigStep
    } else {
        AdvanceProgress "Logginning onto MySewNet" -BigStep
        try {
            $code = Invoke-OAuth2AuthorizationEndpoint -uri $authorization_endpoint @idparams
        } catch {
            Complete-Progress
            return $false
        }
        AdvanceProgress "Authenicating onto MySewNet" -BigStep
        $tokens = Invoke-OAuth2TokenEndpoint -uri $token_endpoint  @code  
        AdvanceProgress "Completed Logon to MySewNet" -BigStep

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
        AdvanceProgress "Reading file list from MySewNet $(" ".padright($tries,[char]2160))" -BigStep
        try {
            $result = Invoke-RestMethod -Headers $authHeader -Uri $requestUri -Method GET -ContentType 'application/json'
        } catch {
            # slow down and see if trying again helps
            Start-Sleep -seconds 2
        }
        $tries++
        
    } while ($null -eq $result -and $tries -lt 3)
    
    if ($null -eq $result){
        write-host "MySewNet error, please try again later" -ForegroundColor Red
    } else {
        # Recurse the structure and add a Path attribute to the collection with the directory names
        CloudMetaAddPath "" $result
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
            # write-Debug "Meta _Path" ($metafolder.folders | where-object { $_.path -like $foldername})
            return $metafolder.folders | where-object { $_.path -like $foldername}
        } else {
            foreach($fid in $metafolder.folders) {
                $retdir = FindCloudidfromPath $foldername  $fid 
                if ($retdir) {
                    # write-debug "Ret $retdir"
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
    # BUGS TODO
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
        LogAction -File "Error Deleting folder id: '$id'" -Action ("StatusDescription:" + $_.Exception.Response.StatusCode.value__ + " "+ $_.Exception.Response.StatusDescription)
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
        LogAction -File "Error Deleting file id: '$id'" -Action ("StatusDescription:" + $_.Exception.Response.StatusCode.value__ + " "+ $_.Exception.Response.StatusDescription)
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
            LogAction -File "Error Moving file id: '$fileid'" -Action ("StatusDescription: "+ $_.Exception.Response.StatusCode.value__ + " " + $_.Exception.Response.StatusDescription)
		
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
            # write-debug "Folder " $fld.Name " ($name) exists as " $fld.id " for in " $fld.parentfolderid " ($inFoldID) "
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
        LogAction -File "Error Pushing the folder: '$name'" -Action ("StatusDescription: "+ $_.Exception.Response.StatusCode.value__ + " " + $_.Exception.Response.StatusDescription)
		$result = ""
		$myError = $_
	} 

	if ("" -eq $result) {
		$eDetails = ($myError.errorDetails.Message|convertfrom-json )
		write-warning ("Error Pushing the folder: '$name'  [ " + $eDetails.message + " ]" )
        $script:CloudStatusGood = $false
	} else {
		# write-debug "Folder '$name' saved as '$result'"
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
    if (test-Path -LiteralPath $filepath) {
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
		# write-warning ("StatusDescription: " + $_.Exception.Response.StatusCode.value__ + " " + $_.Exception.Response.StatusDescription)
		$result = ""
		$myError = $_
	} 

	if ("" -eq $result) {
		$eDetails = ($myError.errorDetails.Message|convertfrom-json )
		write-warning ("Error Pushing the file: '$name'  [ "  + $myError.Exception.Response.StatusDescription + $eDetails.message + " : " + " ]") 
        $script:CloudStatusGood = $false

	} else {
		# write-debug "File '$name' saved as '$result'"
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


function DoCleanCollection {
    
    write-host "Scanning for files to clean up in collection"

    # this creates a BUG for below as we are removing files that might be synced without removeing them from the list
    $filesToRemove = DuplicateFiles -Path $EmbroidDir 
    if ($filesToRemove) { 
        CheckAndRemove -RemoveFiles $filesToRemove -DeleteWithoutRecycle $HardDelete -why "you have multiple copies of the same file"
        }
    if (!$KeepAllTypes) {
        write-host "Scanning for folders to clean up in collection"
        $pe = $preferredSewType | ForEach-Object { ".$_" }
        $filesToRemove = DuplicateFileNames -Path $EmbroidDir -ExtensionsOrder $pe
        if ($filesToRemove) { 
            CheckAndRemove -RemoveFiles $filesToRemove -DeleteWithoutRecycle $HardDelete -why "you have multiple files of different embroidery types"
            }
        }
            
        # Look for lone directories
        $maxtries = 5
        $proceed = $false
        while ($maxtries -gt 0) {
            $Movelist = @()
            # 
            $cleanList =  Get-ChildItem -Path $EmbroidDir -Recurse -Directory 
                | Where-Object { ($_.GetFiles().count -eq 0 -and $_.GetDirectories().count -eq 1) -or ($_.GetFiles().count -eq 1 -and $_.GetDirectories().count -eq 0)}  

            $cleanList =  $cleanList | ForEach-Object -Begin { $stack = [System.Collections.Stack]::new() } -Process { $stack.Push($_) } -End 	{ while ($stack.Count -gt 0) { $stack.Pop() } } 
            if ($cleanList -and -not $proceed)  {
                $proceed = MyPause -Message "Clean up Directory structure.  (Remove empty and move lone files up)" -Choice $true -ChoiceDefault $false
            }
            if ($proceed) {
                $cleanList | ForEach-Object {
                $thisdir = $_
                if ($thisdir -and $thisdir.GetDirectories().count -eq 1) {
                    AdvanceProgress -Area $_ -stat "Moving Lone Directories"
                    $subdir = $thisdir.GetDirectories()
                    if ($subdir.GetFileSystemInfos().count -eq 0) {
                        if (-not $KeepEmptyDirectory) {
                            remove-item $subdir
                        }
                    } else {
                        $Movelist += [PSCustomObject]@{
                            From = $subdir.FullName
                            To = $($subdir.parent).FullName
                        }
                        $movethese = Get-ChildItem -LiteralPath $subdir.FullName -Recurse  -Filter "*"
                        foreach ($movethis in $movethese ) {
                            $DestPath = Split-Path -Path $(Split-Path -Path $movethis.FullName -Parent) -Parent
                            Move-Item -Path $movethis.FullName -Destination $DestPath
                        }
                        if (-not $KeepEmptyDirectory -and $subdir.GetFileSystemInfos().count -eq 0) {
                                remove-item $subdir -force
                        }
                    }
                } else {
                    #lone file
                    AdvanceProgress -Area $_ -stat "Moving Lone File"
                        
                    $movethese = Get-ChildItem -LiteralPath $thisdir.FullName -Recurse  -Filter "*"
                    foreach ($movethis in $movethese ) {
                        $DestPath = Split-Path -Path $(Split-Path -Path $movethis.FullName -Parent) -Parent
                        if (-not $(test-path -LiteralPath $DestPath)) {
                            Move-Item -Path $movethis.FullName -Destination $DestPath
                            $Movelist += [PSCustomObject]@{
                                From = $thisdir
                                To = $($thisdir.parent).FullName
                            }
                        }
                    }
                }
            }
            # Then we need to fix the mysewing files list
            $script:mysewingfiles = $mysewingfiles | ForEach-Object {
                if ($_.DirectoryName -in $MoveList.From ) {
                    foreach ($MoveLoc in ($MoveList)) {
                        if ($MoveLoc.From -like $_.DirectoryName) {
                            $_.DirectoryName = $moveLoc.To
                            $_.FullName =  join-path -path $moveLoc.To -ChildPath $_.Name
                            $_.RelPath = $_.DirectoryName.Substring($script:EmbroidDir.Length)
                            if (test-path -LiteralPath $_.FullName) {
                                $_.FileInfo = get-item -LiteralPath $_.FullName
                            }
                        }
                    }
                    
                }
                $_
            }
        } else {
            $maxtries = -1

        }
            $maxtries--
         
    }
}  
function FoldupDirPath {
    param (
        [string]$directoryPath,
        $minimumlen = 2
    )

    $directoryPathLower = $directoryPath.ToLower()

    do {
        $folding = $false
        foreach ($r in $foldupDirs) {
            if ($directoryPathLower.EndsWith("\$r")) {
                if ($(Split-Path $directoryPath -IsAbsolute) -and ($directoryPath.Length -lt $script:EmbroidDir.Length)) {
                    # Emergency panic exit
                    $directoryPath = $script:EmbroidDir
                    Write-Warning "Folding Directories has a problem, please change the 'foldupDir' List"
                    return $directoryPath
                }
                # Strip off the directory name and preserve the case of the directory and files
                if (($directoryPath.Length - $r.Length - 1) -gt $minimumlen) {
                    $directoryPath = $directoryPath.Substring(0, $directoryPath.Length - $r.Length - 1)
                    $directoryPathLower = $directoryPath.ToLower()
                    $folding = $true
                }
            }
        }
    } while ($folding)
    # TODO Check for nested duplicate names

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
        $targetdir = $script:EmbroidDir
    } else { 
        # Move anything that is not a Embrodery type file (alltypes)
        $dtype = "Instructions"
        $Excludes = ($allTypesStar + $TandCs )
        # if it is from a nested zip, then keep the zip because we don't yet expand that 
        if (!($isFromNestedRelative)) {
            $Excludes += @("*.zip")
        }
        $objs = Get-ChildItem $fromPath -file -Recurse -Exclude $Excludes 
        if ($whichfiles) {
            $objs = $objs | where-Object { $whichfiles -ilike $_.Name } 
        }
        $targetdir = $InstructDir
        }
	$oc = $objs.count
    if ($oc) {
        AdvanceProgress -Area "Copying $dtype" -Stat "Added ${Script:savecnt} files"
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
            if (Test-Path -LiteralPath $newpath -PathType Leaf) {
                # what do if there is already a file with the same name as the folder? Rename the confliciting folder
                $parent = Split-Path -Path $newpath -Parent
                $thisfolder = Split-Path -Path $newpath -Leaf
                $newpath = join-path -path $parent -childpath $($thisfolder.Replace('.','-'))
            }
            if (!(Test-Path -LiteralPath $newpath -PathType Container)) {
                New-Item -Path $newpath -ItemType Directory | Out-Null
                }
            $npath = Join-Path -Path $newpath -ChildPath $newfile 
            if (test-path -LiteralPath $npath) {  # See if the file already exists
                
                $newHash = get-filehash -Algorithm md5 -LiteralPath $_
                $orgHash = get-filehash -Algorithm md5 -LiteralPath $npath
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
            if (test-path -LiteralPath $_) {
                # BUG 
                # this can happen with the same file type is nested within other directories which then get folded to the same directory and are duplicate
                # We could rename the files, but then the code above need to find and match the same rename pattern
                if (test-path -LiteralPath $npath) { 
                    Write-verbose " WARNING File already exists $npath for $($_.FullName) - maybe renaming - overwriting?"
                    }
                else {
                    ChecktoClearNewFilesDirectory
                    if (!($UsingUSBDrive) -or $isEmbrodery) { 
                        if ($NoDirectory) {
                            Copy-Item -Path $_ -Destination $(Join-Path -Path $NewFilesDir  -ChildPath $newfile)
                        } else {
                            $newpath = join-path -path $NewFilesDir -childpath $newdir
                            if (!(test-path ($newpath))) {
                                New-Item -Path ($newpath) -ItemType Directory  | Out-Null
                                }
                            Copy-Item -Path $_ -Destination $(Join-Path -Path $newpath -ChildPath $newfile)
                            }
                        }
                    LogAction $newfile -Action "++Added-MoveFrom"
                    # BUG THIS IS THE LINE THAT ERRORS out wiht Directory Not Found
                    $movedirto = split-path -path $npath -parent
                    if (!(test-path -LiteralPath $movedirto -PathType Container)) {
                        New-Item -Path $movedirto -ItemType Directory  | Out-Null
                    }
                    Move-Item $_ -Destination $npath  -force # -ErrorAction SilentlyContinue
                    $newFileCount += 1
                    Write-Information "+++ Saving ${dtype}:'$_' to ${newdir} & ${newfiledir}"
                    }
                }
        $loopy++
        Show-Progress -Activity "$loopy - Copying $($_.Name)" -Status "Added ${Script:savecnt} files" -PercentComplete $($loopy*100/$oc)
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
    $thelist = (Get-ChildItem -Path $script:EmbroidDir  -Recurse -file -include $PrefSewTypeStar)| ForEach-Object { 
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
            RelPath = $_.DirectoryName.Substring($script:EmbroidDir.Length)
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
        PrepareCloud
        if ($CloudAuthAvailable)  {
            Write-host    $($paramswitch['CloudAPI']).padright($padder)   -ForegroundColor yellow
        } else {
            Write-host "Using API to update mySewNet Cloud requires PSAuthClient to be installed to use CloudAPI feature" -ForegroundColor Red
            }
        }
    }
}

function  CheckUSBDrive ($USBPath) {

    $driveReady = $true
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
           $driveReady = $False
        }
        else {
          Write-Host "The drive ${driveletter}: is not a removable device." -ForegroundColor Red
        }
      }
      else {
        Write-Host "The drive ${driveletter}: does not exist." -ForegroundColor Red
      }
      
    } elseif ($USBPath.tolower().contains("off")) {
        $script:USBDrive = ""
        $driveReady = $false
        $driveletter = ""
    } else {
        Write-Host "The drive letter is invalid ('$USBPath')." -ForegroundColor Red
    }
    if ($driveReady) {
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
    if ($nameIndex -eq "" -or $null -eq $nameIndex) {
        
        write-Error "** BLANK NAME - '$NameIndex', '$Name', '$directory', '$lastWriteTime' "
        start-sleep -Milliseconds 100
    }
    $isFoundnewfile = $true
    $Directory = FoldupDirPath -directoryPath $Directory
    if ($quickmysewfiles[$NameIndex]) {
        # BUG duplicate filename but different checksum??? 
        # TODO use  LastWriteTime to overcome
        # Check for duplicate FileHash
        # Find the instances then get the full name then get hash to compare
        # If there are multiple versions of the file (with different extensions then this Verbose will show several values)
        foreach ($q in $quickmysewfiles[$NameIndex]) {
            $d = $mysewingfiles[$q-1].LastWriteTime.date
            if ($LastWriteTime.date.AddDays(1) -ge $d -and $LastWriteTime.date.AddDays(-1) -le $d) {
                return ""
                }
            if ($Directory -eq $mysewingfiles[$q-1].DirectoryName) {
                return ""
                }
            $hash = $null


            if ($TmpPath.Exists -or $keepPath.Exists) {
                if ($TmpPath.Exists) {
                    $hash = (get-filehash -Algorithm md5 $TmpPath).Hash
                } else {
                    $hash = (get-filehash -Algorithm md5 $KeepPath).Hash
                }
                # TODO need to retest the file compare beyond name and date to Hash
                if (!($mysewingfiles[$q-1].Hash)) {
                    # TODO This will cause a duplicate later
                    $mysewingfiles[$q-1].Hash = $(get-filehash -Algorithm md5 $mysewingfiles[$q-1].FullName -ErrorAction SilentlyContinue).Hash 
                    }
                if ($mysewingfiles[$q-1].Hash -eq $hash) {
                    return ""
                    }
            } else {
                # we need to extract it to compare
                return "-"
            }
        }
    }

  
    $fullName = join-path -Path $Directory -ChildPath $Name
    $fileinfo = New-Object System.IO.FileInfo($fullname) 
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
        } 
    return $Name        
    
}


<#
.SYNOPSIS
    Searches for .PES files within a zip file, including nested zip files.

.DESCRIPTION
    This script recursively searches for files with specific extensions within a zip file. If the zip file contains nested zip files,
    those nested zip files are also checked for .PES files. The script avoids unnecessary extractions unless a zip
    within a zip is found.

.PARAMETER zipFilePath
    The path to the zip file to search.

.PARAMETER fileExtensionToSearch
    An array of file extensions to search for (e.g., @("PES", "VP3", "DST")).

.EXAMPLE
    Search-ZipFile -zipFilePath "C:\path\to\your\zipfile.zip" -fileExtensionToSearch @("PES", "VP3", "DST")

#>
function Search-ZipFile {
    param (
        [string]$zipFilePath,
        [string[]]$fileExtensionsToSearch
    )
    Write-Output "Check: $zipFilePath"
    # Create a temporary directory to extract files
    $tempDir = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), [System.Guid]::NewGuid().ToString())
    [System.IO.Directory]::CreateDirectory($tempDir) | Out-Null

    try {
        # Open the zip file
        $zipArchive = [System.IO.Compression.ZipFile]::OpenRead($zipFilePath)

        foreach ($entry in $zipArchive.Entries) {
            if ($entry.FullName -match "\.zip$") {
                # If the entry is a zip file, extract it to a temporary location and search it recursively
                $nestedZipPath = [System.IO.Path]::Combine($tempDir, $entry.FullName)
                $nestedzipParentPath = split-path -path $nestedZipPath -parent
                [System.IO.Directory]::CreateDirectory( $nestedzipParentPath) | Out-Null
                
                [System.IO.Compression.ZipFileExtensions]::ExtractToFile($entry, $nestedZipPath)
                if (test-path -path $nestedZipPath) {
                    Search-ZipFile -zipFilePath $nestedZipPath -fileExtensionsToSearch $fileExtensionsToSearch
                    }
            } else {
                foreach ($extension in $fileExtensionsToSearch) {
                    if ($entry.FullName -match "\.$extension$") {
                        # If the entry matches one of the file extensions, output the file path
                        Write-output "File found: $($entry.FullName)"
                    }
                }
            }
        }
    } finally {
        $zipArchive = $null
        # Clean up the temporary directory
        Remove-Item -Path $tempDir -Recurse -Force
    }
}



function ExpandAZip {
    param (
        $zippath,
        $RelativePath
    )
    $resultTmpDir = (Join-Path $tmpdir -childpath $RelativePath).trim("\")
    # Check for long path names inside the zip file
    
    Show-progress -Activity "Expanding Zip Archive.. Please wait " -Status (split-path -path $zippath -leaf)
    $bigzip = (get-item $zippath).Length -gt $use7zipsize
    if (($bigzip -or $havewarning) -and (Test-Path "C:\Program Files\7-Zip\7z.exe")) {
        Set-Alias sevenz "C:\Program Files\7-Zip\7z.exe"
        sevenz x $zippath -o"$resultTmpDir" -y
    } else {
        Expand-Archive -Path $zippath -DestinationPath $resultTmpDir -Force
    }
    Complete-Progress
    return $resultTmpDir
}

function ProcessZipContents {
    param (
        $zips,       # FullName
        $Base,
        $isNested = $false,
        $NestedName = ""       # Is it a zip in a zip
    )

    function BuildTmpAPath {
        param (
            $NewMadeDir,
            $FileFullName
        )
        if ($NewMadeDir) {
            $buildAPath = join-path -path $tmpdir -ChildPath $NewMadeDir 
        } else {
            $buildAPath = $tmpdir
            }
        $buildAPath = join-path -path $buildAPath -childpath $FileFullName
        if (test-path -LiteralPath $buildAPath) {
            $NewtempPath = get-item -LiteralPath  $buildAPath 
        } else {
            write-warning "Could not access: $($buildAPath.substring($tmpdir.Length))"
            # TODO this should not be counted
            $NewtempPath = "NOACCESS"
        }
            
        return $NewtempPath
    }

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
    $isZipExtracted = $false
    $isnew = $false
    $numnew = 0
    foreach ($thistype in $preferredSewType) {
        $filesInThisList = @()
        $ts = "*."+ $thistype
        if ($zipfilelist.Entries.Name -ilike $ts) {
            $isnew = $false
            # like our extension, but does not start with . 
            $SpecificExtensionFiles = $zipfilelist.Entries | where-object {$_.Name -like $ts -AND $_.Name.substring(0,1) -ne "."} 
            AdvanceProgress  "Checking Zips - Looking at $($_.Name) - looking at '$($ts.substring(2).ToUpper())' type"  -stat "Added $Script:savecnt files"
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
                if ($isZipExtracted) {
                    $tempPath = BuildTmpAPath -NewMadeDir $madeDir -FileFullName $fileInZip.FullName
                } else {
                    $tempPath = $null
                }
                # Can not add a file we don't have access too
                if ("NOACCESS" -eq $tempPath) {
                    $isnewfile = ""
                } else {
                    $isnewfile = AddToSewList -Name $fileInZip.Name -Directory $dirn -NameIndex $n -LastWriteTime $fileInZip.LastWriteTime -RelativePath $relativepath -TmpPath $tempPath
                }
                if ("" -ne $isnewfile -and -not $isZipExtracted) { 
                    # TODO We should not expand the ZIP, if the file type found is a duplicate of another type... hard to determine
                    $isZipExtracted = $true
                    ExpandAZip -zippath $zips -RelativePath $madeDir |out-null
                    #try again
                    $tempPath = BuildTmpAPath -NewMadeDir $madeDir -FileFullName $fileInZip.FullName
                    if ("NOACCESS" -eq $tempPath) {
                        $isnewfile = ""
                    } elseif ($isnewfile -eq "-") {
                        # try again after expand
                        $isnewfile = AddToSewList -Name $fileInZip.Name -Directory $dirn -NameIndex $n -LastWriteTime $fileInZip.LastWriteTime -RelativePath $relativepath -TmpPath $tempPath
                    }
                }
                        
                if ("" -ne $isnewfile ) {
                    Write-verbose "New file '${filenameInZip}'"
                    $isnew  = $true
                    $filesInThisList += $isnewfile
                } else {
                    if ($VerbosePreference -eq  "Continue") {
                        $fileInstance = $mysewingfiles | where-object {$_.NameIndexed -eq $filenameInZip}
                        $fiName = $fileInstance.FullName
                        Write-verbose "Duplicate zfile '${filenameInZip}' to ${fiName}"
                        }
                    }
                }
            
                # we found a new file in the Zip
                if ($isnew) { 
                    $numnew += $(MoveFromDir -fromPath $tmpdir -isEmbrodery $true -files $filesInThisList -whichfiles $ts)
                }               
            }
        }
        if ($isZipExtracted) {
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
            AdvanceProgress  "Checking Zips"  "Added $Script:savecnt files"
        } else {
            Write-host $("- Found: '$zf'").padright(65) " nothing new" 
            }
        # we extracted the Zip already and now let's check for instructions but only do this at the very end of nested files
        if ($isZipExtracted -and (!($isNested))) { 
            $numnew += $(MoveFromDir -fromPath $tmpdir -isEmbrodery $false)
            Get-ChildItem -Path $tmpdir -Recurse | Remove-Item -force -Recurse
            }
   
        
    }   
#======================================================================================
#
# Building out all the directory structures and File lists
#
#======================================================================================
function BuildTypeLists() {
    $script:PrefSewTypeStar = $script:preferredSewType | ForEach-Object { "*.$_" }
    if ($script:PrefSewTypeStar.count -eq 0 -or $null -eq $script:PrefSewTypeStar) {
        write-error "Miss configuration of 'preferredSewType', can not continue"
        throw [System.Exception] "Halting the script!"
        return
    }
    $script:SewTypeMatch = $script:preferredSewType -join '|'
    if ($script:foldupDir.count -eq 0) {
        $script:foldupDir = $defaultfoldupDir
        }
    if ($script:TandCs.count -eq 0 ) {
        $script:TandCs = $defaultTandCs
    }
    if ($script:goodInstructionTypes.count -eq 0) {
        $script:goodInstructionTypes = $defaultgoodInstructionTypes
    }
    $script:foldupDirs = $foldupDir + $preferredSewType | ForEach-Object { $_.ToLower() }
    if ($foldupDirs.count -eq 0 -or $null -eq $foldupDir) {
        write-error "Miss configuration of 'foldupDir', can not continue"
        throw [System.Exception] "Halting the script!"
    }
    $script:allTypesStar = $alltypes | ForEach-Object {"*.$_"}
}



<#
.SYNOPSIS
Do the configuration for Setup Processes

.DESCRIPTION
Create a ICON on the desktop.  Add link to powershell script.  Automatically updated to PSH if that option is available during Setup mode.
Use Folder Dialog box to select root folder.  Prompt for other settings options.  Install PSAuthenicate Module if Cloud is used.

.NOTES
Someday this should be done in a GUI
#>
Function DoSetup() {


    write-host "   ".padright(86) -BackgroundColor Yellow -ForegroundColor Black
    # Load the System.Windows.Forms assembly
    Add-Type -AssemblyName System.Windows.Forms
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
    
    # Instantiate a FolderBrowserDialog object
    $DirectoryBrowser = New-Object System.Windows.Forms.FolderBrowserDialog -Property @{ 
        SelectedPath = $script:EmbroidDir
        Description = "Select the Directory for Embroidery Files.  This will be used as a reference location for you to look at the current collection of files"
        ShowNewFolderButton = $true
        # https://learn.microsoft.com/en-us/dotnet/api/system.environment.specialfolder
        rootFolder = 40 # CommonDocuments
        }
    # Instantiate a FolderBrowserDialog object
    $USBBrowser = New-Object System.Windows.Forms.FolderBrowserDialog -Property @{ 
        SelectedPath = $script:USBDrive
        Description = "Select the USB device for Embroidery Files to be copied to"
        ShowNewFolderButton = $false
        # https://learn.microsoft.com/en-us/dotnet/api/system.environment.specialfolder
        rootFolder = 17 # MyComputer 
        }
    do {
        write-host "In order to Setup you will be answer questions to configure this script".padright(70)   -BackgroundColor Blue -ForegroundColor White
        write-host $($paramstring['EmbroidDir']) "?" -NoNewline
        # Show the dialog box and store the selected folder path
        if ($DirectoryBrowser.ShowDialog() -eq "OK") {
            $script:EmbroidDir = $DirectoryBrowser.SelectedPath
        } else {
            $continuesetup = $false
            write-host " "
            write-host "Stopping setup - No settings changed"  -BackgroundColor Blue -ForegroundColor White
            Write-Host ( "End") -ForegroundColor Green
            MyPause 'Press any key to Close' | out-null
            return $true
        }
        write-host $script:EmbroidDir
        write-host "How do you want to transfer your files to your machine (USB, Mysewnet or neither)"
        if (myPause -Message "Are you using a USB Drive?" -Choice $true -ChoiceDefault ($script:USBDrive -ne "")) {
            
            do {
                if ($USBBrowser.ShowDialog() -eq "OK") {
                    $script:USBDrive = $USBBrowser.SelectedPath
                } else {
                    $script:USBDrive = ""
                }
                
                if ($script:USBDrive -eq "") {
                    $notvalid = myPause -Message "Do you still want to use a USB Drive?" -Choice $true
                } else {
                    $udrive = CheckUSBDrive -USBPath $script:USBDrive
                    $notvalid = ($udrive -eq "")
                    }
            } while ($notvalid)
            $CloudAPI = $false
            if ($script:USBDrive) {
                write-host "USB Drive Selected - $script:USBDrive"
            }
        } else {
            $script:USBDrive = ""
        }
        if ($script:USBDrive -eq "") {
            $script:UsingUSBDrive = $false
            if (myPause -Message "Are you using MySewnet Cloud" -Choice $true -ChoiceDefault $script:CloudAPI) {
                $script:CloudAPI = $true
                $script:USBDrive = ""
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
                $script:CloudAPI = $false
                $script:USBDrive = ""
                $script:UsingUSBDrive = $false
            }
        }
        $dd = Read-Host "How many days back do you want to always look when checking the Download folder? (currently $script:DownloadDaysOld)"
        if ($dd -gt 0) {
            $script:DownloadDaysOld = $dd
        }         
        $script:KeepAllTypes = myPause $parambool['KeepAllTypes'] -Choice $true -ChoiceDefault $script:KeepAllTypes
        $script:KeepEmptyDirectory = myPause $parambool['KeepEmptyDirectory'] -Choice $true -ChoiceDefault $script:KeepEmptyDirectory
        if ($script:CloudAPI) {
            $script:DragUpload = myPause $parambool['DragUpload'] -Choice $true -ChoiceDefault $script:DragUpload
            $script:ShowExample =myPause $parambool['ShowExample'] -Choice $true -ChoiceDefault $script:ShowExample
            }
        
        $val = $alltypes  -join ', '
        Write-host "All the different Embroidery file types: `n$val" 
        
        
        Write-host "What are the preferred types of files for your machine in order of preference"
        write-host "Current list is: " $($script:preferredSewType -join ", " )  -ForegroundColor Yellow
        
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
                    $script:preferredSewType = $ptype
                }
            }
        } while ($problemext)
        write-host "  All Settings ".padright(70) -BackgroundColor Blue -ForegroundColor White
        ShowPreferences -showall $true
        $savep = mypause -Message "Do you want to save these settings?" -Choice $true
        $continuesetup = -not $savep
    } while ($continuesetup)
    $script:FirstRun = $script:FirstRun -or $(Test-path -path $script:ConfigFile)
    SaveAllParams
    $script:instructDir = $script:EmbroidDir
    if (-not ((Test-Path($script:EmbroidDir)) -and (Test-Path($instructDir)))) {
        write-host "  Creating Directory for Embroidery files collection ".padright(70) -BackgroundColor Yellow -ForegroundColor Black
        
        if (!(Test-Path($script:EmbroidDir))) {
            New-Item -ItemType Directory -Path $script:EmbroidDir | out-null
            write-host "Creating Directory '$script:EmbroidDir' for Embroidery files" -BackgroundColor Green -ForegroundColor Black
            LogAction -File $script:EmbroidDir -Action "Created-Directory"
        }
        if (!(Test-Path($script:InstructDir))) {
            New-Item -ItemType Directory -Path $script:instructDir | out-null
            write-host "Creating Directory '$script:instructDir' for Instructions files" -BackgroundColor Green -ForegroundColor Black
            LogAction -File $script:instructDir -Action "Created-Directory"
        }
    }
    if ($script:CloudAPI -and ((get-module -name PSAuthClient).count -lt 1)) {
        write-host "You will need to install the Powershell module in order to access MySewnet use the following command:" -ForegroundColor Yellow
        write-host "      install-module -name PSAuthClient" -ForegroundColor Blue
        write-host "with completing that you will not be able to use the CloudAPI feature" -ForegroundColor Yellow

    }
    write-host "All Setup " -BackgroundColor Yellow -ForegroundColor Black
    if (-not $script:FirstRun) {
        $script:FirstRun = mypause -Message "Would you like to run trigger the script to collect *ALL* your Embroidery files that you have every downloaded? " -Choice $true -ChoiceDefault $script:FirstRun
    }
    
    Return $false
}
function SetNewFilesDir ()
{
    if ("" -ne $USBDrive) {
        do {
            $driveletter = CheckUSBDrive $USBDrive
            $needadrive = "" -eq $driveletter
            if ($needadrive) {
                $needadrive = MyPause "USB Drive $usbDrive is not ready, do you want to use your USB Stick (insert it now)" -choice $true
                if (-not $needadrive) {
                    $driveletter = ""
                    $script:USBDrive = ""

                }
            }
        } while ($needadrive)
        if ("" -ne $driveletter) {
            
            $script:NewFilesDir = $USBDrive
            # Don't wipe someone's USB drive
            $Script:clearNewFiles = $False
            $Script:DragUpload = $False
            $Script:ShowExample = $false
            $Script:UsingUSBDrive = $True
            
        }
    } 
    # if USBDrive is not ready or was not selected then use the internal temporary space
    if ("" -eq $USBDrive) {
        $Script:UsingUSBDrive = $false
        $Script:clearNewFiles = $true
        $Script:NewFilesDir = ${env:temp} + "\cleansew.new"
        if (-not (Test-Path -Path $NewFilesDir )) { New-Item -ItemType Directory -Path ($NewFilesDir )}
    }
}

#
############################### MAIN continues ##############
#

PrepareCloud

$doit = !$Testing

# This is for development testing and debugging

if ($env:COMPUTERNAME -eq "DESKTOP-R3PSDBU_") { # -and $Testing) {
    $docsdir = "d:\Users\kjeff\"
    $downloaddir = "d:\Users\kjeff\downloads"
    $doit = $true
    }

BuildTypeLists

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
    ($latestTag, $newdescription) = Get-LatestGitHubTag -RepositoryOwner $GitOwner -RepositoryName $GitName
    $script:LastCheckedGithub = get-date -format "g"
    if ($latestTag) {
        Write-Verbose "Latest tag in D-Jeffrey/Embroidery-File-Organize $latestTag"
        if ($latestTag -gt $ECCVERSION) {
            Write-host "  *** Newer version ($latestTag) of this script is available" -ForegroundColor Green
            Write-host $newdescription
            $upgrademe = MyPause -Message "Do you want to upgrade" -Choice $true -Timeout 300 -ChoiceDefault $false
            if ($upgrademe) {
                $upgradescript = Join-Path -Path $PSScriptRoot -ChildPath "install.ps1"
                if (test-path $upgradescript) {
                    powershell -ExecutionPolicy bypass -file $upgradescript -"$ECCVERSION"
                    return
                } else {
                    Write-Warning "Automatic upgrade script '$upgradescript' can not be found to run, continuing without upgrade"
                }
            }
        }
    }
    else {
        Write-Verbose "Failed to retrieve the latest tag for  D-Jeffrey/Embroidery-File-Organize from github."
    }
    }

########################

if ($setup) {
    if (DoSetup){
        break
    }
}

if ($FirstRun) {
    Write-Host " ".padright(15) $("** Checking ALL Zip files **".padright(70)) -ForegroundColor white -BackgroundColor blue
}

# Clean out the old tmp working space
AdvanceProgress  "Cleaning up temporary work space" -BigStep
Get-ChildItem -Path  ($tmpdir ) -Recurse | Remove-Item -force -Recurse
if (-not (Test-Path -Path $tmpdir )) { New-Item -ItemType Directory -Path ($tmpdir )}
Complete-Progress

$failed = $false

# TODO change for USBDrive
# We will put all the new files in here for now
 $UsingUSBDrive = $False

$flatdir = -not $noDirectory

doWinForm | out-null
Complete-Progress  
if ($script:SetExiting) {
        write-host "Stopping" -ForegroundColor green
        break
}

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

AdvanceProgress  "Starting" 

if ($CloudAPI -and $CloudAuthAvailable) {
    if (-not (LoginSewnetCloud)) {
        write-host "Login failed or canceled - Stopping" -ForegroundColor Red
        Complete-Progress
        return
    }
}

SaveAllParams
if (( $EmbroidDir.tolower().contains("\onedrive") )) {
    Write-Host "The Embroidery files directory '$EmbroidDir' is within OneDrive ---- Warning" -ForegroundColor Yellow
    
    }
$InstructDir = $EmbroidDir
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
if ($sync) {
    write-Host "Sync Mode"  -ForegroundColor Green
}


$beginTimer = Get-Date

Add-Type -assembly "system.io.compression.filesystem"
Function CalculateSize()
{
    $byExt = @{}
    AdvanceProgress  "Calculating size" -BigStep
    $script:validsize = $false
    $eSize = 0
    $EmbSize = 0
    Get-ChildItem -Path $EmbroidDir  -Recurse -file  | ForEach-Object { 
        $eSize +=  $_.Length
        $byExt[$_.Extension.replace(".","")] +=  $_.Length
    }
    $EmbSize = $(foreach ($_ in $preferredSewType) { if ($_ -in $byExt.Keys) { $byExt[$_]} })  |  Measure-Object -Sum | Select-Object -ExpandProperty Sum

    $script:validsize = $true
    return ($esize, $EmbSize)
}

($librarySizeBefore, $libraryEmbSizeBefore) = CalculateSize
# "Starting with All files: $(niceSize $librarySizeBefore) - Embroidery files:  $(niceSize $libraryEmbSizeBefore)"

AdvanceProgress  "Loading file list" -BigStep
# Get a list of all the existing files in mySewnet
$mysewingfiles = LoadSewfiles
$quickmysewfiles = BuildHashofMySewingFiles

if ($FirstRun) {
    $DownloadDaysOld = 20*365
    $zipdepth = 50
} else {
    $zipdepth = 0
}
# $mysewingfiles | ft


$havewarning = $false
$afterdate = (Get-Date).AddDays(- $DownloadDaysOld )
Get-ChildItem -Path $downloaddir  -file -filter "*.zip" -depth $zipdepth | Where-Object { (($_.CreationTime -gt $afterdate) -OR ($_.LastWriteTime -gt $afterdate)) -and ($_.gettype().Name -eq 'FileInfo')} |
  
    ForEach-Object {
        AdvanceProgress  "Checking Zips - Looking at $($_.Name)"  -stat "Added $Script:savecnt files"
        Write-Verbose "Checking ZIP '$($_.FullName)'" 
        ProcessZipContents -zips $_.FullName -Base $_.BaseName
    }

# Look for Files which are not part of a ZIP file, just the selected file types that we are looking for that is in the download directory
$DownloadDaysOld = 365*20  # 20 years of downloads (when you download files, it keeps the old data)
$ppp = 0
foreach ($thistype in $preferredSewType) {

    write-Information "Working on File type: *.$ts"
    Get-ChildItem -Path $downloaddir  -file  -Depth ($zipdepth + 1) -Recurse|
             Where-Object { $_.Extension -like ".$thistype" -and $_.CreationTime -gt (Get-Date).AddDays(- $DownloadDaysOld ) } |
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
            if ($NoDirectory) {
                $l = ""
            } else {
                $l = (split-path -Path $fullname -Parent).Substring($downloaddir.Length).trim('\')
            }
            $d = (join-path -path $EmbroidDir -childpath $l).Trim('\')
            $isnew = AddToSewList -NameIndex $findfile -Name $f -Directory $d -LastWriteTime $thisfile.LastWriteTime -keepPath $thisfile -RelativePath $l 
    
            if ("" -eq $isnew) {
                Write-verbose "Duplicate file '$($findfile)'"
            } else {
                if (test-path -LiteralPath $fullname) { 
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
                    if (test-path -LiteralPath $fd) {
                        Copy-Item -Path $fd -Destination $InstructDir -ErrorAction SilentlyContinue 
                        Write-Information "+++ Copied instructions from Downloads :'$($_.Name)' to $InstructDir"
                        LogAction -File $($_.Name) -Action "++Added-from-Download" -isInstrution $true
                    }
                }
            }
        $ppp++
        AdvanceProgress  "Copying from Downloads" 
        }    
    }


    # clean up the zip file mess
    foreach ($tz in $tempziplist) {
        remove-item -Path $tz 
    }
    Complete-Progress "Copying from Downloads" 
    
# $mysewingfiles | ft

# TODO check for BUGS ?? 

if ($CleanCollection) {
    DoCleanCollection
}
  


#  Clear out empty Directories
if (-not $KeepEmptyDirectory) {
    show-progress -Activity "Clearing Empty Directories"
    $tailr = 0    # Loop thru 8 times to remove empty directories, then go back and check to see if you made any more emty
    while ($tailr -le 8 -and (tailRecursion $EmbroidDir) ) {
         $tailr++
         show-progress -Activity "Clearing Empty Directories" -percentcomplete ($tailr * 100 / 8) -status "Round $tailr"
    }
    complete-progress  "Clearing Empty Directories"
}

# Push to the Cloud & optionally Sync to the Cloud
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
        $cm++
        $n = $($_.FileInfo.Name).PadLeft(22)
        Show-Progress -Activity "Matching files to Cloud: $n" -PercentComplete ($cm * 100 / $mysewingfiles.count) -Status "$cm of $($mysewingfiles.count)"
        $thisfile.CloudRef = GetFileIDfromCloud $_.N
        if ($thisfile.CloudRef)   {
            $cf = $cf +1
            }
        }
    Complete-Progress -Activity "Matched files to Cloud "
    write-host "Found $cf cloud file which match local collection of $($MySewingfiles.count) files"
    # $MySewingfiles | Out-GridView -Title "MySewingfiles List"

    if ($sync) {
        $pool = $webcollection
        $cloudfileremove = @()
        do {
            $cloudfileremove  += $pool | Select-Object -ExpandProperty Folders | select-Object -ExpandProperty Files| Where-Object {$_.Name -notin ($mysewingfiles.N)} 
            $pool = $pool | Select-Object -ExpandProperty Folders 
            } while (($pool.files.count + $pool.folders.count) -gt 0)
        if ($cloudfileremove) {
            $i = 0
            $fc = $cloudfileremove.count
            $cloudfileremove | ForEach-Object  {
                $i++
                Show-Progress -PercentComplete $($i*100/$fc) "Removing files from cloud :" -Status $("$i of $fc " + $_.Name)
                LogAction -File $_.Name -Action "--Deleted-Sync"
                DeleteCloudFile -id $_.id | Out-Null
            }
            Complete-Progress "Removing files from cloud"
        }
        $filestopush = ($mysewingfiles | Where-Object { ($_.Push -and $_.Push.contains('\'))  -or ($_.CloudRef -eq $null)}).count

        }
     else {
        $filestopush = ($mysewingfiles | Where-Object { ($_.Push -and $_.Push.contains('\')) }).count
        }
    $spacereq = 0
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
                    AdvanceProgress "Delete Cloud Folder $($_.Name)"
                    DeleteCloudFolder -id $_.id | Out-null
                    }
            }
         #   } while ($emtpyfoldersList.count -gt 0)
        }
    if ($script:CloudStatusGood) {
        write-host "Beginning push to MySewnet: $filestopush files" -ForegroundColor Green
        $i = 0
        $fc = $MySewingfiles.count
        if ($filestopush -or $sync) {
            $MySewingfiles | ForEach-Object {
                $thisfile = $_
                $i++
                $complete = $i*100/$fc
                if ($thisfile.push -and $thisfile.push.contains("\") ) {
                    Show-Progress -PercentComplete $complete "Pushing files to the Cloud : $($thisfile.N)" -Status "$i of $fc"
                    PushCloudFileToDirectory -filepath ($thisfile.FullName) -folderpath $thisfile.push | Out-Null
                    LogAction -File ($thisfile.push +  $thisfile.N) -Action "^^Cloud-Added"
                    $thisfile.push = ""
                    PushCloudFileToDirectory -filepath ($thisfile.FullName) -folderpath $thisfile.push | Out-Null
                    LogAction -File ($thisfile.push +  $thisfile.N) -Action "^^Cloud-Added"
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
                            LogAction -File ($thisfile.RelPath + "\" + $thisfile.N) -Action ">>Cloud-Move"
                        }

                    } else {
                        Show-Progress -PercentComplete $complete "Syncing files to the Cloud : $($thisfile.N)" -Status "$i of $fc"
                        $spl = $thisfile.DirectoryName.substring($EmbroidDir.Length)
                        PushCloudFileToDirectory -filepath ($thisfile.FullName) -folderpath $spl | Out-null
                        LogAction -File ($spl + "\" + $thisfile.N) -Action "^^Cloud-Added-Sync"
                        }
                    }
                    if (!($script:CloudStatusGood)) {
                        write-error "We have a problem with the Cloud - STOPPING"
                        break
                    }
                    
                }
                Complete-Progress "Pushed files to Cloud "
            }
        }
    }
#  Sync to the USB Drive if required
elseif ($UsingUSBDrive) {
    if ($sync) {
        #TODO this need to include the -not KeepAll Filter
        $cm = 0
        $MySewingfiles | ForEach-Object {
            $cm++
            $n = $($_.FileInfo.Name).PadLeft(22)
            Show-Progress -Activity "Matching files to USB: $n"   -PercentComplete ($cm * 100 / $mysewingfiles.count) -Status "$cm of $($mysewingfiles.count)"
            $usbfile = join-path -path $USBDrive -ChildPath $_.relPath | join-path -ChildPath $_.fileinfo.Name
            if (test-path ($usbfile)) {
                $_.CloudRef = $usbfile
                
            } else {
                $spacereq = $spacereq + $_.FileInfo.Length
            }
        }
        Complete-Progress
        
        $spaceAvailable = Get-PSDrive -Name $($USBDrive.Substring(0,1))
        $spaceAvailable = $spaceAvailable.Free
        if ($spacereq -gt ($spaceAvailable+1000)) {
            write-host "Error: not enough free space available on $USBDrive requires $(NiceSize $spacereq) (only $(niceSize $spaceAvailable) free)" -ForegroundColor Red
            write-host "Stopping" -ForegroundColor Red
            break
        }
        $cm = 0
        $fileToSync = $mysewingfiles | where-object Cloudref -eq $null
        if ($fileToSync) {
            $fileToSync | foreach-object {
                $n = $($_.FileInfo.Name).PadLeft(25)
                Show-Progress -Activity "Copying missing files: $n"  -PercentComplete ($cm * 100 / $fileToSync.count) -Status "$cm of $($fileToSync.count)"
                $usbfile = New-Object System.IO.FileInfo($(join-path -path $USBDrive -ChildPath $_.relPath | join-path -ChildPath $_.fileinfo.Name))
                $usbfile.Directory.Create()
                copy-item -path $_.fileinfo -Destination $usbfile -Force
                $_.CloudRef = $usbfile
                $Script:addsizecnt += $_.FileInfo.Length
                LogAction -File $usbfile.fileinfo.fullName -Action "++Sync-USB"
                $cm++
                }   
            $script:savecnt = $script:savecnt + $fileToSync.count
                
            }
        # TODO
        
        $filesToRemoveUSB = get-childitem -path $($USBDrive + "\") -Recurse -File |where-object {$_.Extension.replace(".","") -in ($alltypes)} 
        $filesToRemoveUSB = $filesToRemoveUSB.FullName | Sort-Object 
        $removeFiles = @()
        $unum = 0
        $usblist = $mysewingfiles.CloudRef | sort-object  
        if ($usblist) {
            foreach ($f in $filesToRemoveUSB) {
                if ($f -eq $usblist[$unum]) {
                    $unum++
                } else {
                    $removeFiles += $f
                    while ($f -gt $usblist[$unum]) {
                        $unum++
                    }
                } 
                show-progress -Activity "Checking files to remove" -PercentComplete $($unum*100/$filesToRemoveUSB.count) -status $f
            }
        }
        $fc = $removeFiles.count
        if ($removeFiles) {
            $cm= 0
            $removeFiles | foreach-object {
                $cm++
                $n = $($_).PadLeft(26)
                $n = $n.substring($n.Length-25)
                Show-Progress -Activity "Removing extra files: $n"  -PercentComplete $($cm * 100 / $fc) -Status "$cm of $fc"
                
                try {
                    remove-item -path $_ -force -ErrorAction SilentlyContinue | Out-Null
                    LogAction -File $_ -Action "--Remove-USB"
                } catch {
                    write-output " Problem deleting - $_ -" 
                }
                # $Script:removesizecnt += $_.FileInfo.Length
            }
        }
        if (($script:savecnt -or $removeFiles) -and -not $KeepEmptyDirectory) {
            show-progress -Activity "Clearing Empty Directories from USB"
            $tailr = 0    # Loop thru 8 times to remove empty directories, then go back and check to see if you made any more emty
            while ($tailr -lt 8 -and (tailRecursion $USBDrive -Purge $true) ) {
                    $tailr++
                    show-progress -Activity "Clearing Empty Directories" -percentcomplete ($tailr * 100 / 8) -status "Round $tailr"
            }
            complete-progress  "Clearing Empty Directories from USB"
        }
    }
}

write-host "Calculating size"
 
if ($Script:savecnt -gt 0) {
    if (-not $CloudAPI) {
        OpenForUpload
    }
    ($librarySizeAfter, $libraryEmbSizeAfter) = CalculateSize
    
} else {
    $librarySizeAfter = $librarySizeBefore
    $libraryEmbSizeAfter = $libraryEmbSizeBefore
}

Complete-progress
if ($Script:dircnt -gt 0 -or $filesToRemove.length -gt 0) {
    $filecnt = $filesToRemove.length
    # bug in counting this
    Write-Host "Cleaned up - Directories removed: '$Script:dircnt    Files removed : '$filecnt' ($sizecnt)." -ForegroundColor Green
    }
if ($Script:savecnt -gt 0) {
    if ($Sync) {$what = "Synced" } else { $what = "Added"}
    write-host "+++ $what files to Embriodery Collection: '$($Script:savecnt)' files $(niceSize $Script:addsizecnt) " -ForegroundColor Green
    write-host "File size before   Total: $(niceSize $librarySizeBefore) - Embroidery files: $(niceSize $libraryEmbSizeBefore)"
    write-host "          after    Total: $(niceSize $librarySizeAfter) - Embroidery files: $(niceSize $libraryEmbSizeAfter)"
    if ($UsingUSBDrive) {
        $f = $(join-path -Path $USBDrive -ChildPath $markdrive)
        Set-Content -Value "Created/updated with Embroidery Collection Cleanup version: $ECCVERSION" -Path $f -Force -ErrorAction SilentlyContinue
        Set-ItemProperty -Path $F -Name Attributes -Value ([System.IO.FileAttributes]::Hidden) -ErrorAction SilentlyContinue
    }
    }
else {
    # TODO calc instruction size
    write-host "*** Embroidery collection size is: $(niceSize $libraryEmbSizeBefore) ****" -ForegroundColor Green 
}
if ($CloudAPI -and $CloudAuthAvailable) {
    if ($script:CloudStatusGood) {
        $updatedMeta = ReadCloudMeta
        write-Host "Cloud Storage currently is: " -ForegroundColor Green
        write-Host "     Used of Total: " (NiceSize $updatedMeta.storage.usedSize) "of" (NiceSize $updatedMeta.storage.totalSize)
        write-Host "   Space remaining: " (NiceSize $updatedMeta.storage.availableSize)
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
Write-Host "Time to complete the job: $minutes minutes and $seconds seconds" -ForegroundColor Blue

Complete-Progress
if ($UsingUSBDrive -and $USBEject) {
    Show-Progress -Activity "Ejecting USB Drive" -PercentComplete 50
    # Eject the USB drive
    if (-not $(EjectAndCheckUSB)) {
        start-sleep 2
        if (-not $(EjectAndCheckUSB)) {
            Write-host "Warning: Some files may not be completely written. Please try ejecting the USB again." -ForegroundColor Red
        }
    }
    Complete-Progress
    }
Write-Host "End" -ForegroundColor Green
if ($null -eq $psEditor.Workspace) {
MyPause 'Press any key to Close' -Timeout 120 |  out-null
}
# End of script