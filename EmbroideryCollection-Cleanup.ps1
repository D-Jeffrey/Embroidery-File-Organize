#
# EmbroideryCollection-Cleanup.ps1
#
# Deal with the many different types of embroidery files, put the right format types in mySewingnet Cloud
# We are looking to keep the ??? files types only from the zip files.
#
# Orginal Author: Darren Jeffrey Dec 2021
#                           Last Feb 2023
#

param
(
  [Parameter(Mandatory = $false)]
  [int32]$DownloadDaysOld = 7,                        # How many days old show it scan for Zip files in Download
  [int32]$aSetSizeis = 10,
  [Switch]$KeepAllTypes,                              # Keep all the different types of a file (duplicate name but different extensions)
  [Switch]$CleanCollection,                           # Cleanup the Collection Folder to only EmbrodRootDirtop files
  [Switch]$CleanCollectionIgnoreDir,                  # Cleanup the Collection Folder to only EmbrodRootDirtop files and look for duplicates ignoring directories meanings
  [string]$EmbrodRootDirtop = "Embroidery",           # You may want to change this directory name inside of the 'Collection ' Directory
  [string]$instructions = "Embroidery Instructions",  # This is a Directory name inside of "Documents" where instructions are saved
  [Switch]$includeEmbFilesWithInstruction,            # Put the Instruction along with the Embrodery files (NOT recommended as the PDFs tend to take up a lot of space)
  [Switch]$HardDelete,                                # Delete the files rather than sending to recycle bin
  [switch]$KeepEmptyDirectory,                        # If you don't want this to remove extra empty directories from Collection folders'
  [Switch]$Testing                                   # Run it and see what happens
  )

# ******** CONFIGURATION 
$preferredSewType = ('vp4', 'vp3',  'vip', 'pcs', 'dst')
$alltypes =('hus','dst','exp','jef','pes','vip','vp3','xxx','sew',
    'vp4','pcs','vf3','csd','zsk','emd','ese','phc','art','ofm','pxf','svg','dxf')
$goodInstructionTypes = ('pdf','doc', 'docx', 'txt','rtf'. 'mp4', 'ppt', 'pptx', 'gif', 'jpg', 'png', 'bmp','mov', 'wmv', 'avi','mpg', 'm4v' )
$TandCs = @('TERMS-OF-USAGE.*', 'planetappliquetermsandconditions.*')
$foldupDir = @('images','sewing helps','Designs', 'Design Files')


# This will be the directory created in MySewnet and managed by this program.  If you want to put other files in MySewnet, just put them in a different directory hierarchy
#   $EmbrodRootDirtop
# We will save the Instructions for any of the files into this instructions directory name under your user Documents folders
#   $instructions 


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

$ECCVERSION = "0.0.2"
write-host "Embroidery Collection Cleanup version: $ECCVERSION" -ForegroundColor Cyan

$filecnt = 0
$sizecnt = 0
$Global:dircnt = 0
$Global:savecnt = 0
$Global:addsizecnt = 0
$p = 0

$shell = New-Object -ComObject 'Shell.Application'
$downloaddir = (New-Object -ComObject Shell.Application).NameSpace('shell:Downloads').Self.Path
$docsdir =[environment]::getfolderpath("mydocuments")
$homedir = "${env:HOMEDRIVE}${env:HOMEPATH}"
$tmpdir = ${env:temp} + "\cleansew.tmp"

$doit = !$Testing

$SewTypeStar = @()
$foldupDirs = $foldupDir.tolower()
foreach ($t in $preferredSewType) {
    $SewTypeStar += "*." + $t
    $foldupDirs += $t
    }
$excludetypes =@()

foreach ($a in $alltypes) {
    $found = $false
    foreach ($t in $preferredSewType) {
        if ($t -match $a) { $found = $true }
        }
    if ($found) {
        if (!$includeEmbFilesWithInstruction) {
            $excludetypes += ("*." +$a)
            }
        }
    else {
        $excludetypes += ("*." +$a)
               
        }
    }


# This is for development testing and debugging

if ($env:COMPUTERNAME -eq "DESKTOP-R3PSDBU" -and $Testing) {
    $homedir  = "d:\Users\kjeff"
    $docsdir = "d:\Users\kjeff\OneDrive\Documents"
    $downloaddir = "d:\Users\kjeff\downloads"
    $doit = $true
    }

$FilesCollection = $docsdir
$CollectionTypeofStr = "Default Documents"

$cfgfile = "${env:LOCALAPPDATA}\Mysewnet\mySewnetCache\cclibrary.config"
if (test-path -path $cfgfile) {
    $CollectionTypeofStr = "MySewnet Cloud"
    get-content -path $cfgfile | where-object {$_ -like "files-path=*" } | 
        foreach { ($mp, $FilesCollection) = $_ -split "=" }
        if ($FilesCollection.substring($FilesCollection.length-1,1) -in @('\','/')) {
            $FilesCollection = $FilesCollection.substring(0,$FilesCollection.length-1)
        }
    write-verbose "Found MySewnet configuration file $FilesCollection" 
    if ((get-itemproperty -path "Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Classes\.vp3\ShellEx\{e357fccd-a995-4576-b01f-234630154e96}").'(default)' -ne "{370F9E36-A651-4BB3-89A9-A6DB957C63CC}") {
        write-host "** Install the Explorer Plug-in https://download.mysewnet.com/MSW/ so the pattern images appear in Windows Explorer" -ForegroundColor Yellow
        }
      
    }
#
# Building out all the directory structures and File lists
#


if (-not $doit){
    $PSDefaultParameterValues = @{
  "Copy-Item:WhatIf"=$True
  "Move-Item:WhatIf"=$True
  "Remove-Item:WhatIf"=$True
}
}

$EmbrodRootDir = $FilesCollection  + "\" + $EmbrodRootDirtop + "\" 
$instructionRoot = $docsdir + "\" + $instructions + "\"


Get-ChildItem -Path  ($tmpdir ) -Recurse | Remove-Item -force -Recurse


Function ShowSomeProgress ([string]$Area, [string]$stat = $null)
{
    $Global:p++
    If ($stat -eq $null -or $stat -eq "") {
        write-progress -PercentComplete ($Global:p % 100 ) $Area
    } else {
        write-progress -PercentComplete ($Global:p % 100 ) $Area -status $stat
    }
}
Function deleteToRecycle ($file) {
        if ($file.GetType().Name -ne "String") {
            $file = $file.FullName
            }
        if ($HardDelete) {
                Remove-Item -Path $file        # Handled by WhatIf
        } else {
            if ($doit){
                $shell.NameSpace(0).ParseName($file).InvokeVerb('delete')
            }
        }
}

Function MyPause ($message, [bool]$choice=$false, $boxmsg, [int]$timeout=0)
{
    # Check if running Powershell ISE
    if ($psISE)
    {
        Add-Type -AssemblyName System.Windows.Forms
        if ($choice) {
            $x = [System.Windows.Forms.MessageBox]::Show("$boxmsg",'Cleanup Collection Folders', 'YesNo', 'Question')
        } else {
            [System.Windows.Forms.MessageBox]::Show("$message")
            }
        return ($x -eq 'Yes')
    }
    else
    {
        $secondsRunning = 0;
        if ($timeout -gt 0) { $host.ui.RawUI.FlushInputBuffer()}  # get around bug in commandline
        Write-Host "$message" -ForegroundColor Yellow
        $isSpace = $true
        while( (-not $Host.UI.RawUI.KeyAvailable) -and ($secondsRunning -lt $timeout) ){

            Start-Sleep -Seconds 1
            $secondsRunning++
            }
        if ($Host.UI.RawUI.KeyAvailable -or $timeout -eq 0) {
            $x = $host.ui.RawUI.ReadKey("NoEcho,IncludeKeyDown")
            # Return true is Space press
            $isSpace = $x.VirtualKeyCode -eq 32
            }
        return ($isSpace)
    }
}

Function tailRecursion ([string]$path, [int]$depth=0) {
        

    $found=$false
    foreach ($childDirectory in Get-ChildItem -Force -LiteralPath $Path -Directory) {
        tailRecursion $childDirectory.FullName ($depth+1)
        }
    $currentChildren = Get-ChildItem -Force -LiteralPath $Path
    $isEmpty = $currentChildren -eq $null
    # Don't remove the very top directory, but do remove sub-directories if empty
    if ($isEmpty -and $depth -gt 0) {
        Write-Verbose "Removing empty folder: '${Path}'" 
        deleteToRecycle $Path 
        $found = $true
        
        $Global:dircnt++
        
        ShowSomeProgress "Removing Directory" $Path
    }
    return ($found)
}

#-----------------------------------------------------------------------
# Move From Directory to either Embrodery or Instrctions directory
function MoveFromDir ( 
        [string] $fromPath, 
        [boolean]$isEmbrodery = $false,        # is this to the Embrodery directory (true) Or to the Instruction Directory (false)
        [string]$whichfiles = "*.*", 
        [string[]]$files = $null
        
    ) 
{
    
    ShowSomeProgress "Copying" "Added ${Global:savecnt} files"
    if ($isEmbrodery) { 
        $dtype = 'Embroidery' 
        $objs = Get-ChildItem -Path $fromPath -include $SewTypeStar -File -Recurse -filter $whichfiles
        $targetdir = $EmbrodRootDir
    } else { 
        # Move anything that is not a Embrodery type file (alltypes)
        $dtype = 'Instructions'
        $objs = Get-ChildItem -Path $fromPath -exclude ($excludetypes +$TandCs)  -File -Recurse -filter $whichfiles
        $targetdir = $instructionRoot
        }

    $objs | ForEach-Object {
        if (($files -eq $null) -or ($_.Name -in $files)) {                
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
            
            
            if (!(Test-Path -Path ($targetdir + $newdir) -PathType Container)) {
                $null = New-Item -Path ($targetdir + $newdir) -ItemType Directory  }
                $npath = (Join-Path -Path ($targetdir + $newdir) -ChildPath $newfile )
            if (test-path $npath) {  # See if the file already exists
                if (!$isEmbrodery) {
                    if (Get-FileHash($npath) -eq Get-Filehash($_)) {
                        Remove-Item -Path $_
                        Write-Verbose "Removed Duplicate ${dtype} file :'$_'" 
                        }
                    }
                Write-Verbose "Skipping ${dtype}:'$_' to ${newdir}" 
            } else {
                $_ | Move-Item -Destination $npath  # -ErrorAction SilentlyContinue
                Write-Verbose "Saving ${dtype}:'$_' to ${newdir}"
                if ($isEmbrodery) { 
                    $Global:addsizecnt = $Global:addsizecnt + (Get-Item -Path $npath).Length 
                    $Global:savecnt = $Global:savecnt + 1
                    }
                }
            } 
            else {
                Write-Verbose "Skipping ${_.Name}" 
            }
    
    ShowSomeProgress "Copying" -Status "Added ${Global:savecnt} files"
    }
}

#-----------------------------------------------------------------------
# Format a Size string in KB/MB/GB 
#
function niceSize ($sz)   {
    $ext = " B"
    if ($sz -gt 1024) {
        $ext = " KB"
        $sz = $sz/1024
        if ($sz -gt 1024) {
            $ext = " MB"
            $sz = $sz/1024
            if ($sz -gt 1024) {
                $ext = " GB"
                $sz = $sz/1024
                }
                if ($sz -gt 1024) {
                    $ext = " TB"
                    $sz = $sz/1024
                    }
            }
        }
    return ([Math]::Round($sz,1)).toString() + $ext
    }


#-----------------------------------------------------------------------


Write-Host " ".padright(15) "Begin Embroidery Filing".padright(60) -ForegroundColor white -BackgroundColor blue
Write-Host " ".padright(15) $("Checking for Zips in the last $DownloadDaysOld days".padright(60)) -ForegroundColor white -BackgroundColor blue

$failed = $false
if (!( test-path -Path $FilesCollection )) {
    Write-Host "Can not find the Main directory for $$CollectionTypeofStr ($FilesCollection).  Stopping" 
    $failed = $true
    }
    
if (!( test-path -Path $EmbrodRootDir)) {
    Write-Host "Can not find the main directory ($EmbrodRootDir) within $FilesCollection.  Stopping" 
    Write-Host "Usually create '$EmbrodRootDirtop' in the with $FilesCollection directory.  Create the directory if this is your first time."
    $failed = $true
    }
if (!( test-path -Path $instructionRoot)) {
    Write-Host "Can not find the Instruction Directory ($instructionRoot).  Stopping"
    Write-Host "Usually created in the Documents directory ($docsdir).  Create the directory if this is your first time"
    $failed = $true
    } 
if ($failed) {
    Write-host " ".padright(80) -BackgroundColor DarkRed -ForegroundColor White
    write-host "See instructions at https://github.com/D-Jeffrey/Embroidery-File-Organize"
    $failed = MyPause 'Press any key to Close' 

    break
    }
    
Write-Host    "Download source directory           : $downloaddir" 
Write-host    "$CollectionTypeofStr sub folder directory : $EmbrodRootDir" 
Write-host    "Instructions directory              : $instructionRoot"
Write-host    "File types                          : $SewTypeStar"
Write-host    "Age of files in Download directory  : $DownloadDaysOld"
Write-host    "Clean the Collection folder         : $($CleanCollection -or $CleanCollectionIgnoreDir)"
Write-host    "Keep all variations of files types  : $keepAllTypes"
if ($Testing) {
    Write-Host "Testing Mode                        : $Testing" -ForegroundColor Yellow
    }
Write-Verbose "Rollup match pattern                : $foldupDirs"
Write-Verbose "Ignore Terms Conditions files       : $TandCs"
Write-Verbose "Excludetypes                        : $excludetypes"

$cont = (MyPause 'Press Start to continue, any other key to stop (Auto starting in 3 seconds)'  $true 'Click Yes to start' 3) 

if (!$cont) { 
    Break
    }
Add-Type -assembly "system.io.compression.filesystem"

ShowSomeProgress  "Calculating size"
$librarySizeBefore = 0
if ($CollectionTypeofStr -eq "MySewnet Cloud") {
    Get-ChildItem -Path ($EmbrodRootDir + "..")  -Recurse -file  | ForEach-Object { $librarySizeBefore = $librarySizeBefore + $_.Length}
    }
else {
    Get-ChildItem -Path ($EmbrodRootDir)  -Recurse -file  | ForEach-Object { $librarySizeBefore = $librarySizeBefore + $_.Length}
    }

ShowSomeProgress  "Loading file list"
$mysewingfiles = $null
# Get a list of all the existing files in mySewnet
        
$mysewingfiles = (Get-ChildItem -Path $EmbrodRootDir  -Recurse -file -include $SewTypeStar)| ForEach-Object { 
    if ($keepAllTypes) {
        $n = $_.Name} 
    else {
        $n = $_.BaseName
    }
    [PSCustomObject]@{ 
        Name = $n
        N = $_.Name 
        Base = $_.BaseName
        Ext = $_.Extension
        DirectoryName = $_.DirectoryName
        DirBase = $_.DirectoryName + $_.BaseName
        Priority = $preferredSewType.Indexof($_.Extension.substring(1,$_.Extension.Length-1))
        } 
    }

if ($mysewingfiles -eq $null) {
    $mysewingfiles =    @([PSCustomObject]@{ 
            Name ="mysewingfiles.placeholder"
            N = "mysewingfiles.placeholder"
            Base  = "mysewingfiles"
            Ext  = "placeholder"
            DirectoryName = "DirectoryName"
            DirBase = "DirBase"
            Priority = 100
            })
    }

# $mysewingfiles | ft
$isNewInstruct = $false
Get-ChildItem -Path $downloaddir  -file -filter "*.zip" | Where-Object { $_.CreationTime -gt (Get-Date).AddDays(- $DownloadDaysOld ) } |
  
    ForEach-Object {
        $thisFile = $_
        $thisfileBase = $_.BaseName
        Write-Verbose "Checking ZIP '$zips'" 
        $zips = $_.FullName
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
                foreach ($f in $($myfl.Name )) {
                    $fs = $f
                    if ($f -match "\.") { $fs = $($f -split "\.")[0] }
                    if (-not $keepAllTypes) {
                        $f = $fs
                        }
                    if ($f -in $mysewingfiles.Name) {
                        # Check for duplicate FileHash
                        # Find the instances then get the full name then get hash to compare
                        $fileInstance = $mysewingfiles | where-object {$.Name -eq $f}
                        $fiName = $fileInstance.DirBase + "." + $fileInstance.Ext
                        Write-verbose "Comparing FileHash '${f}' to ${fiName}"
                        
                        get-FileHas($fiName)
                        somehow we need to queue this comparison

                        Write-verbose "Duplicate file '${f}'"
                    } else {
                        Write-verbose "New file '${f}'"
                        $numnew++
                        $isnew  = $true
                        if ($keepAllTypes) {
                            $filesOfThisType += $f
                        } else {
                            $filesOfThisType += ($f + "." + $t)
                        if ($keepAllTypes) {
                           $n = $_.Name 
                        } else {
                            $n = $fs
                            }
                        $mysewingfiles +=  
                            [PSCustomObject]@{ 
                                Name = $n
                                N = $f
                                Base = $fs
                                DirectoryName = "????????????"
                                DirBase = ""
                                Priority = $preferredSewType.Indexof($t)
                                } 
                            }
                        }
                    }
                
                # we found a new file in the Zip.  If we have not expanded this Zip, then do it now
                if ($isnew) { 
                    if (-not $isNewInstruct) { 
                        if ($filesOfThisType.count -gt $aSetSizeis) {     # This is a set we should keep together
                            $madeDir = "\" + $thisfileBase -replace ' \(([0-9]+)\)'
                            
                            }
                        $resultTmpDir = $tmpdir + $madeDir
                        $isNewInstruct = $true
                        Expand-Archive -Path $zips -DestinationPath $resultTmpDir
                        }
                  
                    MoveFromDir $tmpdir $true $ts $filesOfThisType
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
            MoveFromDir $tmpdir $false
            Get-ChildItem -Path $tmpdir -Recurse | Remove-Item -force -Recurse
            }
   
        $filelist = $null      # Close Zipfile
        ShowSomeProgress  "Checking Zips"  "Added $Global:savecnt files"

    }

# $mysewingfiles | ft

# Look for Files which are not part of a ZIP file, just the selected file types that we are looking for that is in the download directory
$DownloadDaysOld = 365*10  # 10 years of downloads
$ppp = 0
foreach ($t in $preferredSewType) {

    $ts = "*."+ $t
    Get-ChildItem -Path $downloaddir  -file -include $ts -Depth 1 -Recurse| Where-Object { $_.CreationTime -gt (Get-Date).AddDays(- $DownloadDaysOld ) } |
        ForEach-Object {
            $f = $_.Name
            $fs = $_.BaseName
            $d = $_.DirectoryName 
            if (-not $keepAllTypes) {
                    $f = $fs
                    }
            if ($f -in $mysewingfiles.Name) {
                Write-verbose "Duplicate file '${f}'"
            } else {
                if (!(test-path -path (join-path -Path $EmbrodRootDir -ChildPath $_.Name))) { 
                    $_ | Copy-Item -Destination $EmbrodRootDir -ErrorAction SilentlyContinue 
                    Write-Verbose "Copied from Download :'$($_.Name)' to $EmbrodRootDir"
                    $fd = join-Path -Path $d -ChildPath ($fs +".pdf")
                

                    if (test-path -path $fd) {
                        Copy-Item -Path $fd -Destination $instructionRoot -ErrorAction SilentlyContinue 
                        Write-Verbose "Copied instructions from Download :'$($_.Name)' to $instructionRoot"
                        }
                    $Global:addsizecnt = $Global:addsizecnt + $_.Length 
                    $Global:savecnt = $Global:savecnt + 1 
                
                    $mysewingfiles +=  [PSCustomObject]@{ 
                                Name = $f
                                N = $f
                                Base = $fs
                                DirectoryName = $d
                                Priority = $preferredSewType.Indexof($t)
                       
                    }
                }    
            }
        $ppp++
        if ($ppp % 5 -eq 0) {
            ShowSomeProgress  "Copying from Downloads" 
            }
        }    
    }

            
# $mysewingfiles | ft
if ($CleanCollection -or $CleanCollectionIgnoreDir) {
    write-host "Scanning for files to clean up in mySewnet"
    $filteredfiles = @([PSCustomObject]@{
        Base = "filteredfiles" 
        DirBase = ""
        Ext = "PSCustomObject"
            })
    # Reload MySewingfiles
    $mysewingfiles = (Get-ChildItem -Path $EmbrodRootDir  -Recurse -file -include $SewTypeStar)| ForEach-Object { 
        if ($keepAllTypes) {
            $n = $_.Name} 
        else {
            $n = $_.BaseName
        }
        [PSCustomObject]@{ 
            Name = $n
            N = $_.Name 
            Base = $_.BaseName
            Ext = $_.Extension
            DirectoryName = $_.DirectoryName
            Priority = $preferredSewType.Indexof($_.Extension.substring(1,$_.Extension.Length-1))
            } 
        }

    # Find the priroity files types and keep those
    foreach ($pt in $SewTypeStar) {
        $mysw = $mysewingfiles | where-object {$_.Ext -like $pt} 
        foreach ($m in $mysw) {
            if ($m.Base -notin $filteredfiles.Base) {
                $filteredfiles += [PSCustomObject]@{ 
                                Base = $m.Base.ToLower()
                                DirBase = ($m.DirectoryName + $m.Base).ToLower()
                                Ext = $m.Ext.ToLower()
                }
            }
            $ppp++
            if ($ppp % 20 -eq 0) {
                ShowSomeProgress  "Scanning" 
            }
        
        }
        
    }
 
    # $filteredfiles | out-GridView -wait
       
    $filesToRemove = @()
    
    Get-ChildItem -Path $EmbrodRootDir  -Recurse -file -Exclude $SewTypeStar  -filter "*.*" -include $excludetypes | ForEach-Object {
       if ($_.BaseName -in $filteredfiles.Base) {   # Is this a file we already have a good version of?
            $sizecnt = $sizecnt + $_.Length
            Write-Verbose "Pending Removing '$_'" 
            $filesToRemove += $_ 
            $filecnt = $filecnt + 1
            }
        $ppp++
        if ($ppp % 20 -eq 0) {
            ShowSomeProgress   "Checking For unneeded formats"
            }
        }
    if (!$KeepAllTypes) {
        Get-ChildItem -Path $EmbrodRootDir  -Recurse -file -Include $SewTypeStar -filter "*.*" | ForEach-Object {
            $best = -2
            if ($CleanCollectionIgnoreDir) {
                if (($_.BaseName) -in $filteredfiles.Base) {   # Check is this is the based version of the file
                    $best = $filteredfiles.Base.Indexof(($_.BaseName).ToLower())
                    }
                }
             else {
                if (($_.DirectoryName + $_.BaseName) -in $filteredfiles.DirBase) {   # Check is this is the based version of the file
                    $best = $filteredfiles.DirBase.Indexof(($_.DirectoryName + $_.BaseName).ToLower())
                    }
                }
            if ($best -eq -1) {
                    write-host "Problem with IndexOf - $_    $best"
                }
            elseif ($best -ge 0) {
                if ($filteredfiles[$best].Ext -ne ($_.Extension.ToLower())) {
                    $sizecnt = $sizecnt + $_.Length
                    Write-Verbose "Pending Removing '$_'" 
                    $filesToRemove += $_ 
                    $filecnt = $filecnt + 1
                    }
                }
            $ppp++
            if ($ppp % 20 -eq 0) {
                ShowSomeProgress   "Checking MySewnet"
                }
            }
        }
    $fcr = $filesToRemove.count
    if ($fcr  -gt 0) {
        write-host "Found ${fcr} files that should be removed to clean up extras"
        $filesToRemove.FullName | Out-GridView -Title "Files that will be removed if you press Space (Close this Windows to continue)" -wait
        $cont = (MyPause 'Press Start to remove those files, any other key to keep them'  $true 'Click Yes to remove them') 

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
            ForEach ($f in $filesToRemove) {
                deleteToRecycle ($f)
                ShowSomeProgress  ($howDeleted  + "extra files from MySewnet") "$fcs of $fcr - $($f.Name)"
                $fcs++
                }
            }
        }

    write-host "Moving Instructions to proper Instructions directory"
    foreach ($g in $goodInstructionTypes) {
        MoveFromDir $EmbrodRootDir $false ("*."+ $g)        
        }
    }


#  Clear out empty Directories
if (-not $KeepEmptyDirectory) {
    $tailr = 0    # Loop thru 8 times to remove empty directories, then go back and check to see if you made any more emty
    while ($tailr -le 8 -and (tailRecursion $EmbrodRootDir) ) {
         $tailr++
    }
}
write-host "Calculating size"
ShowSomeProgress  "Calculating size"
$librarySizeAfter = 0
if ($CollectionTypeofStr -eq "MySewnet Cloud") {
    Get-ChildItem -Path ($EmbrodRootDir  + "..") -Recurse -file  | ForEach-Object { $librarySizeAfter = $librarySizeAfter + $_.Length}
    }
else {
    Get-ChildItem -Path ($EmbrodRootDir  ) -Recurse -file  | ForEach-Object { $librarySizeAfter = $librarySizeAfter + $_.Length}
    }


  
$librarySizeBefore = niceSize $librarySizeBefore
$librarySizeAfter = niceSize $librarySizeAfter
$sizecntB = niceSize  $sizecnt
$addsizecntB = niceSize $Global:addsizecnt
write-progress -PercentComplete  100  "Done"
if ($Global:dircnt -gt 0 -or $filecnt -gt 0) {
    Write-Host "Cleaned up - Directories removed: '$Global:dircnt    Files removed : '$filecnt' ($sizecntB)." -ForegroundColor Green
    }
if ($Global:savecnt -gt 0) {
    write-host "Added files to ${CollectionTypeofStr}: '${Global:savecnt}' ($addsizecntB) " -ForegroundColor Green
    }
Write-host "   *** $CollectionTypeofStr size is now : $librarySizeAfter was $librarySizeBefore ****   "  -ForegroundColor Green 
$none = MyPause 'Press any key to Close'
Write-Host ( "End") -ForegroundColor Green
