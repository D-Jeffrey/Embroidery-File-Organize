#
# MySewingNet-Cleanup
#
# Deal with the many different types of embroidery files, put the right format types in mySewingnet Cloud
# We are looking to keep the ??? files types only from the zip files.
#
# Orginal Author: Darren Jeffrey Dec 2021
#                           Last Dec 2022
#

param
(
  [Parameter(Mandatory = $false)]
  [Switch]$includeEmbFilesWithInstruction,            # Put the Instruction along with the Embrodery files (NOT recommended as the PDFs tend to take up a lot of space)
  [Switch]$CleanSewNet,                               # Cleanup the MySewingNet Folder to only EmbrodRootDirtop files
  [int32]$DownloadDaysOld = 7,                        # How many days old show it scan for Zip files in Download
  [Switch]$KeepAllTypes,                              # Keep all the different types of a file (duplicate name but different extensions)
  [Switch]$Testing,                                   # Run it and see what happens
  [string]$EmbrodRootDirtop = "Embroidery",           # You may want to change this directory name inside of the MySewingNet Directory
  [string]$instructions = "Embroidery Instructions"   # This is a Directory name inside of "Documents" where instructions are saved
)

# ******** CONFIGURATION 
$preferredSewType = ('vp4', 'vp3',  'vip', 'pcs', 'dst')
$alltypes =('hus','dst','exp','jef','pes','vip','vp3','xxx','sew',
    'vp4','pcs','vf3','csd','zsk','emd','ese','phc','art','ofm')
$TandCs = @('TERMS-OF-USAGE.*', 'planetappliquetermsandconditions.*')
$foldupDir = @('images','sewing helps')


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

$dircnt = 0
$filecnt = 0
$sizecnt = 0
$savecnt = 0
$addsizecnt = 0
$p = 0

$shell = New-Object -ComObject 'Shell.Application'


$downloaddir = (New-Object -ComObject Shell.Application).NameSpace('shell:Downloads').Self.Path
$docsdir =[environment]::getfolderpath("mydocuments")
$homedir = "${env:HOMEDRIVE}${env:HOMEPATH}"
$mySewTypeStar = @()
$foldupDirs = $foldupDir
foreach ($t in $preferredSewType) {
    $mySewTypeStar += "*." + $t
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

# $foldupDirs = "(" + $foldupDirs + ")"
$MySewNetCloud = $homedir + "\mySewnet Cloud"

$tmpdir = ${env:temp} + "\cleansew.tmp"
$doit = !$Testing

if ($env:COMPUTERNAME -eq "DESKTOP-R3PSDBU" -and $Testing) {
    $MySewNetCloud = "d:\Users\kjeff\mySewnet Cloud"
    $docsdir = "d:\Users\kjeff\OneDrive\Documents"
    $doit = $true
    }


$EmbrodRootDir = $MySewNetCloud  + "\" + $EmbrodRootDirtop + "\" 
$instructionRoot = $docsdir + "\" + $instructions + "\"


Get-ChildItem -Path  ($tmpdir ) -Recurse | Remove-Item -force -Recurse


Function deleteToRecycle ($file) {        
    $shell.NameSpace(0).ParseName($file.FullName).InvokeVerb('delete')
}

Function MyPause ($message, [bool]$choice=$false, $boxmsg)
{
    # Check if running Powershell ISE
    if ($psISE)
    {
        Add-Type -AssemblyName System.Windows.Forms
        if ($choice) {
            $x = [System.Windows.Forms.MessageBox]::Show("$boxmsg",'Cleanup MySewingNet', 'YesNo', 'Question')
        } else {
            [System.Windows.Forms.MessageBox]::Show("$message")
            }
        return ($x -eq 'Yes')
    }
    else
    {
        Write-Host "$message" -ForegroundColor Yellow
        $x = $host.ui.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        # Return true is Space press
        return ($x.VirtualKeyCode -eq 32)
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
        if ($doit) { 
            deleteToRecycle $Path 
            $found = $true
            }
        $dircnt + $dircnt + 1
        $p += 1
        write-progress -PercentComplete ($p % 100 ) "Removing Directory"
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
    
    write-progress -PercentComplete  ($p % 100 ) "Copying" -Status "Added $savecnt files"
    if ($isEmbrodery) { 
        $dtype = 'Embroidery' 
        $objs = Get-ChildItem -Path $fromPath -include $mySewTypeStar -File -Recurse -filter $whichfiles
        $targetdir = $EmbrodRootDir
    } else { 
        # Move anything that is not a Embrodery type file (alltypes)
        $dtype = 'Instructions'
        $objs = Get-ChildItem -Path $fromPath -exclude ($excludetypes +$TandCs)  -File -Recurse -filter $whichfiles
        $targetdir = $instructionRoot
        }

    $objs | ForEach-Object {
        if (($files -eq $null) -or ($_.Name -in $files)) {                
            $newdir =  (Split-Path(($_.FullName).substring(($fromPath.Length), 
                                ($_.FullName).Length - ($fromPath.Length) )))
            
            $newfile = $_.Name
        
            # take off the directory name if it is one of the rollup names
        
            foreach ($r in $foldupDirs) {
                if ($newdir.ToLower().EndsWith("\"+$r.ToLower())) {
                    #strip off the directory name and perserve the case of the directory and files
                    $newdir = $newdir.substring(0,($newdir.tolower().Replace("\"+$r,'')).length)
                    }
                }
            
            if ($doit) {
                if (!(Test-Path -Path ($targetdir + $newdir) -PathType Container)) {
                    $null = New-Item -Path ($targetdir + $newdir) -ItemType Directory  }
                    $npath = (Join-Path -Path ($targetdir + $newdir) -ChildPath $newfile )
                if (test-path $npath) {
                    Write-Verbose "Skipping ${dtype}:'$_' to ${newdir}" 
                } else {
                    $_ | Move-Item -Destination $npath  # -ErrorAction SilentlyContinue
                    Write-Verbose "Saving ${dtype}:'$_' to ${newdir}"
                    $addsizecnt = $addsizecnt + $_.Length 
                }
            }
            else { 
                Write-Verbose "Would save ${dtype}:'$_' to ${newdir}" 
                }
            
        } else 
        {
                Write-Verbose "Skipping ${_.Name}" 
        }
    $savecnt = $savecnt + 1
    $p += 1
    write-progress -PercentComplete  ($p % 100 ) "Copying" -Status "Added $savecnt files"
    }
}

#-----------------------------------------------------------------------
# Format a Size string in KB/MB/GB 
#
function niceSize ($sz)   {
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
    return ([Math]::Round($sz,1)).toString() + $ext
    }


#-----------------------------------------------------------------------


Write-Host $(" " * 15) "Begin MySewnet Process"  $(" " * 28) -ForegroundColor white -BackgroundColor blue
Write-Host $(" " * 15) "Checking for Zips in the last $DownloadDaysOld days" $(" " * (14-[Math]::floor([Math]::log10($DownloadDaysOld)))) -ForegroundColor white -BackgroundColor blue

if (!( test-path -Path $MySewNetCloud )) {
    Write-Host "Can not find the Main directory for MySewnet ($MySewNetCloud).  Stopping" -BackgroundColor DarkRed -ForegroundColor White
    write-host "See instructions at   https://github.com/D-Jeffrey/MySewingNet-Cleanup"
    MyPause 'Press any key to Close'

    break
    }

if (!( test-path -Path $EmbrodRootDir)) {
    Write-Host "Can not find the MySewnet Directory ($EmbrodRootDir).  Stopping" -BackgroundColor DarkRed -ForegroundColor White
    Write-Host "Usually create '$EmbrodRootDirtop' in the with the MySewnet directory.  Create the directory if this is your first time"
    write-host "See instructions at   https://github.com/D-Jeffrey/MySewingNet-Cleanup"
    MyPause 'Press any key to Close'

    break
    }
if (!( test-path -Path $instructionRoot)) {
    Write-Host "Can not find the Instruction Directory ($instructionRoot).  Stopping" -BackgroundColor DarkRed -ForegroundColor White
    Write-Host "Usually created in the Documents directory.  Create the directory if this is your first time"
    write-host "See instructions at   https://github.com/D-Jeffrey/MySewingNet-Cleanup"
    MyPause 'Press any key to Close'

    break
    }


Write-Host    ("Download source directory          : " + $downloaddir)
Write-host    ("mySewnet Cloud sub folder directory: " + $EmbrodRootDir)
Write-host    ("Instructions directory             : " + $instructionRoot)
Write-host    ("File types                         : " + $mySewTypeStar)
Write-host    ("Age of files in Download directory : " + $DownloadDaysOld)
Write-host    ("Clean the mysewingNet cloud folder : " + $CleanSewNet)
Write-host    ("Keep all variations of files types : " + $keepAllTypes)
if ($Testing) {
    Write-Host    ("Testing                            : " + $Testing) -ForegroundColor Yellow
    }
Write-Verbose ("Rollup match pattern               : " + $foldupDirs)
Write-Verbose ("Ignore Terms Conditions files      : " + $TandCs)
Write-Verbose ("Excludetypes                       : " + $excludetypes)

$cont = (MyPause 'Press Start to continue, any other key to stop'  $true 'Click Yes to start') 

if (!$cont) { 
    Break
    }
Add-Type -assembly "system.io.compression.filesystem"

$librarySizeBefore = 0
Get-ChildItem -Path ($EmbrodRootDir + "..")  -Recurse -file  | ForEach-Object { $librarySizeBefore = $librarySizeBefore + $_.Length}


$mysewingfiles = $null
# Get a list of all the existing files in mySewnet
        
$mysewingfiles = (Get-ChildItem -Path $EmbrodRootDir  -Recurse -file -include $mySewTypeStar)| ForEach-Object { 
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
        Priority = $preferredSewType.Indexof($_.Extension)
        } 
    }

if ($mysewingfiles -eq $null) {
    $mysewingfiles =    @([PSCustomObject]@{ 
            Name ="mysewingfiles.placeholder"
            N = "mysewingfiles.placeholder"
            Base  = "mysewingfiles"
            Ext  = "placeholder"
            Priority = 100
            })
    }

# $mysewingfiles | ft

Get-ChildItem -Path $downloaddir  -file -filter "*.zip" | Where-Object { $_.CreationTime -gt (Get-Date).AddDays(- $DownloadDaysOld ) } |
    ForEach-Object {
        $thisFile = $_
        $zips = $_.FullName
        Write-Verbose "Checking ZIP '$zips'" 
        $zips = $_.FullName
        $filelist = [io.compression.zipfile]::OpenRead($zips).Entries.Name
        
        $isNewInstruct = $false
        foreach ($t in $preferredSewType) {
            $filesOfThisType = @()
            $ts = "*."+ $t
            if ($filelist -like $ts) {
                Write-host "Found ZIP: '$zips'" 
                $isnew = $false
           
                foreach ($f in $($filelist -like $ts)) {
                    $fs = $f
                    if ($f -match "\.") { $fs = $($f -split "\.")[0] }
                    if (-not $keepAllTypes) {
                        $f = $fs
                        }
                    if ($f -in $mysewingfiles.Name) {
                        Write-verbose "Duplicate file '${f}'"
                    } else {
                        Write-verbose "New file '${f}'"
                        $isnew  = $true
                        if ($keepAllTypes) {
                            $filesOfThisType += $f
                        } else {
                            $filesOfThisType += ($f + "." + $t)
                            }
                        if ($keepAllTypes) {
                           $n = $_.Name} 
                        else {
                            $n = $fs
                            }
                        $mysewingfiles +=  
                            [PSCustomObject]@{ 
                                Name = $n
                                N = $f
                                Base = $fs
                                Priority = $preferredSewType.Indexof($t)
                                } 
                            }
                        }
                    }
                    
                # we found a new file in the Zip.  If we have not expanded this Zip, then do it now
                if ($isnew) { 
                    Write-host "** New files in ZIP: '$zips'" 
                    if (-not $isNewInstruct) { 
                        $isNewInstruct = $true
                        Expand-Archive -Path $zips -DestinationPath $tmpdir
                    }
                
                
                    MoveFromDir $tmpdir $true $ts $filesOfThisType
                }
                
            }
        # we extracted the Zip already and now let's check for instructions
        if ($isNewInstruct) { 
            MoveFromDir $tmpdir $false
            Get-ChildItem -Path $tmpdir -Recurse | Remove-Item -force -Recurse
            
        }
        $filelist = $null      # Close Zipfile
        $p += 1
        write-progress "Checking Zips" -PercentComplete ($p % 100 ) -Status "Added $savecnt files"

    }

# $mysewingfiles | ft

# Look for Files which are not part of a ZIP file, just the selected file types that we are looking for that is in the download directory
$DownloadDaysOld = 365*10  # 10 years of downloads
foreach ($t in $preferredSewType) {
    $filesOfThisType = @()
    $ts = "*."+ $t
    Get-ChildItem -Path $downloaddir  -file -include $ts -Depth 1 -Recurse| Where-Object { $_.CreationTime -gt (Get-Date).AddDays(- $DownloadDaysOld ) } |
        ForEach-Object {
              $f = $_.Name
              if (-not $keepAllTypes) {
                    $f = $($f -split "\.")[0]
                    }
             if ($f -in $mysewingfiles.Name) {
                        Write-verbose "Duplicate file '${f}'"
            } else {
                if (!(test-path -path (join-path -Path $EmbrodRootDir -ChildPath $_.Name))) { 
                    $_ | Copy-Item -Destination $EmbrodRootDir -ErrorAction SilentlyContinue
                    Write-Verbose "Copied from Download :'$($_.Name)' to $EmbrodRootDir"
                    $addsizecnt = $addsizecnt + $_.Length
                    $savecnt = $savecnt + 1 
                    $n = $f                  
                    $mysewingfiles +=  [PSCustomObject]@{ 
                                Name = $n
                                N = $f
                                # Base = $($f -split "\.")[0]
                                Priority = $preferredSewType.Indexof($t)
                        } 
                    }
                }    
            }
      
    }
            
# $mysewingfiles | ft
if ($CleanSewNet) {
    Get-ChildItem -Path $EmbrodRootDir  -Recurse -file -Exclude $mySewTypeStar  | ForEach-Object {
     
            $sizecnt = $sizecnt + $_.Length
            Write-Verbose "Removing '$_'" 
            if ($doit) { deleteToRecycle  $_ }
            $filecnt = $filecnt + 1
            $p += 1
            write-progress -PercentComplete  ($p % 100 ) "Updating MySewnet"
        }
    
    MoveFromDir $EmbrodRootDir $false
    }


#  Clear out empty Directories
$tailr = 0
while ($tailr -le 8 -and (tailRecursion $EmbrodRootDir) ) {
     $tailr = $tailr + 1
     write-progress -PercentComplete  ($p % 100 ) "Clean directories"
}

$librarySizeAfter = 0
Get-ChildItem -Path ($EmbrodRootDir  + "..") -Recurse -file  | ForEach-Object { $librarySizeAfter = $librarySizeAfter + $_.Length}

  
$librarySizeBefore = niceSize $librarySizeBefore
$librarySizeAfter = niceSize $librarySizeAfter
$kb = niceSize  $sizecnt
$akb = niceSize $addsizecnt
write-progress -PercentComplete  100  "Done"
Write-host "   *** MySewnet Cloud size is now : $librarySizeAfter was $librarySizeBefore ****   "  -ForegroundColor Blue Blue -BackgroundColor White
Write-Host ("Cleaned up - Dirs removed: '$dircnt' Saved from zip files: '$savecnt' ($akb) Removed: '$filecnt' ($kb).") -ForegroundColor Green

MyPause 'Press any key to Close'
Write-Host ( "End") -ForegroundColor Green
