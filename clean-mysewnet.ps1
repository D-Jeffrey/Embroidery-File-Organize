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
  [bool]$includeEmbFilesWithInstruction = $false,
  [bool]$CleansewingNet = $false,             # Cleanup the MySewingNet Folder to only Embrodery files
  [int32]$DownloadDaysOld = 7,                       # How many days old show it scan for Zip files in Download
  [bool]$keepAllTypes = $false,                     # Keep all the different types of a file (duplicate name but different extensions)
  [string]$EmbrodRootDirtop = "Embroidery",
  [string]$instructions = "Embroidery Instructions"  
)

# This will be the directory created in MySewnet and managed by this program.  If you want to put other files in MySewnet, just put them in a different directory hierarchy
#   $EmbrodRootDirtop
# We will save the Instructions for any of the files into this instructions directory name under your user Documents folders
#   $instructions 

$mysewtypes = ('vp4', 'vp3',  'pcs', 'dst')

# ----------------------------------------------------------------------
# this is a list of all the different types of embrodiary files that are considered.  
# The 'mysewtypes' should be from the list below based on what is best for your machine
$alltypes =('hus','dst','exp','jef','pes','vip','vp3','xxx','sew',
    'vp4','pcs','VF3','CSD','ZSK','EMD','ESE','PHC','ART','OFM')

# Term and Conditions added by various store that add up space with the same document type over and over, using up your MySewing Cloud space
# This is a file name pattern so TC.* will match TC.doc or TC.pdf
$TCs = @('TERMS-OF-USAGE.*', 'planetappliquetermsandconditions.*')
# What directories should be flattened to bring the Embroidery files higher up so they are not nested instead of sub-folders.  
# The names are for Directories you want to remove the sub-folder and moved the contents up
$rollup = @('images','sewing helps')



# ----------------------------------------------------------------------
#
#
$Testing = $true

$dircnt = 0
$filecnt = 0
$sizecnt = 0
$savecnt = 0
$addsizecnt = 0
$p = 0

$shell = New-Object -ComObject 'Shell.Application'


$downloaddir = (New-Object -ComObject Shell.Application).NameSpace('shell:Downloads').Self.Path
$docsdir =[environment]::getfolderpath(“mydocuments”)
$homedir = "${env:HOMEDRIVE}${env:HOMEPATH}"
$mysewtype = @()
$rollups = $rollup
foreach ($t in $mysewtypes) {
    $mysewtype += "*." + $t
    $rollups += $t
    }
$excludetypes =@()

foreach ($a in $alltypes) {
    $found = $false
    foreach ($t in $mysewtypes) {
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

# $rollups = "(" + $rollups + ")"
$EmbrodRootDir = $homedir + "\mySewnet Cloud\" + $EmbrodRootDirtop + "\" 
$instructionRoot = $docsdir + "\" + $instructions + "\"

$tmpdir = ${env:temp} + "\cleansew.tmp"
$doit = !$Testing





Get-ChildItem -Path  ($tmpdir ) -Recurse | Remove-Item -force -Recurse

if ($testing -and $env:COMPUTERNAME -eq "DESKTOP-R3PSDBU") {
    $EmbrodRootDir = "d:\Users\kjeff\mySewnet Cloud\Embroidery\"
    $instructionRoot = "d:\Users\kjeff\OneDrive\Documents\Embroidery Instruction\"
    $doit = $true
    }

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
            Write-Verbose "Removing empty folder: '${Path}'." 
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
            [boolean]$Moveit = $true
        ) 
    {
        
        
        write-progress -PercentComplete  ($p % 100 ) "Copying"
        if ($isEmbrodery) { 
            $dtype = 'Embroidery' 
            $objs = Get-ChildItem -Path $fromPath -include $mysewtype -File -Recurse 
            $targetdir = $EmbrodRootDir
        } else { 
            # Move anything that is not a Embrodery type file (alltypes)
            $dtype = 'Instructions'
            $objs = Get-ChildItem -Path $fromPath -exclude ($excludetypes +$TCs)  -File -Recurse 
            $targetdir = $instructionRoot
            }

        $objs | ForEach-Object {
                        
            $newdir =  (Split-Path(($_.FullName).substring(($fromPath.Length), 
                                ($_.FullName).Length - ($fromPath.Length) )))
            
            $newfile = $_.Name
        
            # take off the directory name if it is one of the rollup names
        
            foreach ($r in $rollups) {
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
                Write-Verbose "Skipping ${dtype}:'$_' to ${newdir}." 
                } else {
                if ($Moveit) {
                    $_ | Move-Item -Destination $npath -ErrorAction SilentlyContinue
                } else {
                    $_ | Copy-Item -Destination $npath -ErrorAction SilentlyContinue
                    }
                Write-Verbose "Saving ${dtype}:'$_' to ${newdir}."
                $addsizecnt = $addsizecnt + $_.Length 
                }
            }
            else { 
                Write-Verbose "Would save ${dtype}:'$_' to ${newdir}." 
                }

        $savecnt = $savecnt + 1
        $p += 1
        write-progress -PercentComplete  ($p % 100 ) "Copying"
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
            $sz = [Math]::Round($sz/1024)
            if ($sz -gt 1024) {
                $ext = " GB"
                $sz = $sz/1024
                }
                if ($sz -gt 1024) {
                    $ext = " TB"
                    $sz = $sz/1024
                    }
            }
        return ([Math]::Round($sz)).toString() + $ext
        }


#-----------------------------------------------------------------------


Write-Host $(" " * 15) "Begin Sewnet Process"  $(" " * 30) -ForegroundColor white -BackgroundColor blue
Write-Host $(" " * 15) "Checking for Zips in the last $DownloadDaysOld " $(" " * (18-[Math]::floor([Math]::log10($DownloadDaysOld)))) -ForegroundColor white -BackgroundColor blue

if (!( test-path -Path $EmbrodRootDir)) {
    Write-Host "Can not find the MySewnet Directory ($EmbrodRootDir).  Stopping" -BackgroundColor DarkRed -ForegroundColor White
    write-host "See instructions at   https://github.com/D-Jeffrey/MySewingNet-Cleanup"
    MyPause 'Press any key to Close'

    break
    }
if (!( test-path -Path $instructionRoot)) {
    Write-Host "Can not find the Instruction Directory ($instructionRoot).  Stopping" -BackgroundColor DarkRed -ForegroundColor White
    write-host "See instructions at   https://github.com/D-Jeffrey/MySewingNet-Cleanup"
    MyPause 'Press any key to Close'

    break
    }


Write-Host    ("Download source directory          : " + $downloaddir)
Write-host    ("mySewnet Cloud sub folder directory: " + $EmbrodRootDir)
Write-host    ("Instructions directory             : " + $instructionRoot)
Write-host    ("File types                         : " + $mysewtype)
Write-host    ("Age of files in Download directory : " + $DownloadDaysOld)
Write-host    ("Clean the mysewingNet cloud folder : " + $CleansewingNet)
Write-Verbose ("Rollup match pattern               : " + $rollups)
Write-Verbose ("Ignore Terms Conditions files      : " + $TCs)
Write-Verbose ("Testing                            : " + $Testing)
Write-Verbose ("Excludetypes                       : " + $excludetypes)

$cont = (MyPause 'Press Start to continue, any other key to stop'  $true 'Click Yes to start') 

if (!$cont) { 
    Break
    }
Add-Type -assembly "system.io.compression.filesystem"

$librarySizeBefore = 0
Get-ChildItem -Path ($EmbrodRootDir + "..")  -Recurse -file  | ForEach-Object { $librarySizeBefore = $librarySizeBefore + $_.Length}



# Get a list of all the existing files in mySewnet
if ($keepAllTypes) {
    $mysewingfiles = (Get-ChildItem -Path $EmbrodRootDir  -Recurse -file -include $mysewtype).Name 
} else {
    $mysewingfiles = (Get-ChildItem -Path $EmbrodRootDir  -Recurse -file -include $mysewtype).BaseName |foreach {$_ + ".*"}
}


Get-ChildItem -Path $downloaddir  -file -filter "*.zip" | Where-Object { $_.CreationTime -gt (Get-Date).AddDays(- $DownloadDaysOld ) } |
    ForEach-Object {
        $zips = $_.FullName
        Write-Verbose "Checking ZIP '$zips'." 
        $zips = $_.FullName
        $filelist = [io.compression.zipfile]::OpenRead($zips).Entries.Name
        
        foreach ($t in $mysewtypes) {
            if ($filelist -match($t)) {
                Write-host "Found ZIP: '$zips'." 
                $isnew = $true
                foreach ($m in $mysewingfiles) {
                
                    foreach ($f in $filelist) {
                        
                        if ($f -like $m) {
                            $isnew  = $false
                            Write-verbose "Duplicate file '${f}'."

                        }
                    }
                } 
            if ($isnew) {  
                Write-host "** New files in ZIP: '$zips'." 
                Expand-Archive -Path $zips -DestinationPath $tmpdir
                MoveFromDir $tmpdir $false
                MoveFromDir $tmpdir $true
                

                
                Get-ChildItem -Path $tmpdir -Recurse | Remove-Item -force -Recurse

                }
            }
            
        }
        
        $p += 1
        write-progress "Checking Zips" -PercentComplete ($p % 100 ) 

    }

$DownloadDaysOld = 365*100  # 100 years of downloads
Get-ChildItem -Path $downloaddir  -file -include $mysewtype -Depth 1 -Recurse| Where-Object { $_.CreationTime -gt (Get-Date).AddDays(- $DownloadDaysOld ) } |
    ForEach-Object {
        if (!(test-path -path (join-path -Path $EmbrodRootDir -ChildPath $_.Name))) { 
            $_ | Copy-Item -Destination $EmbrodRootDir -ErrorAction SilentlyContinue
            Write-Verbose "Copied from Download :'$_.Name}' to ${tree}."
            $addsizecnt = $addsizecnt + $_.Length
            $savecnt = $savecnt + 1 
            }
         }
            

if ($cleansewingNet) {
    Get-ChildItem -Path $EmbrodRootDir  -Recurse -file -Exclude $mysewtype  | ForEach-Object {
     
            $sizecnt = $sizecnt + $_.Length
            Write-Verbose "Removing '$_'." 
            if ($doit) { deleteToRecycle  $_ }
            $filecnt = $filecnt + 1
            $p += 1
            write-progress -PercentComplete  ($p % 100 ) "Updating MySewnet"
        }
    
    MoveFromDir $EmbrodRootDir, $false
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
Write-host "*** MySewnet Cloud size is now : $librarySizeAfter was $librarySizeBefore ****"  -ForegroundColor Blue
Write-Host ("Cleaned up - Dirs removed: '$dircnt' Save: '$savecnt' ($akb) Removed: '$filecnt' ($kb).") -ForegroundColor Green


$addsizecnt = $addsizecnt + $_.Length

MyPause 'Press any key to Close'
Write-Host ( "End") -ForegroundColor Green
