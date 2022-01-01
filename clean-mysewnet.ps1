#
# MySewingNet-Cleanup
#
# Deal with the many different types of embroidery files, put the right format types in mySewingnet Cloud
# We are looking to keep the ??? files types only from the zip files.
#
# Orginal Author: Darren Jeffrey Dec 2021
#

param
(
  [Parameter(Mandatory = $false)]
  [bool]$includeEmbFilesWithInstruction = $false,
  [bool]$CleansewingNet = $false,             # Cleanup the MySewingNet Folder to only Embrodery files
  [int32]$DownloadDaysOld = 7              # How many days old show it scan for Zip files in Download  
)

$treetop = "Embroidery"
$instructions = "Embroidery Instructions"

$mysewtypes = ('vp3', 'vp4')

# ----------------------------------------------------------------------

$alltypes =('hus','dst','exp','jef','pes','vip','vp3','xxx','sew',
    'vp4','pcs','VF3','CSD','ZSK','EMD','ESE','PHC','ART','OFM')


$TCs = @('TERMS-OF-USAGE.*', 'planetappliquetermsandconditions.*')
$rollup = @('images','sewing helps')


# ----------------------------------------------------------------------
#
#
$Testing = $false

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
$tree = $homedir + "\mySewnet Cloud\" + $treetop + "\" 
$savetree = $docsdir + "\" + $instructions + "\"

$tmpdir = ${env:temp} + "\cleansew.tmp"
$doit = !$Testing





Get-ChildItem -Path  ($tmpdir ) -Recurse | Remove-Item -force -Recurse

if ($testing) {
    $savetree = "d:\Users\kjeff\OneDrive\Documents\Embroidery Instruction\"
    $tree = "d:\Users\kjeff\mySewnet Cloud\Embroidery\"
    $doit = $true
    }

function deleteToRecycle ($file) {        
	$shell.NameSpace(0).ParseName($file.FullName).InvokeVerb('delete')
}

Function pause ($message, [bool]$choice=$false, $boxmsg)
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



function movefromdir (
        [string] $fromPath, 
        [boolean]$Moveit = $true,
        [boolean]$isEmb = $false
    ) 
{
    
    
    write-progress -PercentComplete  ($p % 100 ) "Copying"
    if ($isEmb) { 
        $dtype = 'Embroidery' 
        $objs = Get-ChildItem -Path $fromPath -include $mysewtype -File -Recurse 
        $targetdir = $tree
    } else { 
        $dtype = 'Instructions'
        $objs = Get-ChildItem -Path $fromPath -exclude ($excludetypes +$TCs)  -File -Recurse 
        $targetdir = $savetree
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
              Write-Verbose "Skiping ${dtype}:'$_' to ${newdir}." 
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


Write-Host "Begin Sewnet Process" -ForegroundColor white -BackgroundColor blue

Write-Host "Checking for Zips in the last $DownloadDaysOld " -ForegroundColor white -BackgroundColor blue



if (!( test-path -Path $tree)) {
    Write-Host "Can not find the MySewnet Directory ($tree).  Stopping" -BackgroundColor DarkRed -ForegroundColor White
    write-host "See instructions at   https://github.com/D-Jeffrey/MySewingNet-Cleanup"
    pause 'Press any key to Close'

    break
    }
if (!( test-path -Path $savetree)) {
    Write-Host "Can not find the Instruction Directory ($savetree).  Stopping" -BackgroundColor DarkRed -ForegroundColor White
    write-host "See instructions at   https://github.com/D-Jeffrey/MySewingNet-Cleanup"
    pause 'Press any key to Close'

    break
    }


Write-Host    ("Download source directory          : " + $downloaddir)
Write-host    ("mySewnet Cloud sub folder directory: " + $tree)
Write-host    ("Instrctions directory              : " + $savetree)
Write-host    ("File types                         : " + $mysewtype)
Write-host    ("Age of files in Download directory : " + $DownloadDaysOld)
Write-host    ("Clean the mysewingNet cloud folder : " + $CleansewingNet)
Write-Verbose ("Rollup match pattern               : " + $rollups)
Write-Verbose ("Ignore Terms Conditions files      : " + $TCs)
Write-Verbose ("Testing                            : " + $Testing)
Write-Verbose ("Excludetypes                       : " + $excludetypes)

$cont = (pause 'Press Start to continue, any other key to stop'  $true 'Click Yes to start') 

if (!$cont) { 
    Break
    }
Add-Type -assembly "system.io.compression.filesystem"

$treesizebefore = 0
Get-ChildItem -Path ($tree + "..")  -Recurse -file  | ForEach-Object { $treesizebefore = $treesizebefore + $_.Length}



# Get a list of all the existing files in mySewnet
$mysewingfiles = (Get-ChildItem -Path $tree  -Recurse -file -include $mysewtype).Name


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
                        if ($f -match $m) {
                            $isnew  = $false
                            Write-verbose "Duplicate file '${f}'."

                        }
                    }
                } 
            if ($isnew) {  
                Write-host "** New files in ZIP: '$zips'." 
                Expand-Archive -Path $zips -DestinationPath $tmpdir
                movefromdir $tmpdir $true $false
                movefromdir $tmpdir $true $true
                

                
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
        if (!(test-path -path (join-path -Path $tree -ChildPath $_.Name))) { 
            $_ | Copy-Item -Destination $tree -ErrorAction SilentlyContinue
            Write-Verbose "Copied from Download :'$_.Name}' to ${tree}."
            $addsizecnt = $addsizecnt + $_.Length
            $savecnt = $savecnt + 1 
            }
         }
            

if ($cleansewingNet) {
    Get-ChildItem -Path $tree  -Recurse -file -Exclude $mysewtype  | ForEach-Object {
     
            $sizecnt = $sizecnt + $_.Length
            Write-Verbose "Removing '$_'." 
            if ($doit) { deleteToRecycle  $_ }
            $filecnt = $filecnt + 1
            $p += 1
            write-progress -PercentComplete  ($p % 100 ) "Updating MySewnet"
        }
    
    movefromdir $tree, $true, $false
    }


#  Clear out empty Directories
$tailr = 0
while ($tailr -le 8 -and (tailRecursion $tree) ) {
     $tailr = $tailr + 1
     write-progress -PercentComplete  ($p % 100 ) "Clean directories"
}

$treesizeafter = 0
Get-ChildItem -Path ($tree  + "..") -Recurse -file  | ForEach-Object { $treesizeafter = $treesizeafter + $_.Length}
function nicesize ($sz) 
{
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
  
$treesizebefore = nicesize $treesizebefore
$treesizeafter = nicesize $treesizeafter
$kb = nicesize  $sizecnt
$akb = nicesize $addsizecnt
write-progress -PercentComplete  100  "Done"
Write-host "MySewnet Cloud size is now : $treesizeafter was $treesizebefore."
Write-Host ("Cleaned up - Dirs removed: '$dircnt' Save: '$savecnt' ($akb) Removed: '$filecnt' ($kb).") -ForegroundColor Green


$addsizecnt = $addsizecnt + $_.Length
pause 'Press any key to Close'
Write-Host ("Done ") -ForegroundColor Green


