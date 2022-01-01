# MySewingNet-Cleanup
A powershell script to deal with the many different types of embroidery files, put the right format types in [mySewingnet](https://mysewnet.com/) Cloud

## parameters
  >-includeEmbFilesWithInstruction = $false
  >-CleansewingNet = $false     
  >-DownloadDaysOld = 7   

  - -includeEmbFilesWithInstruction = $false
     Put a copy of the Embrodery in with the instructions in addition to putting them into the mysewnet cloud folders
  - -CleansewingNet = $false     
     Cleanup the MySewingNet Folder to only Embrodery files
  - -DownloadDaysOld = 7   

## functions

Checks the download directory for the Embrodery files types of any age and all the zip files which have been downloaded in *DownloadDaysOld*.  
Any Embrodery files found are copied in to the mysewing cloud folder under *treetop* directory (set below).
Any zip files found are scanned to see if they have Embrodery file types that we are interested in.  If they are files which do not yet exist in the
mysewingnet cloud, then extract that zip to a temporary location, pull out all the relevant files (formats we want) with the directory hierachy (adjusted).  Also pull out any related documentation and put it into the *instructions* folder location within the user documents on the computer with the directory hierachy (adjusted).
**TODO** add a get other types function

## directory hierachy (adjusted)
When vendors build zip files and put in all the different formats, it means digging for files.  the adjusted version of this will get rid of sub folders if they exist above 
## requirements

It was designed to work with [mySewnet Cloud](https://cloud.mysewnet.com/) which is a type of file share service for sewing machines.

$treetop = "Embroidery"
$instructions = "Embroidery Instructions"

depending on the types Embrodery files your machine uses then you may what to change the sewing file types of 
$mysewtypes = ('vp3', 'vp4')


### nice to know

$TCs = @('TERMS-OF-USAGE.*', 'planetappliquetermsandconditions.*')
$rollup = @('images','sewing helps')
sewtypes rollup

![powershell running](docs/2022-01-01_13-15-43.gif)
