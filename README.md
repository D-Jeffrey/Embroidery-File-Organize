# MySewingNet-Cleanup
A powershell script to deal with the many different types of embroidery files, put the right format types in [mySewingnet](https://mysewnet.com/) Cloud

## parameters
`-includeEmbFilesWithInstruction = $false` Put a copy of the Embrodery in with the instructions in addition to putting them into the mysewnet cloud folders
`-CleansewingNet = $false`  Clean out non embroidery files from the mysewingnet cloud directory since it is limited on the amount of space you have to work with (unless you are using the Silver or Platinum plans.  The files are deleted to the **recycle bin** so they can be restored.
`-DownloadDaysOld = 7`  determine how old of zip files to look for (in days) 

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

`$treetop = "Embroidery"` is the directory name within your mysewingnet cloud that this program will put all the files and clean our file types that do not match the type you set.  It must exist in in the root directory of the **mySewingnet cloud folders**, order for the program to run.
`$instructions = "Embroidery Instructions"`  this is where all the instructions are saved (outside of mysewingnet).  It must exist within the users **Documents** folder in order for the program to run

Depending on the types Embrodery file extensions your machine uses then you may what to change the sewing file types of for you machine.
`$mysewtypes = ('vp3', 'vp4')`


### nice to know

Ignore files which are terms and conditions (it does not mean you can ignore the laws, just don't save so many copies of the files.
`$TCs = @('TERMS-OF-USAGE.*', 'planetappliquetermsandconditions.*')`
This are the directories (plus the if the directory name equals the format type)
`$rollup = @('images','sewing helps')`


![powershell running](docs/2022-01-01_13-53-31.gif)
