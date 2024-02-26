# How to Install

~~If you have MySewnet Cloud~~

1. First download the script go to https://github.com/D-Jeffrey/Embroidery-File-Organize/blob/main/EmbroideryCollection-Cleanup.ps1 and click on the Download icon (just below History on the right side of the page).  Move to the file `EmbroderyCollection-Cleanup.ps1` it to a a folder like **`Documents\scripts`**
2. Select the file and select the Properties click on `unblock`
3. Put the file in a directory such as 
  - For the location of the item put in something like `Embroidery Instructions` inside of the folder `Documents`
4. Open a command prompt (Start -> Run -> `CMD`)
    cd to the directory of the 
  - create a short cut to the script with a command such as 
```
Powershell -ex bypass -command ".\EmbroderyCollection-Cleanup.ps1 -setup"
```
  - For the name use something like` Embrodery Extractor`
3. Create a directory in **Documents** call `Embroidery Instructions`
4. If you have a lot of duplicate and extract files in your Mysewnet Cloud folder then create a folder in **Downloads** and move all the files into that folder.
  - Don't worry will put all the right files back in there for you
  
5. Create a folder in **Documents** call `Embroidery`


6. (Optional) Edit the PS1 file and set the types of patterns that work best for you.
      `$preferredSewType = ('vp3', 'vp4')`
   - See [File Types](File-Types.md)

You should be ready to go now... run the program by clicking on the shortcut and watch the results..

Or run Powershell and at the command prompt using the commandline options such as:
```
PS>  cd C:\Users\darre\Documents\Embroidery
PS>  EmbroderyCollection-Cleanup.ps1 -Testing
PS>  EmbroderyCollection-Cleanup.ps1 -?
PS>  EmbroderyCollection-Cleanup.ps1 -DownloadDaysOld 720

```
7. I also suggest you download and install **[MySewnet Embroidery Software](https://download.mysewnet.com/)** & **[Explorer Plug-in Software](https://download.mysewnet.com/)** from Mysewnet, their thumbnail file preview feature is very well done.
![explorer with preview](images/2022-12-27_10-56-25.gif)







$ScriptFromGitHub = Invoke-WebRequest https://raw.githubusercontent.com/D-Jeffrey/Embroidery-File-Organize/main/install.ps1
Invoke-Expression $($ScriptFromGitHub.Content)