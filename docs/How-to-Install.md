# How to Install

~~If you have MySewnet Cloud~~

1. First download the script (select the green Code button and choose Download ZIP).  Open the Zip and save the file `EmbroderyCollection-Cleanup.ps1` it to a a folder like **Documents** `scripts`
2. Select the file and select the Properties click on `unblock`
3. Put the file in a directory such as 
  - For the location of the item put in something like `Embroidery` inside of the folder `Documents`
  - create a short cut to the script with a command such as 
  `Powershell.exe -NoLogo -ExecutionPolicy Bypass -File "%USERPROFILE%\Documents\EmbroderyCollection-Cleanup.ps1"`
  - For the name use something like` Embrodery Extractor`
3. Create a directory in **Documents** call `Embroidery Instructions`
4. If you have a lot of duplicate and extract files in your Mysewnet Cloud folder then create a folder in **Downloads** and move all the files into that folder.
  - Don't worry will put all the right files back in there for you
  
~~If you have Mysewnet Cloud, create a directory in there called `Embroidery`~~

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
