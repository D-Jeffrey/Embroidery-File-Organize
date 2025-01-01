# General Tips
## Pin the Embroidery folder Quick Access
If you are looking for your files on a regular basis, then consider using the PIN to Quick Access function in Windows Explorer. [Pin, remove, and customize in Quick access](https://support.microsoft.com/en-us/windows/pin-remove-and-customize-in-quick-access-7344ff13-bdf4-9f40-7f76-0b1092d2495b)

## Restore build Windows Zip
If you have installed a 3rd party bloatwear zip package, know what Windows has a nice native function.  To restore the native Zip file integration in Windows, you only have to do this from an elevated (run as administrator) command prompt:

Click the `Windows Start` > type 'cmd' and select the  "Run as Administrator" option to option the command prompt.  Type :
```
cmd /c assoc .zip=CompressedFolder
```
and press Enter.
You may have to restart the computer.

## Upgrade your Powershell to Version 7
The performance is better with Powershell 7

[Installing Powershell version 7](https://learn.microsoft.com/en-us/powershell/scripting/install/installing-powershell-on-windows)
--or-- 
[https://aka.ms/PSWindows](https://aka.ms/PSWindows)

### Nice to know

- If you have `7zip` installed on your computer, (the native version, not MS store version), then it will be used when working with large zip files as it is much faster.

- It will create folders when there is a number of files `-SetSize` that are in a given zip file using the name of the zip file.  You most likely will want to rename it and give it a new name which reflects the folder.

- Ignore files which are terms and conditions (it does not mean you can ignore the laws, just don't save so many copies of the files.
`$TandCs = @('TERMS-OF-USAGE.*', 'planetappliquetermsandconditions.*')`  Edit your config (`EmbroideryCollection.cfg`) file to adjust these values. 
This are the directories (plus the if the directory name equals the format type)  
`$foldupDir = @('images','sewing helps','Designs', 'Design Files')` Edit your config file.
