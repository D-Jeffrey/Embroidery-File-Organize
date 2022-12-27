# Buggy software

*I can not be believe that that they ask people to pay to use this crappy service*
I have spent the last 8 hours rebuilding my wife MySewnet folders. After the combination of the web page and sync client wiped out 95% of the files.

So the Pfaff version of the client on my wife's machine works okay.  (It has lots of sync issues where it needs to time-out while syncing files, I assume that is caused by the number of session that it has running concurrently and not processing some REST error messages)

The mySewnet Windows sync client software is buggy.  As well, as the web pages are buggy.  I have no idea who tested this, but they certainly did not do any volume testing or a number of folders.
Examples of the bugs (if anyone from Mysewnet ever read this page):
- On the web page, when you select multiple items (using the checkbox by the Name) and you have several hundred items, it may indicate that you have selected the items, but not actually check the boxes.  It will not show you that it may have selected folders and then when you delete, it's list of selected items which not match what is checked on the page.
- in the client if you move folders, depending on how many files and folders you have the client will not realize that you have move folders or files and assume they are added, then it will bring a copy of that file back from the cloud onto your computer. This maybe be a good way to mess up the customer and cause them to need to purchase more space in order to keep the unexpected growth under control.
- the sync client gets confused and started creating folders when it is looking at a file and then it gets into a loop repeatly trying to sync that folder.  One way to clear that is to quit the sync program and start it again.
- it has major issues with nested folders, if you create nested folder by copying a group of folders around, then the client may end up creating a flatten structure copy.
- several times, I have had all files disappear or get duplicated, which means keeping all the ZIP files sitting around, in case you need to reload them.
- I guess the intent is to make you purchase from them, rather than 3rd parties they can put the files back into your cloud after it gets all messed up.

While I am on the **ranting throne**, have you seen the junk they call **[QCT](https://graceframe.com/en/product/quiltmotion-qct5)** for longarm sewing machines?  They charge thousands for a bunch of Visual Basic 6 programs (which is **[EOL](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/visual-basic-6-support-policy#the-visual-basic-60-support-lifetime)** since April 8, 2008) which crash regularly, has crappy keyboard overlap which gets in the way all the time, a useless file naviation replacement for openfile, and a workflow process which in great for influences, cause without hours and **hours* of watching videos, there is no way you would figure it out.
