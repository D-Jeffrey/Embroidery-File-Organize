## Organize your embroidery pattern files, you have too many copies of the same pattern and you can’t find the right one…. 
I created this for my wife in order to avoid having to buy a MySewnet subscription. Additionally, when you end up with too many variations of patterns on your sewing machine it's really hard to find anything. In the case of my wife, she has over 2600 different patterns on their machine and still only uses 67 megabytes of space and she can find the pattern that she is looking for.  The other problem I found is a lot of the patterns come with a PDF or images which you don't need to put into your sewing machine or into my sewing net.  My annoyance with the service started when we could not get a free trial in Canada which was supposed to be include with the Pfaff Icon that my wife bought.  There are benefits to the subscription service and nice looking patterns, but over $300 per year is a little steep.  Regardless even with a subscription you still end up with a message of folders and files and hard to fine the right pattern type connection.  This is here to HELP manage and organize those files.

The program is designed to make sure that you don't duplicate patterns either by putting more than one copy of the same pattern or by putting more than one copy of pattern variation such as a different file type.  It also allows you to move files, around put them into folders and will respect where they've landed and not resort them again.

So let's take a look at the problem if you look at it as a purchased which you purchased from one of the many excellent pattern providers out there.  Some of them come with all the files dumped into the zip file with no directory structure.  Some of the zip files come with a directory structure which may be two or three directories deep sorted by the various file types. Other ones may have large PDFs handsome bitmap pictures of what the pattern is that you're making.  To do is only pluck out the right patterns that you need to be loaded onto your sewing machine and the instructions that go along with them and put them into different places on your computer away from the Mysewnet cloud sync client. Then when you download some new zip files the program should I might could check those new ones find out what's missing and push them into the right directories as well.

Let’s take a look at an example of how much of a mess this will create.   This is the mess you get from extracting all the files into one directory with whatever folder structure it comes with.  (SpaceSniifer)

There was a 'Nursery Rhymes 'pattern folder itself is 184Mb by itself.  And I really only need the VP3 files.

So I have almost 1300 folders with 20,000 files using 1.5 GB of space.  Keep in mind that some of these might be pattern for the Cricut because those and can mixed up the folds.
Rather than extracting all those zips into a mess, which by the way the top level directory there are 1057 files taking up 140 Mb.

Let’s assume you just move that mess over into “Embroidery” directory within MySewnet and run 
```
C:\Users\kjeff\Documents\scripts\EmbroderyCollection-Cleanup.ps1 -CleanCollection 
```

First pass, it came back with a huge list of files that it want to delete 

```
Found 11080 files that should be removed to clean up extras 

*** MySewnet Cloud size is now : 166.8 MB was 257.5 MB ****  
```

```
PS c:\Users\kjeff\Downloads> C:\Users\kjeff\Documents\scripts\EmbroderyCollection-Cleanup.ps1.ps1 -CleanCollectionIgnoreDir 
                Begin MySewnet Process                             
                Checking for Zips in the last 7 days               
Download source directory          : d:\Users\kjeff\downloads
mySewnet Cloud sub folder directory: d:\Users\kjeff\mySewnet Cloud\Embroidery\
Instructions directory             : d:\Users\kjeff\OneDrive\Documents\Embroidery Instructions\
File types                         : *.vp4 *.vp3 *.vip *.pcs *.dst
Age of files in Download directory : 7
Clean the mysewingNet cloud folder : False
Keep all variations of files types : False
Scanning for files to clean up in mySewnet
Found 15540 files that should be removed to clean up extras
Switching to Fast quick delete without recycle
Moving Instructions to proper Instructions directory
Calculating size
   *** MySewnet Cloud size is now : 281.4 MB was 1.6 GB ****   
Cleaned up - Dirs removed: '0' Saved from zip files: '0' (0 KB) Removed: '15540' (578.7 MB).
End 
```

What happens if we just let it run as it normally should be checking all the downloaded files in the download directory and move them into the proper place.  It only takes 64 MB of space instead of the orginal 1.6 GB we started with.

```
PS D:\Users\kjeff\Downloads> C:\Users\darre\OneDrive\Documents\scripts\EmbroderyCollection-Cleanup.ps1 -Testing 
                Begin Embroidery Filing                                     
                Checking for Zips in the last 7 days                        
Download source directory           : d:\Users\kjeff\downloads
MySewnet Cloud sub folder directory : D:\Users\kjeff\mySewnet Cloud\Embroidery\
Instructions directory              : d:\Users\kjeff\OneDrive\Documents\Embroidery Instructions\
File types                          : *.vp4 *.vp3 *.vip *.pcs *.dst
Age of files in Download directory  : 7
Clean the Collection folder         : False
Keep all variations of files types  : False
Testing Mode                        : True
* New  : '...\12644-907.pcs.zip'                                   2 new patterns
* New  : '...\12644-907.vp3 (1).zip'                               3 new patterns
- Found: '...\12644-907.vp3 (2).zip'                               nothing new
- Found: '...\12644-907.vp3.zip'                                   nothing new
* New  : '...\17445225.zip'                                        40 new patterns
* New  : '...\17489954 (1).zip'                                    29 new patterns
- Found: '...\17489954 (2).zip'                                    nothing new
- Found: '...\17489954.zip'                                        nothing new
* New  : '...\17529881.zip'                                        36 new patterns
* New  : '...\17532767.zip'                                        8 new patterns
* New  : '...\17561335.zip'                                        15 new patterns
* New  : '...\17589307.zip'                                        25 new patterns
* New  : '...\17660766.zip'                                        26 new patterns
* New  : '...\17661990 (1).zip'                                    18 new patterns
- Found: '...\17661990.zip'                                        nothing new
* New  : '...\17700580 (1).zip'                                    7 new patterns
- Found: '...\17700580 (2).zip'                                    nothing new
- Found: '...\17700580.zip'                                        nothing new
* New  : '...\17710087.zip'                                        15 new patterns
* New  : '...\17838253.zip'                                        13 new patterns
* New  : '...\17841716.zip'                                        7 new patterns
* New  : '...\17872666.zip'                                        74 new patterns
* New  : '...\17912524.zip'                                        66 new patterns
* New  : '...\17942480.zip'                                        27 new patterns
* New  : '...\17942481.zip'                                        4 new patterns
* New  : '...\17960951.zip'                                        38 new patterns
* New  : '...\17973434.zip'                                        24 new patterns
* New  : '...\17981105.zip'                                        33 new patterns
* New  : '...\17986490.zip'                                        13 new patterns
* New  : '...\17992030.zip'                                        5 new patterns
* New  : '...\18096530.zip'                                        22 new patterns
* New  : '...\18130241.zip'                                        28 new patterns
* New  : '...\18147942.zip'                                        21 new patterns
* New  : '...\18187327 (1).zip'                                    36 new patterns
- Found: '...\18187327.zip'                                        nothing new
* New  : '...\18198823.zip'                                        3 new patterns
- Found: '...\3d-layered-hummingbird-svg-mandala-jennifermaker.zip'  nothing new
- Found: '...\3d-paper-birds-jennifermaker.zip'                    nothing new
* New  : '...\61062.dst (1).zip'                                   28 new patterns
- Found: '...\61062.dst.zip'                                       nothing new
- Found: '...\61062.hus (1).zip'                                   nothing new
- Found: '...\61062.hus.zip'                                       nothing new
* New  : '...\61062.pcs (1).zip'                                   16 new patterns
- Found: '...\61062.pcs.zip'                                       nothing new
- Found: '...\61062.pes.zip'                                       nothing new
- Found: '...\61062.sew.zip'                                       nothing new
- Found: '...\61062.xxx.zip'                                       nothing new
* New  : '...\ae 3d christmas gift tags.zip'                       16 new patterns
* New  : '...\ae 3d ornaments.zip'                                 24 new patterns
* New  : '...\ae hand stitched christmas pillows.zip'              30 new patterns
* New  : '...\angels.zip'                                          95 new patterns
* New  : '...\ap2932a_vp3.zip'                                     10 new patterns
* New  : '...\ap3029a_vp3.zip'                                     10 new patterns
- Found: '...\art-ithbellapuppyset.zip'                            nothing new
- Found: '...\art-ithkittyset.zip'                                 nothing new
* New  : '...\away in a manger.zip'                                102 new patterns
- Found: '...\beautifulthings (1).zip'                             nothing new
- Found: '...\beautifulthings.zip'                                 nothing new
- Found: '...\beyond-diet-tools.zip'                               nothing new
- Found: '...\cc06222.zip'                                         nothing new
* New  : '...\christmas bible bookmarks.zip'                       10 new patterns
- Found: '...\color lesson practice doc (1).zip'                   nothing new
- Found: '...\color lesson practice doc (2).zip'                   nothing new
- Found: '...\color lesson practice doc (3).zip'                   nothing new
- Found: '...\color lesson practice doc (4).zip'                   nothing new
- Found: '...\color lesson practice doc (5).zip'                   nothing new
- Found: '...\color lesson practice doc (6).zip'                   nothing new
- Found: '...\color lesson practice doc.zip'                       nothing new
- Found: '...\disney love songs.zip'                               nothing new
- Found: '...\diy-white-board-calendar-jennifermaker.zip'          nothing new
- Found: '...\drawer-dividers-noglue-jennifermaker (1).zip'        nothing new
- Found: '...\drawer-dividers-noglue-jennifermaker.zip'            nothing new
* New  : '...\dst-ithbellapuppyset.zip'                            49 new patterns
* New  : '...\dst-ithkittyset.zip'                                 59 new patterns
- Found: '...\easter-cross-card.zip'                               nothing new
- Found: '...\easy summer dress.zip'                               nothing new
- Found: '...\easy-heart-box-jennifermaker.zip'                    nothing new
- Found: '...\edp13891-2.dst.zip'                                  nothing new
- Found: '...\edp13891-2.vp3.zip'                                  nothing new
- Found: '...\engraved-ornaments-jennifermaker.zip'                nothing new
- Found: '...\envelope2bybird.zip'                                 nothing new
* New  : '...\esp80149-1.vp3.zip'                                  1 new patterns
- Found: '...\exp-ithbellapuppyset.zip'                            nothing new
- Found: '...\exp-ithkittyset.zip'                                 nothing new
- Found: '...\faux-leather-earrings-jennifermaker.zip'             nothing new
- Found: '...\filigree-ornament.zip'                               nothing new
* New  : '...\fundamentals-designs.zip'                            94 new patterns
- Found: '...\giant-paper-poinsettia-jennifermaker (1).zip'        nothing new
- Found: '...\giant-paper-poinsettia-jennifermaker (2).zip'        nothing new
- Found: '...\giant-paper-poinsettia-jennifermaker.zip'            nothing new
- Found: '...\grab-and-go-wristlet-8211-svg-8211-for-die-cutting-machines-00504913900.zip'  nothing new
* New  : '...\hand stitched autumn quilt.zip'                      160 new patterns
- Found: '...\hand-drawn-spring-flower-set.zip'                    nothing new
- Found: '...\handsanitzer-pdf-svg-patterns-00500080500.zip'       nothing new
- Found: '...\heart-box-jennifermaker.zip'                         nothing new
- Found: '...\heart-explosion-box-jennifermaker (1).zip'           nothing new
- Found: '...\heart-explosion-box-jennifermaker.zip'               nothing new
- Found: '...\heart-lantern-lg-jennifermaker.zip'                  nothing new
- Found: '...\heart-rainbow-svg-v2_6c773b90-0aeb-46f3-ab46-87af3a2fbc95 (1).zip'  nothing new
- Found: '...\heart-rainbow-svg-v2_6c773b90-0aeb-46f3-ab46-87af3a2fbc95.zip'  nothing new
- Found: '...\holiday-ornaments-stl-files-3d-prints.zip'           nothing new
- Found: '...\hus-ithbellapuppyset.zip'                            nothing new
- Found: '...\hus-ithkittyset.zip'                                 nothing new
- Found: '...\images-ithbellapuppyset.zip'                         nothing new
- Found: '...\images-ithkittyset.zip'                              nothing new
* New  : '...\inspirograph quilt (1).zip'                          100 new patterns
- Found: '...\inspirograph quilt.zip'                              nothing new
- Found: '...\jef-ithbellapuppyset.zip'                            nothing new
- Found: '...\jef-ithkittyset.zip'                                 nothing new
* New  : '...\journaling quilt.zip'                                72 new patterns
- Found: '...\k2442.vp3.zip'                                       nothing new
- Found: '...\l9449.vp3.zip'                                       nothing new
- Found: '...\m18051.vp3.zip'                                      nothing new
- Found: '...\m18612.vp3.zip'                                      nothing new
- Found: '...\m18615.vp3.zip'                                      nothing new
- Found: '...\m18618.vp3.zip'                                      nothing new
- Found: '...\m20713.vp3.zip'                                      nothing new
- Found: '...\m28367.dst.zip'                                      nothing new
- Found: '...\m28367.vp3.zip'                                      nothing new
- Found: '...\m28370.dst.zip'                                      nothing new
- Found: '...\m28370.vp3.zip'                                      nothing new
- Found: '...\m28373.vp3.zip'                                      nothing new
- Found: '...\m28715.vp3.zip'                                      nothing new
- Found: '...\m33042.dst.zip'                                      nothing new
- Found: '...\m33042.vp3.zip'                                      nothing new
- Found: '...\m33948.vp3.zip'                                      nothing new
- Found: '...\m5772.vp3.zip'                                       nothing new
- Found: '...\mf_i_love_glitter.zip'                               nothing new
- Found: '...\msg2pst.zip'                                         nothing new
- Found: '...\mysewnetcloudsyncsetup.zip'                          nothing new
* New  : '...\nursery rhymes bonus.zip'                            60 new patterns
* New  : '...\nursery rhymes.zip'                                  190 new patterns
- Found: '...\nymphette.zip'                                       nothing new
- Found: '...\onedrive-2018-01-28 (1).zip'                         nothing new
- Found: '...\onedrive-2018-01-28 (2).zip'                         nothing new
- Found: '...\onedrive-2018-01-28.zip'                             nothing new
- Found: '...\onedrive-2019-11-02.zip'                             nothing new
* New  : '...\pa-2021christmasornament-rudolph (1).zip'            2 new patterns
- Found: '...\pa-2021christmasornament-rudolph (2).zip'            nothing new
- Found: '...\pa-2021christmasornament-rudolph.zip'                nothing new
* New  : '...\pa-doyouwanttobuildasnowman-dolljointarms.zip'       5 new patterns
* New  : '...\pa-doyouwanttobuildasnowman.zip'                     40 new patterns
* New  : '...\pa-inthehoop-darlingdolls-dst.zip'                   176 new patterns
- Found: '...\pa-inthehoop-darlingdolls-exp.zip'                   nothing new
- Found: '...\pa-inthehoop-darlingdolls-hus.zip'                   nothing new
- Found: '...\pa-inthehoop-darlingdolls-jef.zip'                   nothing new
- Found: '...\pa-inthehoop-darlingdolls-sew.zip'                   nothing new
- Found: '...\pa-inthehoop-darlingdolls-vip.zip'                   nothing new
- Found: '...\pa-inthehoop-darlingdolls-vp3 (1).zip'               nothing new
- Found: '...\pa-inthehoop-darlingdolls-vp3.zip'                   nothing new
- Found: '...\pa-inthehoop-darlingdolls-xxx.zip'                   nothing new
* New  : '...\pa-inthehoopbabyblitzen.zip'                         77 new patterns
* New  : '...\pa-inthehoopdarlingbunny-dst.zip'                    83 new patterns
- Found: '...\pa-inthehoopdarlingbunny-exp.zip'                    nothing new
- Found: '...\pa-inthehoopdarlingbunny-hus.zip'                    nothing new
- Found: '...\pa-inthehoopdarlingbunny-jef.zip'                    nothing new
- Found: '...\pa-inthehoopdarlingbunny-pes.zip'                    nothing new
- Found: '...\pa-inthehoopdarlingbunny-sew.zip'                    nothing new
- Found: '...\pa-inthehoopdarlingbunny-vip.zip'                    nothing new
- Found: '...\pa-inthehoopdarlingbunny-vp3.zip'                    nothing new
- Found: '...\pa-inthehoopdarlingbunny-xxx.zip'                    nothing new
* New  : '...\pa-inthehoopreindeerstable.zip'                      76 new patterns
* New  : '...\pa-inthehoopsantaslittlereindeer.zip'                135 new patterns
* New  : '...\pa-inthehoopsillysnowballfight.zip'                  18 new patterns
* New  : '...\pa-ithabbydollandpony (1).zip'                       82 new patterns
- Found: '...\pa-ithabbydollandpony (2).zip'                       nothing new
- Found: '...\pa-ithabbydollandpony.zip'                           nothing new
* New  : '...\pa-sallyfaceaddon-darlingdolls.zip'                  5 new patterns
* New  : '...\pa-taylorfaceaddon-darlingdolls.zip'                 5 new patterns
* New  : '...\pa-unicornhornforpony (1).zip'                       4 new patterns
- Found: '...\pa-unicornhornforpony.zip'                           nothing new
* New  : '...\pa-willafaceaddon-darlingdolls.zip'                  5 new patterns
- Found: '...\paper-poppies.zip'                                   nothing new
- Found: '...\paper-star-lanterns-jennifermaker.zip'               nothing new
- Found: '...\patch-8x12ponybody-october2021 (1).zip'              nothing new
- Found: '...\patch-8x12ponybody-october2021.zip'                  nothing new
- Found: '...\patch-ithkittyset.zip'                               nothing new
* New  : '...\pa_christmasfingerpuppets.zip'                       4 new patterns
- Found: '...\pes-ithbellapuppyset.zip'                            nothing new
- Found: '...\pes-ithkittyset.zip'                                 nothing new
- Found: '...\quiltlabelsfromsewcanshe_aiid1428726.zip'            nothing new
- Found: '...\sew-ithbellapuppyset.zip'                            nothing new
- Found: '...\sew-ithkittyset.zip'                                 nothing new
* New  : '...\sewing machine patchwork cover.zip'                  30 new patterns
- Found: '...\sewitpretty-2.zip'                                   nothing new
- Found: '...\sewmuchfabric.zip'                                   nothing new
* New  : '...\she5467a_vp3.zip'                                    10 new patterns
- Found: '...\snowflake-tree-basic-jennifermaker.zip'              nothing new
* New  : '...\snowflakes.zip'                                      2 new patterns
- Found: '...\snowman-face-hat-jennifermaker.zip'                  nothing new
* New  : '...\spring flower doodles.zip'                           30 new patterns
- Found: '...\spring_romance (1).zip'                              nothing new
- Found: '...\spring_romance.zip'                                  nothing new
* New  : '...\stippled quilt block.zip'                            7 new patterns
- Found: '...\summer-flower-edge-card-by-bird.zip'                 nothing new
* New  : '...\sunflower crazy quilt.zip'                           50 new patterns
* New  : '...\sunny sunflower.zip'                                 11 new patterns
- Found: '...\text on path practice doc.zip'                       nothing new
- Found: '...\thread-catcher-svg-pattern-00539650700.zip'          nothing new
- Found: '...\tool-holder-jennifermaker-stl.zip'                   nothing new
- Found: '...\towel-hanger-pdf-pattern-00510022900.zip'            nothing new
* New  : '...\utz2377.vp3.zip'                                     2 new patterns
- Found: '...\vip-ithbellapuppyset.zip'                            nothing new
- Found: '...\vip-ithkittyset.zip'                                 nothing new
- Found: '...\vp3-ithbellapuppyset.zip'                            nothing new
- Found: '...\vp3-ithkittyset.zip'                                 nothing new
- Found: '...\winter-wreath-jennifermaker.zip'                     nothing new
- Found: '...\x13113.vp3.zip'                                      nothing new
- Found: '...\x13237.vp3.zip'                                      nothing new
- Found: '...\x13734.vp3.zip'                                      nothing new
- Found: '...\x13774.dst.zip'                                      nothing new
- Found: '...\x13774.vp3.zip'                                      nothing new
- Found: '...\x14789.vp3.zip'                                      nothing new
- Found: '...\x14808.vp3.zip'                                      nothing new
- Found: '...\x14959.vp3.zip'                                      nothing new
- Found: '...\x14990.vp3.zip'                                      nothing new
- Found: '...\x15445.dst.zip'                                      nothing new
- Found: '...\x15445.vp3 (1).zip'                                  nothing new
- Found: '...\x15445.vp3.zip'                                      nothing new
- Found: '...\x15561.vp3.zip'                                      nothing new
- Found: '...\x3686.vp3.zip'                                       nothing new
- Found: '...\xxx-ithbellapuppyset.zip'                            nothing new
- Found: '...\xxx-ithkittyset.zip'                                 nothing new
Calculating size
Added files to MySewnet Cloud: '2639' (64.6 MB) 
   *** MySewnet Cloud size is now : 64.6 MB was 0 B ****   
End



