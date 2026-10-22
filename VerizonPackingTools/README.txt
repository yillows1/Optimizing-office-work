things to remember to make:

how the flow of the work works now with these tools
question.  how does this work or does it even work with one that doesnt need splits

Answer: It should now work with the ones the don't need splits so 4RR/4RP and etc still will function as intended


DISCLAIMER: 
Still haven't implemented anything that stores to the website database



Basic use:
1) Boot up PackingList.exe
2) Pick folder where the application will read the file from
3) Hit open input sheet to fill out a page, or you can just drag and drop the file in but that's probably not implemented yet so just hit open for now
4) fill out the sheet and save
5) Hit rename files for helping yourself keep track (the app will rename it based on the FRN or Packing List Number) - if you messed up, the program looks for the Incoming FRN at the beginning I think so if you want it to
rename the excel sheet, change the name to something completely wrong and it will pick up on it correctly.  Otherwise if it starts with FRN it assumes its correct since checking every excel sheet would be slow

(FYI the steps down here may be a bit slower as they read and scan or write on the excel sheets so wait till it says imported or if something pops up)
6) Hit Import Input Sheet and select the one that you want to Import 
7) hit scan item list
8) choose if the NEAR MISSES are close enough.  
(personally, I would say if its a number too short and its otherwise exactly the same I consider that the same, and if a 1 and 7 get messed up or like 1 and a 2 or most likely a 3 and an 8 will be the same)
9) If you have any RITM or Addresses that should be either merged into one of the other loads or into a new load, input them in the popup that appears now
10) now once ur on this screen where both the bottom boxes are filled with a few things, then you can double check and if you don't like the way it's split or you need to wait for more information, 
this is the part where you should bail before things get written. 
If you are good with this, continue

(BTW Steps 11 and 12 are interchangable you can do it in either order)

11) you can now click on any of the emails (the list on the middle of the screen that stars with the RITM then Sender etc) and click "Reply to Selected Email".  this will automatically prep the reply email system
which will set up the reply with all the items found in that load
12) Once you send all your emails, you can hit "Export Folder..." to select a folder to put all the excel and pdfs and folders in.  Basically, the App will create these things for you:
- Packing List (per load so A,B,C all will be covered.. ignore this for ones that only have 1 load it still makes it anyways)
- Summaries Folder which contains
    - Reel Summary Report + Date Range (if your packing list falls in the range of an already existing summary, it adds into it instead of creating a new one)
- FRN ... Folder which contains
    - PDF from email with "Delivery FRN..."
    - Scrap Pick Up Request for Load FRN...

13) [OPTIONAL] Once you're done with printing out the files and moving them to the scanned docs folder, you can delete everything EXCEPT Summaries and the Summaries inside of it.  
Those files will be needed for the SplitMaterial.exe


When the total Weights for each material comes in, you can use SplitMaterial.exe


Things to remember:
After doing split materials, the summary file's tabs aren't always in order since sometimes they will be in the order you made them since I kept breaking the program when I tried to get them to be in the correct order (FIXED for R, inconsistent for P)
Although for reference it should order it correctly on the summary tab but just the actual physical P tabs can end up out of order
but this only happens if you input data out of order.  This does not happen to the R ones though I figured out a fix idk why it doesnt work as consistently in P tabs but that may be because
the split_material_app grabs the first tabs and then changes the first tabs if you change the weights in any way and then writes the other tabs and what not so it probably gets messed up because of that

the little checkbox for sorting RP families should work for the loads that literally just have 1 load so like FRN261009-3R.  No a,b,or c at the end and things like that.  This should still work
even if you don't click the checkmark just it was added a while back and i never removed it
so usually the recommended is to leave this box UNCHECKED


ALSO
when using split apps, it capitalizes FRN and whatever based on how you type it in for the summary tab





Not super sure on some things since I tested everything a bunch of times but still mightve missed things so if you find any issues lmk


random things:

if there isnt a good solution it tries to basically create its own solution by doing all the loads except 1 so say you have loads A, B, and C.  A and B are given values that make sense and work but C is then
just given the remainder which will look wrong if you do the math but this ONLY happens if the numbers dont match up in general like theres too much fo99 for it to work or something like that

split material tries to do it correctly, but if thee is no correct answers and the numbers of each load actually add up, then it just kinda does the first loads and the last load gets the rest of the Weights

