i think i need to make the app first before this since all the instructions should be made after the app is created

things to remember to make:

how the flow of the work works now with these tools
question.  how does this work or does it even work with one that doesnt need splits

Answer: It should now work with the ones the don't need splits so 4RR/4RP and etc still will function as intended



Basic use:
1) Boot up PackingList.exe
2) Pick folder where the application will read the file from
3) Hit open input sheet to fill out a page, or you can just drag and drop the file in but that's probably not implemented yet so just hit open for now
4) fill out the sheet and save
5) Hit rename files for helping yourself keep track (the app will rename it based on the FRN or Packing List Number)

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

