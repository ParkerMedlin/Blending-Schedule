# Update Notes

### 3-21-2022
 - Rebuilt the whole spreadsheet on top of a new query which pulls directly from sharepoint instead of using onedrive

### 2-28-2022
 - Added individual refresh buttons for every table bc refreshAll is stupid

### 2-22-2022
 - Updated refresh sub so that it no longer depends on MacrosOn, but that still didn't stop the timestamp sub from triggering during refresh
 - Updated the countPrint sub so it manually fits the print job all columns to one page

### 2-8-2022
 - Added a module so the workbook_open and workbook_close actions would actually work. Select ChemShortages when the workbook is closed or opened.
 - Got rid of MacrosOn/MacrosOff cells. Moved those subs to power button icon.  

### 2-7-2022
 - Standardized the count sheet print sub so that it works on both chemCounts and blendCounts
 - Moved more of the icons over to the chemCounts sheet w/ supporting documentation
 - Removed unnecessary pagesetup formatting in the blendThese print sub
 - Added an indicator to blendCounts that shows whether the blend is on BlendThese (shortage that needs to be blended)
 - Standardized the clear filter macro to work on all tables
 - Standardized clearAndReturn so it just grabs the destination from the active cell
 - Added the lookup macro mapping to the cells on ChemCounts containing chem part numbers, like in chemShortage
 - fixed daily report extra row being copied over

### 2-2-2022
 - Continued work on the mrp-ish report and added a couple others and put them all in a userform like BlendReportSelekta
 - Simplified clearAndReturn subs, removed a few, also removed the printer sub for the pickSheetTable

### 2-1-2022 
 - Finished the chem report sub (called 'realSpecificHours' under the Reports module)
 - Added a column to the bom.master table for weight per gal from Sage
 - Added a sub to ClearnFilter module for wiping the realSpecificHours table
 - Removed pickSheetTable

### 1-31-2022
 - Made countLog sub better by changing all the individual autofilter clears to ListObject.Autofilter.ShowAllData
 - Got started on a report that will show us the usage of a given chemical over the course of all upcoming runs in the schedule for all the blends that use that chem

### 1-27-2022
 - Added 2 tables for chem counts and logging those counts
 - updated the countLog sub so that it now has variables that can be set depending which sheet is active (blendCounts vs chemCounts)

### 1-26-2022
 - Select qty cell after the lot num gen macro 

### 1-18-2022
 - cleared out some unnecessary Private Subs

### 1-13-2022
 - Changed MacrosOff so it just clears the cell instead of inserting "Off"

### 1-12-2022
 - Reverted to the prodschedule version of timeTable_query because blendThese was being weird.

### 1-11-2022
 - Took out all the sheet-keyed lookup macros and put them into a userform called ReportSelekta. This userform triggers when you click the blend part number on most sheets. 
 - Rebuilt the checkoutcounts query. It's now called BlendCounts and it's much simpler.
 - Fixed and removed various small things all around the sheets and the modules.  
 - Added text box documentation to the macro-linked icons on blendthese and blendcounts.  

### 1-5-2022
 - Blend These missed a shortage after my "fixing" so I reverted again and im leaving it at the 15 god awful fucked up redundant steps that it has. I'm sure I'll be back for more punishment once I forget about this. 

### 1-4-2022 
 - Happy New Year! the Blend These table was broken this morning! opened up previous version and restored the query and then walked back through my steps i followed last week to make it more efficient. Not worth my time to dig deeper but I suspect that the reason it broke was because I removed the Table.Buffer step, which I now realize had a nested Table.Sort in it. So the table was all out of whack bc it wasn't being sorted correctly anymore. That has been fixed.  

### 12-20-2021
 - removed the following queries:
	 - issueSheet_table
	 - bom_blend_query and table
	 - updated naming conventions for all other queries and tables then went through. Updated all the macros so they now work with the new names. 
	 - massive cleanup of the blendThese query as well
	 - updated the Hx/Dm query so that it keeps the columns i want rather than removing the ones I don't (per <a href="https://www.youtube.com/watch?v=1wdNNgSAZ7k">this video</a>.
 - not gonna keep fooling with it now but noting for later:
	 - the url for querying the prod schedule thru sharepoint is as follows:
	```
	let
		Source = SharePoint.Files("https://adminkinpak.sharepoint.com/sites/PDTN/", [ApiVersion = 15])
	in
		Source
	```

 - Removed the AllCounts clearnreturn subroutine as well as the AllCounts logging sub and the IssueSheet clearnreturn selection statements for sheet code
 - Updated the LotNumGen sub so that it now uses an array instead of copypasta. Also built in logic to max at 5100 for hx and 2925 for dm

### 12-16-2021
 - removed the following functions because they are not needed in the blending schedule:
	 - `blendshtGen1wk`, `blendshtGen2wk`, `blendshtGen3wk`, `hotRoomSheet`
 - changed macrosOn/Off so that it only does cell AA1. Changed all the sheet code so that the selection change based macros only trigger based on value of cell AA1.

### 12-15-2021
 - Updated prod#mergeSheets query to trim rows before preview per code from JRD

### 12-14-2021 
 - Updated the name of the blendData and bom.master tables to match convention (tableName_TABLE)
 - Tried to add a function (cell formula) that would return the exact StartTime value of the run at which we will run short. Right now it is returning zero, and it is doing so very quickly.  Will return to this in the morning.   

### 12-1-2021
 - JRD pointed out a pitfall that I was running into with the empty row cleanup loop on startron report macro; resolved and pushed
     - rows would eventually be skipped because I was using the loop incrementor to store what row i was on, and number of the row I was on would change every time i deleted a row. once I added a separate `Integer` to track the row count, everything worked as intended. 

### 11-30-2021
 - updated lot number lookup macro to reflect new filepath
 - updated history report macro and daily count macro to reflect new filepaths
 - added startron planning macro
     - this will help us plan specifically for startron runs when we have large qtys of different blends all scheduled close together and we can't afford the space required to blend all of them beforehand 
   	 - This macro opens blendData, filters the list once by each different startron PN, copies each version of the list to a new sheet, then makes a table and orders it by start time.
   	 - I can't figure out why the loop I wrote to clear empty rows + the header rows where they copy over, isn't clearing out the very bottom header row in the table. (i am copying over the entire table including headers each time I re-filter the blendData table). I am using Rows(incrementorFromMyFORLoop).EntireRow.Delete to clear out empty rows and header rows. when the for loop gets to row 24, the function deletes row 25.
   	 - ^Leaving as-is because the report is functioning like i need it to. This is a cosmetic issue.
