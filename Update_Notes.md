# Update Notes

### 12-20-2021

 - removed the following queries:
	 - issueSheet_table
	 - bom_blend_query and table
	 - updated naming conventions for all other queries and tables then went through. Updated all the macros so they now work with the new names. 
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
