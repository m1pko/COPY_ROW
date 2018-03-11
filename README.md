# COPY_ROW

This script enables the copy of specific lines (one at a time) from one Google Spreadsheet to another. 
It also allows to create a dropdown box with selectable options read from a specified range in a sheet of the destination spreadsheet.

How to use the script:

1. Open the source Google Spreadsheet;
2. Select Tools > Script editor...;
3. Copy the code from the COPY_ROW_ON_DEMAND.gs file as it is;
4. Save it;
5. Configure the following 6 varables comprised between the comments "//Global variables - START" and "//Global variables - END":
>columnCount - number of columns intended to copy

>sourceSheetNumber - sheet number, in the source spreadsheet, where the row to copy is, 0 being the first sheet

>destinationSheetNumber - sheet number, in the destiny spreadsheet, where to copy the row, 0 being the first sheet

>destinationUrl - destination spreadsheet URL

>menuName - menu name

>subMenuName - sub menu name


6. Select Run
	It will probable ask for your permission o execute the script; allow it (you'll only need to do this once);
7. Return to the source sheet and execute the script by selecting the newly added menu, inserting the line number you which to copy and clicking in the OK button
8. Check the destination sheet!

That's it!
