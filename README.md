Excel-VBA-Macros
================

Excel VBA Programs for data analysis

<p>
----------------------------
ATID Interpret
Taylor Rose
Created September 24, 2014
----------------------------
Install and Use
</p><br>

<ol>
<li>Enable the Developer tab/ribbon and click on the Visual Basic button in the far left corner to bring up Visual Basic Editor (alt+F11 or fn+alt+F11 on Mac parallels)</li>
<li> Go to file > Import File... (Ctl+m) and select ATID Interpret.bas</li>
<li>Close Visual Basic Editor window to return to your spreadsheet. Select the tab you want to store your lookup information. NOTE: ATID Interpret looksup values as strings so you may need to convert the first column from integers to strings. Insert a new first column and use formula =text(b2,"0") in cell A2.</li>
<li>4) Call the atidInterpret function in the cell you'd like the interpretation to reside. The function takes 4 parameters: ATIDString , vlRange , primaryColumn , secondaryColumn. You can see these parameters by calling atidInterpret and hitting Ctl+Shift+A.</li>

Parameters

ATIDString - The ATID colon deliminated string you'd like to interpet, e.g., 28386:20220:20203 or b2
vlRange - The range of the lookup table you're using to interpret ATID, e.g., Sheet2!A:F
primaryColumn - The number of the column you want to pass as your primary interpretation, e.g. , 6
secondayColumn - The number of the column you want to pass as a fallback if the primary column you select is blank or "NULL" on certain rows



