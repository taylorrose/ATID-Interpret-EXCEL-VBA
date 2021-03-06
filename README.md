ATID Interpret
================


Use a lookup table to convert an ATID string to a readable attribution list:

`28386:20220:28386:20203` --> ` 1)Organic Search 2)Wireless 3)Organic Search 4)Unbranded`

<h2>Install and Use</h2>

1. Enable the Developer tab/ribbon and click on the Visual Basic button in the far left corner to bring up Visual Basic Editor ( alt+F11 or fn+alt+F11 on Mac parallels)

2. Go to file > Import File... ( Ctl+m ) and select ATID Interpret.bas

3. Close Visual Basic Editor window to return to your spreadsheet. Select the tab you want to store your lookup information. NOTE: ATID Interpret looksup values as strings so you may need to convert the first column to strings if your values are numbers. Insert a new first column and use formula `=text(b2,"0")` in cell A2.

4. Call the atidInterpret function in the cell you'd like the interpretation to reside. The function takes 5 parameters: `ATIDString` , `delim`, `vlRange` , `primaryColumn` , `secondaryColumn`. You can see these parameters by calling `=atidInterpret` in your fomula bar and hitting Ctl+Shift+A.

<h2>Parameters </h2> 

*Formula Example:* `atidInterpret(ATIDString , delim, vlRange , primaryColumn , secondayColumn)` , *e.g.* , `=atidInterpret(b2,":",Sheet2!A:F,6,5)`


`ATIDString` - The ATID colon deliminated string you'd like to interpet, *e.g.* , `28386:20220:20203` or `b2`
<br>
`delim` - The deliminator you want to split your ATIDString by, *e.g.* , `:` Note: Stored as string so use quotes `":"`
`vlRange` - The range of the lookup table you're using to interpret ATID, *e.g.*, `Sheet2!A:F`
<br>
`primaryColumn` - The number of the column you want to pass as your primary interpretation, *e.g.* , `6`
<br>
`secondayColumn` - The number of the column you want to pass as a fallback if the primary column you select is blank or "NULL" on certain rows, *e.g.* , `5`



