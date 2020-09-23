<div align="center">

## Auto resize flexgrid column widths


</div>

### Description

Automatically resize the columns in any flex grid to give a nice, professional appearance.

Public sub automatically resizes MS Flex Grid columns to match the width of the text, no matter the size of the grid or the number of columns.

Reads first n number of rows of data, and adjusts column size to match the widest cell of text. Will even expand columns proportionately if they aren't wide enough to fill out the entire width of the grid. Configurable constraints allow you to designate

1) Any flex grid to resize

2) Maximum column width

3) the maximum number of rows in depth to look for the widest cell of text.
 
### More Info
 
msFG (MSFlexGrid) = The name of the flex grid to resize .... MaxRowsToParse (integer) = The maximum number of rows (depth) of the table to scan for cell width (e.g. 50) .... MaxColWidth (Integer) = The maximum width of any given cell in twips (e.g. 5000)

Simply drop this public sub into your form or module and access it from anywhere in your program to automatically resize any flex grid.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Jonathan W\. Lartigue](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/jonathan-w-lartigue.md)
**Level**          |Beginner
**User Rating**    |3.9 (31 globes from 8 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/jonathan-w-lartigue-auto-resize-flexgrid-column-widths__1-8547/archive/master.zip)





### Source Code

```
Public Sub AutosizeGridColumns(ByRef msFG As MSFlexGrid, ByVal MaxRowsToParse As Integer, ByVal MaxColWidth As Integer)
Dim I, J As Integer
Dim txtString As String
Dim intTempWidth, BiggestWidth As Integer
Dim intRows As Integer
Const intPadding = 150
With msFG
 For I = 0 To .Cols - 1
  ' Loops through every column
  .Col = I
  ' Set the active colunm
  intRows = .Rows
  ' Set the number of rows
  If intRows > MaxRowsToParse Then intRows = MaxRowsToParse
  ' If there are more rows of data, reset
  ' intRows to the MaxRowsToParse constant
  intBiggestWidth = 0
  ' Reset some values to 0
  For J = 0 To intRows - 1
   ' check up to MaxRowsToParse # of rows and obtain
   ' the greatest width of the cell contents
   .Row = J
   txtString = .Text
   intTempWidth = TextWidth(txtString) + intPadding
   ' The intPadding constant compensates for text insets
   ' You can adjust this value above as desired.
   If intTempWidth > intBiggestWidth Then intBiggestWidth = intTempWidth
   ' Reset intBiggestWidth to the intMaxColWidth value if necessary
  Next J
  .ColWidth(I) = intBiggestWidth
 Next I
 ' Now check to see if the columns aren't as wide as the grid itself.
 ' If not, determine the difference and expand each column proportionately
 ' to fill the grid
 intTempWidth = 0
 For I = 0 To .Cols - 1
  intTempWidth = intTempWidth + .ColWidth(I)
  ' Add up the width of all the columns
 Next I
 If intTempWidth < msFG.Width Then
  ' Compate the width of the columns to the width of the grid control
  ' and if necessary expand the columns.
  intTempWidth = Fix((msFG.Width - intTempWidth) / .Cols)
  ' Determine the amount od width expansion needed by each column
  For I = 0 To .Cols - 1
   .ColWidth(I) = .ColWidth(I) + intTempWidth
   ' add the necessary width to each column
  Next I
 End If
End With
End Sub
```

