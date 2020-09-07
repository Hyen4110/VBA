```
Sub Rownum()
	Dim lastRow As Long
	lastRow = Sheet1.Cells(Rows.Count, 1).End(xlUp).Row
	MsgBox "1번 행의 개수 : " & lastRow


Sub ColNum()
	Dim lastColumn As Long
	lastColumn = Sheet1.Cells(1, Columns.Count).End(xlToLeft).Column
	MsgBox "1번열의 개수 : " & lastColumn
End SUb 
```
