分解成不同的sheet
Sub SaveSeparately()

	Dim sht As Worksheet
	Application.ScreenUpdating = False
	ipath = ThisWorkbook.Path &'\'
	For Each sht In Sheets
		sht.Copy
		ActiveWorkbook.SaveAs ipath & sht.Name& '.xls'
		ActiveWorkbook.Close
	Next
	Application.ScreenUpdating = True
End Sub


Sub 统计工作簿有多少张工作表()
	MsgBox Sheets.Count
End Sub
'keyong


在表格的下一位加入序号1到i
Sub AddSerialNumbers()
	Dim i As Integer

	On Error GoTo Last
	i = InputBox("Enter Value", "Enter Serial Numbers")
	For i = 1 To i
	ActiveCell.Value = i
	ActiveCell.Offset(1, 0).Activate
	Next i
	Last:Exit Sub
End Sub


加入多列
Sub InsertMultipleColumns()
	Dim i As Integer
	Dim j As Integer

	ActiveCell.EntireColumn.Select
	On Error GoTo Last
	i = InputBox("Enter number of columns to insert", "Insert Columns")
	For j = 1 To i
	Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromRightorAbove
	Next j
	Last:Exit Sub
End Sub


加入多行
Sub InsertMultipleRows()
	Dim i As Integer
	Dim j As Integer
	ActiveCell.EntireRow.Select
	On Error GoTo Last
	i = InputBox("Enter number of columns to insert", "Insert Columns")
	For j = 1 To i
	Selection.Insert Shift:=xlToDown, CopyOrigin:=xlFormatFromRightorAbove
	Next j
	Last:Exit Sub
End Sub



自动适应单元长度的大小
Sub AutoFitColumns()
Cells.Select
Cells.EntireColumn.AutoFit
End Sub


=IF(OR(LEN(G90) = 0,AND(ISNUMBER(VALUE(RIGHT(G90,6))),LEN(G90) >= 6)),0,1) = 1
=IF(,0,1)
OR(,) 两个不对
LEN(G90) = 0不对，g90不能为空，应该为空
AND(,) 不对
ISNUMBER(VALUE(RIGHT(G90,6))) 和 LEN(G90) >= 6 应该都是对的
G90从右往左数6位都是数字，G90>=6



Sub RemoveWrapText()
	Cells.Select 
	Selection.WrapText = False
	Cells.EntireRow.AutoFit
	Cells.EntireColumn.AutoFit
End Sub



Sub OpenCalculator()
Application.ActivateMicrosoftApp Index:=0
End Sub

Sub dateInHeader()
	With ActiveSheet.PageSetup
	.LeftHeader = ""
	.CenterHeader = "&D"
	.RightHeader = ""
	.LeftFooter = ""
	.CenterFooter = ""
	.RightFooter = ""
	End With
	ActiveWindow.View = xlNormalView
End Sub


Sub progressStatusBar()
	Application.StatusBar= "Start Printing the Numbers"
	For icntr= 1 To 5000
	Cells(icntr, 1) = icntr
	Application.StatusBar= " Please wait while printing the numbers " & Round((icntr/ 5000 * 100), 0) & "%"
	Next 
	Application.StatusBar= ""
End Sub



12. Highlight Duplicates from Selection
This macro will check each cell of your selection 
and highlight the duplicate values. You can also change the color from the code.
找到相同的值高亮
Sub HighlightDuplicateValues()
	Dim myRange As Range
	Dim myCell As Range
	Set myRange = Selection
	For Each myCell In myRange
	If WorksheetFunction.CountIf(myRange, myCell.Value) > 1 Then
	myCell.Interior.ColorIndex = 36
	End If
	Next myCell
End Sub


Sub TopTen()
	Selection.FormatConditions.AddTop10
	Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
	With Selection.FormatConditions(1)
		.TopBottom = xlTop10Top
		.Rank = 10
		.Percent = False
	End With
	With Selection.FormatConditions(1).Font
		.Color = -16752384
		.TintAndShade = 0
	End With
	With Selection.FormatConditions(1).Interior
		.PatternColorIndex = xlAutomatic
		.Color = 13561798
		.TintAndShade = 0
	End With
	Selection.FormatConditions(1).StopIfTrue = False
End Sub



Sub HighlightGreaterThanValues()
	Dim i As Integer
	i = InputBox("Enter Greater Than Value", "Enter Value")
	Selection.FormatConditions.Delete
	Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:=i
	Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
	With Selection.FormatConditions(1)
	.Font.Color = RGB(0, 0, 0)
	.Interior.Color = RGB(31, 218, 154)
	End With
End Sub


'Test VBS
Sub HLDuplicateValues()
    Dim myRange As Range
    Dim myCell As Range
    Dim i As Long

    i = InputBox("Enter the date as Value: 201807", "Enter Value")

    Set myRange = Worksheets("Check file").Range("C14:C20")
    For Each myCell In myRange
    If WorksheetFunction.CountIf(myRange, myCell.Value) = i Then
    myCell.Interior.ColorIndex = 36
    End If
    Next myCell
End Sub



'28. Highlight Unique Values
'This codes will highlight all the cells from the selection which has a unique value.

Sub highlightUniqueValues()
Dim rng As Range
Set rng = Selection
rng.FormatConditions.Delete
Dim uv As UniqueValues
Set uv = rng.FormatConditions.AddUniqueValues
uv.DupeUnique = xlUnique
uv.Interior.Color = vbGreen
End Sub














