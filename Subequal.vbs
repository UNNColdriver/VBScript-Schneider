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
' This code can judging the four lines