Sub CheckDUNS()

	Dim k As Integer
	k = Range("b65536").End(xlUp).Row
	MsgBox ("The bottom line is: " & k)
	Dim tmp As Integer
	tmp = 30


	'检查第六列DUNS，1 检查是否是九位；2 检查后面6位是否为数字；3 检查是否有重复
	'如果有不符合规定的，标出绿色，现在的问题是如果有xxxxxxxxx之后依然识别有错，不知道这个是否保持？？？？？
	If k > 1 Then
		For i = 14 To tmp
			Cells(i, 7).Select
			If (Len(Selection) = 9 AND IsNumeric(Mid(Selection,4,6)) = True) Then
				'MsgBox ("Fine")
				With Selection.Interior
					.Pattern = xlNone
					.TintAndShade = 0
					.PatternTintAndShade = 0
				End With
			Else
				With Selection.Interior
					.Pattern = xlSolid
					.PatternColorIndex = xlAutomatic
					.ColorIndex = 3
					.TintAndShade = 0
					.PatternTintAndShade = 0
				End With
			End If
		Next
	End If
End Sub