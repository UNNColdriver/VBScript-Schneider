Sub checkDUNS()
	'检查第六列DUNS，1 检查是否是九位；2 检查后面6位是否为数字；3 检查是否有重复
	If k > 1 Then
		For i = 14 To 30
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
					.Color = RGB(255, 255, 0)
					.TintAndShade = 0
					.PatternTintAndShade = 0
				End With
			End If

			If Selection.Interior.Color = RGB(255, 255, 0) Then
				MsgBox ("checked!")
				Selection = "xxxxxxxxx"
				Cells(i, 8) = ""
				Cells(i, 9) = ""
			End If

			If Selection = "xxxxxxxxx" Then
				With Selection.Interior
					.Pattern = xlNone
					.TintAndShade = 0
					.PatternTintAndShade = 0
				End With
			End If
		Next
	End If
End Sub