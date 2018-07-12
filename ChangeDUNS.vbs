Sub ChangeDUNS()

	Dim k As Integer
	k = Range("b65536").End(xlUp).Row
	MsgBox ("The bottom line is: " & k)
	Dim tmp As Integer
	tmp = 30

	If k > 1 Then
		For i = 14 To tmp
			Cells(i, 7).Select
			'如果识别到单元格的颜色是黄色，msgbox已经将检查，把数字改成xxxxxxxxx，把后面的两个表格清空
			If Selection.Interior.ColorIndex = 3 Then
				MsgBox ("checked!")
				Selection = "xxxxxxxxx"
				Cells(i, 8) = ""
				Cells(i, 9) = ""
			End If

			'把标成颜色的区域退回原来的颜色
			If (Selection = "xxxxxxxxx" AND Selection.Interior.ColorIndex = 3) Then
				With Selection.Interior
					.Pattern = xlNone
					.ColorIndex = 2
					.TintAndShade = 0
					.PatternTintAndShade = 0
				End With
			End If
		Next
	End If

End Sub