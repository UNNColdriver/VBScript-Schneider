Sub checkDUNS()

	Dim k As Integer
	k = Range("b65536").End(xlUp).Row
	MsgBox ("The bottom line is: " & k)
	Dim tmp As Integer
	tmp = 30


	'检查第六列DUNS，1 检查是否是九位；2 检查后面6位是否为数字；3 检查是否有重复
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

	If k > 1 Then
		For i = 14 To tmp
			Cells(i, 7).Select
			'Click第二遍的时候识别到单元格的颜色是黄色，msgbox已经将检查，把数字改成xxxxxxxxx，把后面的两个表格清空
			If Selection.Interior.ColorIndex = 3 Then
				MsgBox ("checked!")
				Selection = "xxxxxxxxx"
				Cells(i, 8) = ""
				Cells(i, 9) = ""
			End If


			'Click第三遍之后，如果识别到目标的号码是xxxxxxxxx并且单元格是黄色，把颜色改回白色
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

	' '检查下DUNS是否有重复的，如果有重复的，把重复的两个标出标成绿色RGB(0, 255, 0)
	' Dim d As Object
	' Dim arr As Array
	' Set d = CreateObject("scripting.dictionary")
	' arr = Range("G14: G")
	' For i = 14 To 30
	' 	d(arr(i, 1)) = d(arr(i, 1)) & "," & i
	' Next

	' arr = Range("G14", [G14].End(xlDown))
	' For i = 1 To UBound(arr)
	' 	d.Add arr(i, 1), ""
	' Next

	' For i = 1 To k
	' 	d.Add Cells(i, 7).value = 

Sub 去重复()
	Dim d As Object, sh As Worksheet, arr, i%
	On Error Resume Next
	Set d = CreateObject("scripting.dictionary")

		If Name <> "产品" Then
			arr = Range("a2", [a2].End(xlDown))
			For i = 1 To UBound(arr)
				d.Add arr(i, 1), ""
			Next
		End If
	Sheets("产品").[a2].Resize(d.Count, 1) = Application.Transpose(d.keys)
End Sub


Sub fhxy()
	Dim rng  As Range
	Dim d As Object

	Set d = CreateObject("scripting.dictionary")

	With ActiveSheet
		Arr = .Range("G14:Gtmp")

		For x = 1 To UBound(Arr)
			If Not d.exists(Arr(x, 1)) Then
				d(Arr(x, 1)) = ""
			Else
				If rng Is Nothing Then
					Set rng = Cells(x, "a")
				Else
					Set rng = Union(rng, Cells(x, "a"))
				End If
			End If
		Next x
		[a:a].Interior.ColorIndex = 0
		If Not rng Is Nothing Then
			rng.Interior.ColorIndex = 3
			MsgBox ActiveSheet.Name & "有重复内容。"
		Else
			Exit Sub
		End If
	End With
End Sub



' 	Sub dd()
' 	Dim d As Object, arr, brr, i&, j&, k
' 	Set d = CreateObject("scripting.dictionary")
' 	arr = Range("e1:e" & [a1].CurrentRegion.Rows.Count)
' 	For i = 2 To UBound(arr)
' 		d(arr(i, 1)) = d(arr(i, 1)) & "," & i
' 	Next
' 	ReDim brr(2 To UBound(arr), 1 To 1)
' 	For i = 2 To UBound(arr)
' 		k = Mid(Replace(d(arr(i, 1)), "," & i, ""), 2)
' 		If Len(k) Then brr(i, 1) = "与第" & k & "行重复"
' 	Next
' 	[h2].Resize(UBound(arr) - 1, 1) = brr
' 	Set d = Nothing
' End Sub