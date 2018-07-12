Sub CheckSameDUNS()

	Dim k As Integer
	Dim d As Object
	Dim rng As Range
	Dim tmp As Integer

	k = Range("b65536").End(xlUp).Row
	MsgBox ("The bottom line is: " & k)
	tmp = 30

	Set d = CreateObject("scripting.dictionary")
	arr = Range("G14: G")
	'UBound 是指返回array的最后一个index
	For x = 1 To UBound(arr)
		If Not d.exists(arr(x, 1)) Then
			d(arr(x, 1)) = ""
		Else
			If rng Is Nothing Then
				Set rng = Cells(x, "a")
			Else
				Set rng = Union(rng, Cells(x, "a"))
			End If
		End If
	Next x

	[G:G].Interior.ColorIndex = 0
	If Not rng Is Nothing Then
		rng.Interior.ColorIndex = 3
		MsgBox("DUNS 有重复！")
	Else
		Exit Sub
	End If

End Sub


' Sub fhxy()
' 	Dim rng  As Range
' 	Dim d As Object

' 	Set d = CreateObject("scripting.dictionary")

' 	With ActiveSheet
' 		Arr = .Range("G14:Gtmp")

' 		For x = 1 To UBound(Arr)
' 			If Not d.exists(Arr(x, 1)) Then
' 				d(Arr(x, 1)) = ""
' 			Else
' 				If rng Is Nothing Then
' 					Set rng = Cells(x, "a")
' 				Else
' 					Set rng = Union(rng, Cells(x, "a"))
' 				End If
' 			End If
' 		Next x
' 		[a:a].Interior.ColorIndex = 0
' 		If Not rng Is Nothing Then
' 			rng.Interior.ColorIndex = 3
' 			MsgBox ActiveSheet.Name & "有重复内容。"
' 		Else
' 			Exit Sub
' 		End If
' 	End With
' End Sub


' Sub 去重复()
' 	Dim d As Object, sh As Worksheet, arr, i%
' 	On Error Resume Next
' 	Set d = CreateObject("scripting.dictionary")

' 		If Name <> "产品" Then
' 			arr = Range("a2", [a2].End(xlDown))
' 			For i = 1 To UBound(arr)
' 				d.Add arr(i, 1), ""
' 			Next
' 		End If
' 	Sheets("产品").[a2].Resize(d.Count, 1) = Application.Transpose(d.keys)
' End Sub

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