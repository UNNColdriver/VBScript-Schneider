Sub checkBasicInfo()

	Dim k As Integer
	Dim DITaccount As String
	Dim PlantName As String
	Dim VenusCode As String
	Dim ReportCode As String

	k = Range("b65536").End(xlUp).Row
	MsgBox ("The bottom line is: " & k)
	Dim tmp As Integer
	tmp = 30

	DITaccount = Range("B14").Value
	PlantName = Range("D14").Value
	VenusCode = Range("E14").Value
	ReportCode = Range("F14").Value

	Dim currentYear As Integer
	Dim targetMonth As Integer
	Dim currentTime As String

	currentYear = Year(Date)
	'注意下和实际情况-1
	targetMonth = Month(Date)-2

	currentTime = currentYear &"0"& targetMonth
	MsgBox "The year is: " & currentYear & " The Month is:" & targetMonth & " The time is:" & currentTime


	'检查第一列DIT account
	If k > 1 Then
		For i = 14 To tmp
			Cells(i, 2).Select
			If Cells(i, 2) = DITaccount Then
				' MsgBox ("Fine")
				With Selection.Interior
					.Pattern = xlNone
					.TintAndShade = 0
					.PatternTintAndShade = 0
				End With

			Else
				Cells(i, 2).Select
				Selection = DITaccount
				With Selection.Interior
					.Pattern = xlSolid
					.PatternColorIndex = xlAutomatic
					.Color = RGB(255, 0, 0)
					.TintAndShade = 0
					.PatternTintAndShade = 0
				End With
			End If
		Next
	End If


	'检查第二列日期
	If k > 1 Then
		For i = 14 To tmp
			Cells(i, 3).Select
			If Selection = currentTime Then
				'MsgBox ("Fine")
				With Selection.Interior
					.Pattern = xlNone
					.TintAndShade = 0
					.PatternTintAndShade = 0
				End With

			Else
				Cells(i, 3).Select
				Selection = currentTime
				With Selection.Interior
					.Pattern = xlSolid
					.PatternColorIndex = xlAutomatic
					.Color = RGB(255, 0, 0)
					.TintAndShade = 0
					.PatternTintAndShade = 0
				End With
			End If
		Next
	End If


	'检查第三列Plant name
	If k > 1 Then
		For i = 14 To tmp
			Cells(i, 4).Select
			If Cells(i, 4) = PlantName Then
				' MsgBox ("Fine")
				With Selection.Interior
					.Pattern = xlNone
					.TintAndShade = 0
					.PatternTintAndShade = 0
				End With
				' If Selection.Interior.Color = RGB(255, 0, 0) Then
				' 	Selection.Interior.Color = RGB(255, 0, 0)
				' End If

			Else
				Cells(i, 4).Select
				Selection = PlantName
				With Selection.Interior
					.Pattern = xlSolid
					.PatternColorIndex = xlAutomatic
					.Color = RGB(255, 0, 0)
					.TintAndShade = 0
					.PatternTintAndShade = 0
				End With
			End If
		Next
	End If


	'检查第四列venuscode
	If k > 1 Then
		For i = 14 To tmp
			Cells(i, 5).Select
			If Cells(i, 5) = VenusCode Then
				' MsgBox ("Fine")
				With Selection.Interior
					.Pattern = xlNone
					.TintAndShade = 0
					.PatternTintAndShade = 0
				End With
				' If Selection.Interior.Color = RGB(255, 0, 0) Then
				' 	Selection.Interior.Color = RGB(255, 0, 0)
				' End If

			Else
				Cells(i, 5).Select
				Selection = VenusCode
				With Selection.Interior
					.Pattern = xlSolid
					.PatternColorIndex = xlAutomatic
					.Color = RGB(255, 0, 0)
					.TintAndShade = 0
					.PatternTintAndShade = 0
				End With
			End If
		Next
	End If


	'检查第五列 One reporting code
	If k > 1 Then
		For i = 14 To tmp
			Cells(i, 6).Select
			If Cells(i, 6) =ReportCode Then
				' MsgBox ("Fine")
				With Selection.Interior
					.Pattern = xlNone
					.TintAndShade = 0
					.PatternTintAndShade = 0
				End With
				' If Selection.Interior.Color = RGB(255, 0, 0) Then
				' 	Selection.Interior.Color = RGB(255, 0, 0)
				' End If

			Else
				Cells(i, 6).Select
				Selection = ReportCode
				With Selection.Interior
					.Pattern = xlSolid
					.PatternColorIndex = xlAutomatic
					.Color = RGB(255, 0, 0)
					.TintAndShade = 0
					.PatternTintAndShade = 0
				End With
			End If
		Next
	End If


	'检查第七列 Supplier local code
	If k > 1 Then
		For i = 14 To tmp
			Cells(i, 8).Select
			If IsNumeric(Selection) Then
				With Selection.Interior
					.Pattern = xlNone
					.TintAndShade = 0
					.PatternTintAndShade = 0
				End With
			Else
				Cells(i, 8).Select
				With Selection.Interior
					.Pattern = xlSolid
					.PatternColorIndex = xlAutomatic
					.Color = RGB(255, 0, 0)
					.TintAndShade = 0
					.PatternTintAndShade = 0
				End With
			End If
		Next
	End If


	'检查第八列 Supplier Name
	If k > 1 Then
		For i = 14 To tmp
			Cells(i, 9).Select
			If Len(Selection) = LenB(StrConv(Selection, vbFromUnicode)) Then
				With Selection.Interior
					.Pattern = xlNone
					.TintAndShade = 0
					.PatternTintAndShade = 0
				End With
			Else
				Cells(i, 9).Select
				With Selection.Interior
					.Pattern = xlSolid
					.PatternColorIndex = xlAutomatic
					.Color = RGB(255, 0, 0)
					.TintAndShade = 0
					.PatternTintAndShade = 0
				End With
			End If
		Next
	End If


End Sub
