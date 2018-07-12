Sub checkBasicInfo()

	Dim tableRow As Integer
	Dim tmp As Integer
	Dim DITaccount As String
	Dim PlantName As String
	Dim VenusCode As String
	Dim ReportCode As String
	Dim k As Integer
	k = Range("b65536").End(xlUp).Row

	'获取当前表格的最后一行是多少行
	tableRow = Range("b65536").End(xlUp).Row
	MsgBox ("The bottom line is: " & tableRow)

	'tmp是打算检查到第几行的行数，这里是检测用的，防止一次运行太多的代码
	tmp = 80

	'提取每一行数据的第一个数据，所以必须确保这四项数据Row14中的数据正确
	DITaccount = Range("B14").Value
	PlantName = Range("D14").Value
	VenusCode = Range("E14").Value
	ReportCode = Range("F14").Value

	'检查第一列DIT account是否与第一列的数据相同
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


	'自动检查日期是否正确，标准是当前日期减去1  e.g.201806-1 = 201805
	Dim currentYear As Integer
	Dim targetMonth As Integer
	Dim currentTime As String
	currentYear = Year(Date)
	'注意下和实际情况-1
	targetMonth = Month(Date)-2
	currentTime = currentYear &"0"& targetMonth
	'MsgBox "The year is: " & currentYear & " The Month is:" & targetMonth & " The time is:" & currentTime
	
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


	'检查第三列Plant name是否与第一列相同
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


	'检查第四列venuscode是否与第一列相同
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


	'检查第五列 One reporting code是否与第一列相同
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

	'此列的问题是：如果DUNS出现问题，删除了当前的数据，那么再次检查会标红

	' '检查第七列 Supplier local code是否与第一列相同
	' If k > 1 Then
	' 	For i = 14 To tmp
	' 		Cells(i, 8).Select
	' 		If IsNumeric(Selection) Then
	' 			With Selection.Interior
	' 				.Pattern = xlNone
	' 				.TintAndShade = 0
	' 				.PatternTintAndShade = 0
	' 			End With
	' 		Else
	' 			Cells(i, 8).Select
	' 			With Selection.Interior
	' 				.Pattern = xlSolid
	' 				.PatternColorIndex = xlAutomatic
	' 				.Color = RGB(255, 0, 0)
	' 				.TintAndShade = 0
	' 				.PatternTintAndShade = 0
	' 			End With
	' 		End If
	' 	Next
	' End If

	'检查第八列 Supplier Name是否与第一列相同
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
