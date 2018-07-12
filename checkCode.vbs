Sub checkCode() 
	Dim venus As String
	Dim SCNLDindex As Integer
	Dim SCLTDindex As Integer
	Dim counter1 As Integer
	Dim counter2 As Integer
	Dim tmp As Integer

	k = Range("b65536").End(xlUp).Row
	tmp = 30

	'获取SCNLD和SCLTD的index和位置
	For i = 1 To 25
		Cells(13, i).Select
		If Selection = "SCNLD" Then
			SCNLDindex = i
		End if

		If Selection = "SCLTD" Then
			SCLTDindex = i
		End If
	Next
	MsgBox("gg " & SCNLDindex &"   " & SCLTDindex)

	'如果这SCNLD的下面所有的数字是空的，那么会有msgbox显示说不可以为空
	counter1 = 0
	counter2 = 0
	For j = 14 to tmp
		'判断SCNLD下面是不是空，如果是空，标红，如果不是空
		Cells(j, SCNLDindex).select
		If Selection <> "" Then
			counter1 = counter1 + 1
		End If
	Next
	If counter1 = 0 Then
		MsgBox("The SCNLD column should not be empty")
	End If

	For j = 14 to tmp
		'判断SCNLD下面是不是空，如果是空，标红，如果不是空
		Cells(j, SCLTDindex).select
		If Selection <> "" Then
			counter2 = counter2 + 1
		End If
	Next
	If counter2 = 0 Then
		MsgBox("The SCLTD column should not be empty")
	End If

End Sub













