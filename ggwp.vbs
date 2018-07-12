Public Sub exportAndImport()
    Dim wkBook As Excel.Workbook
    Dim wkComp As VBIDE.VBComponent
    Dim macroPath As String
    Set wkBook = ThisWorkbook
    
    On Error Resume Next
    macroPath = ThisWorkbook.Path & "\" & "export\"
    For Each wkComp In wkBook.VBProject.VBComponents 'export
        If wkComp.Type = vbext_ct_StdModule Then '如果是模块就导出并删除
            wkComp.Export macroPath & wkComp.Name & ".bas"
            wkBook.VBProject.VBComponents.Remove wkComp
        End If
    Next
    macroPath = ThisWorkbook.Path & "\" & "import\"
    tempfile = Dir(macroPath & "*.bas")
    
    While tempfile <> ""
        Set wkComp = wkBook.VBProject.VBComponents.Import(macroPath & tempfile) '导入代码
        wkComp.Name = Left(tempfile, Len(tempfile) - 4)
        tempfile = Dir
    Wend
    Debug.Print "in export and import"
End Sub
    