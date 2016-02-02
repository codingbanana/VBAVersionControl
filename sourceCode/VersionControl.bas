Attribute VB_Name = "VersionControl"
Public Sub ExportSourceFiles()
'https://christopherjmcclellan.wordpress.com/2014/10/10/vba-and-git/
'http://www.experts-exchange.com/articles/1457/Automate-Exporting-all-Components-in-an-Excel-Project.html
'extract source files
'requires reference to 'Microsoft Visual Basic for Applications Extensibility 5.3 library'
Dim component As VBComponent
Dim destPath As String
destPath = ThisWorkbook.Path & "\sourceCode\"
If Len(Dir(destPath, vbDirectory)) = 0 Then
    MkDir destPath
End If

For Each component In Application.VBE.ActiveVBProject.VBComponents
    If component.CodeModule.CountOfLines > 0 Then
    'Determine the standard extention of the exported file.
    'These can be anything, but for re-importing, should be the following:
        Select Case component.Type
            Case vbext_ct_ClassModule: ext = ".cls"
            Case vbext_ct_Document: ext = ".cls"
            Case vbext_ct_StdModule: ext = ".bas"
            Case vbext_ct_MSForm: ext = ".frm"
            Case vbext_ct_ActiveXDesigner
            Case Else: ext = vbNullString
        End Select
        If ext <> vbNullString Then
            Fname = destPath & component.Name & ext
            'Overwrite the existing file
            'Alternatively, you can prompt the user before killing the file.
            If Dir(Fname, vbNormal) <> vbNullString Then Kill (Fname)
            component.Export Fname
        End If
    End If
Next

End Sub

