Attribute VB_Name = "patcher"
Sub Select_Tgt()
Dim flpath As String
With ThisWorkbook.Sheets(1)
    flpath = .Cells(2, 2).Value
    If flpath = "" Then
        flpath = Environ("USERPROFILE") & "\Desktop"
    End If
    .Cells(2, 2).Value = GetFile(flpath)
End With
End Sub

Function GetFile(strPath As String) As String
'Hao Zhang @ 2015.1.29
'returns a file's full path based on user's selection
Dim fl As FileDialog
Dim sItem As String
Set fl = Application.FileDialog(msoFileDialogFilePicker)
With fl
    .Title = "Select a File"
    .AllowMultiSelect = False
    .InitialFileName = strPath
    .Filters.Clear
    .Filters.Add "Macro-Enabled xlsx", "*.xlsm"
    If .Show <> -1 Then GoTo NextCode
    sItem = .SelectedItems(1)
End With
NextCode:
    GetFile = sItem
Set fl = Nothing
End Function

Sub Add_VC()
'Hao Zhang @ 2016.2.2
'Add version control procedure to any xlsm file
'ref:http://www.cpearson.com/excel/vbe.aspx

Application.ScreenUpdating = False
Dim fso As New FileSystemObject
Dim tgtStr As String

tgtStr = ThisWorkbook.Sheets(1).Cells(2, 2).Value
If tgtStr = "" Then
    Call Select_Tgt
End If
tgtStr = ThisWorkbook.Sheets(1).Cells(2, 2).Value

If ThisWorkbook.Sheets(1).CheckBoxes("Check Box 5").Value = 1 Then
'equivalent:
'If ThisWorkbook.Sheets(1).Shapes("Check Box 5").ControlFormat.Value = 1 Then
    tgtStr_dup = fso.getparentfoldername(tgtStr) & "\" & fso.getbasename(tgtStr) & "_dev." & fso.getExtensionName(tgtStr)
    If IsWorkBookOpen(tgtStr) = True Then
        Workbooks(Dir(tgtStr)).SaveCopyAs tgtStr_dup
    Else
        fso.CopyFile tgtStr, tgtStr_dup
    End If
    Set tgtWB = Workbooks.Open(tgtStr_dup)
Else
    If IsWorkBookOpen(tgtStr) = True Then
        Set tgtWB = Workbooks(Dir(tgtStr))
    Else
        Set tgtWB = Workbooks.Open(tgtStr)
    End If
End If

srcPath = ThisWorkbook.Path & "\sourceCode\"
srcfl1 = srcPath & "VersionControl.bas"
srcfl2 = srcPath & "ThisWorkbook.cls"

With tgtWB.VBProject.VBComponents
    For Each Item In tgtWB.VBProject.VBComponents
        If Item.Name = "VersionControl" Then
        .Remove Item
        End If
    Next
    .Import srcfl1
    
    Set TempVBComp = .Import(srcfl2)
    With .Item("Thisworkbook").CodeModule
        .DeleteLines 1, .CountOfLines
        S = TempVBComp.CodeModule.Lines(1, TempVBComp.CodeModule.CountOfLines)
        .InsertLines 1, S
    End With
    On Error GoTo 0
    .Remove TempVBComp
End With

tgtWB.Save
MsgBox "Version Control module added." & Chr(13) & "Now Manually add MS VBA Ext 5.3 library in VBE.", vbOKOnly, "Success"
Application.ScreenUpdating = True

End Sub

Function IsWorkBookOpen(wbPath As String)
'Hao Zhang @ 2015.3.18
'simplified the function, wbPath must be the full path
For Each WB In Workbooks
    '.FullName returns the full path of the workbook
    If WB.FullName = wbPath Then
        IsWorkBookOpen = True
        Exit Function
    Else
        IsWorkBookOpen = False
    End If
Next
End Function
