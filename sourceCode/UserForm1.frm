VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   1380
   ClientLeft      =   48
   ClientTop       =   372
   ClientWidth     =   3708
   OleObjectBlob   =   "UserForm1.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public cursor As Long
Public col As Long
Public EOR As Long

Private Sub UserForm_Initialize()
'Hao Zhang @ 2016.1.26
'load basic info
With ThisWorkbook.Worksheets("Final")
    .Activate
    col = .Rows(1).Find("Flag").Column
    EOR = .Cells(.Rows.Count, 1).End(xlUp).Row
End With
cursor = 2
Range("A" & cursor).Resize(1, col).Select
End Sub

Private Sub btn1_Click()
'Hao Zhang @ 2016.1.26
'navigate to the next flagged row
With ThisWorkbook.Worksheets("Final")
    .Activate
    While cursor <= EOR
        'bring cursor down to the next unflagge row
        If .Cells(cursor, col).Value <> "good" Then
            Do Until .Cells(cursor, col).Value = "good"
                cursor = cursor + 1
            Loop
        End If
    
        Do While .Cells(cursor, col).Value = "good"
            cursor = cursor + 1
        Loop
    
        'ActiveWindow.ScrollRow = cursor
        .Range("A" & cursor).Resize(1, col).Select
        Exit Sub
    Wend
End With

End Sub

Private Sub btn2_Click()
'Hao Zhang @ 2016.1.26
'navigate to the previous flagged row

'bring cursor up to the head of last flagged row
With ThisWorkbook.Worksheets("Final")
    .Activate
    Do While .Cells(cursor, col).Value <> "good"
        cursor = cursor - 1
    Loop
    If cursor < 2 Then
        cursor = 2
        .Range("A" & cursor).Resize(1, col).Select
        Exit Sub
    End If
    
    'find the head of last good data
    Do While .Cells(cursor, col).Value = "good"
        cursor = cursor - 1
    Loop
    If cursor < 2 Then
        cursor = 2
        .Range("A" & cursor).Resize(1, col).Select
        Exit Sub
    End If
    
    'find the head of last flagged data
    Do While .Cells(cursor, col).Value <> "good"
        cursor = cursor - 1
    Loop
    If cursor < 2 Then
        cursor = 2
        .Range("A" & cursor).Resize(1, col).Select
        Exit Sub
    End If
    
    cursor = cursor + 1
    .Range("A" & cursor).Resize(1, col).Select
End With
End Sub


