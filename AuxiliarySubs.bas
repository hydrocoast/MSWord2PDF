Attribute VB_Name = "AuxSubs"
' Author: Takuya Miyashita (miyashita@hydrocoast.jp)
Option Explicit
Sub DocSearch()
    Dim DirPath As String
    Dim row_ini, col_ini As Long
    Dim WS0 As Worksheet
    ' Path
    ChDrive ThisWorkbook.Path
    ChDir ThisWorkbook.Path
    
    Set WS0 = ThisWorkbook.Worksheets("IO")
    row_ini = WS0.Range("B1").End(xlDown).row
    If row_ini > 100 Then row_ini = 1
    col_ini = 2
    
    DirPath = WS0.Cells(2, 1).Value
    Call DocList(DirPath, row_ini, col_ini)
End Sub
Sub ClearList()
    Dim WS0 As Worksheet
    Dim row As Long
    Set WS0 = ThisWorkbook.Worksheets("IO")
    row = WS0.Range("B1").End(xlDown).row
    With WS0
        .Range(.Cells(2, 2), .Cells(row, 2)).Clear
    End With
End Sub
Sub DocList(PathName, row_ini, col_ini)
    Dim fname As String
    Dim cnt As Long
    cnt = 0
    
    fname = Dir(PathName + "\*.docx")
    Do Until fname = ""
        cnt = cnt + 1
        Call PrintWS(row_ini + cnt, col_ini, PathName + "\" + fname)
        fname = Dir()
    Loop

    fname = Dir(PathName + "\*.doc")
    Do Until fname = ""
        cnt = cnt + 1
        Call PrintWS(row_ini + cnt, col_ini, PathName + "\" + fname)
        fname = Dir()
    Loop

End Sub
Sub PrintWS(row, col, str)
    Dim WS As Worksheet
    Set WS = ThisWorkbook.ActiveSheet
    
    WS.Cells(row, col).Value = str
End Sub
