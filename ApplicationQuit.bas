Attribute VB_Name = "CloseMSWord"
' Convert MS Word Documents to PDFs on MS Excel
' Author: Takuya Miyashita (miyashita@hydrocoast.jp)
Option Explicit
'Close all documents and MS Word applications
Sub CloseMSWordAll()
    Dim objWord As Object
    On Error Resume Next
    Set objWord = GetObject(, "Word.Application")
    On Error GoTo 0
    Do While Not objWord Is Nothing
        objWord.Quit SaveChanges:=wdDoNotSaveChanges
        Set objWord = Nothing
        On Error Resume Next
        Set objWord = GetObject(, "Word.Application")
        On Error GoTo 0
    Loop
End Sub
