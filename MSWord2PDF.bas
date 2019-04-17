Attribute VB_Name = "ConvertMain"
' Convert MS Word Documents to PDFs on MS Excel
' Author: Takuya Miyashita (miyashita@hydrocoast.jp)
Option Explicit
Sub Main()
    Dim SheetName1, SheetName2 As String
    Dim WS As Worksheet
    Dim nfile, chk As Long
    Dim i As Long
    ' Path
    ChDrive ThisWorkbook.Path
    ChDir ThisWorkbook.Path
    'Stop Screen Updating
    'Application.ScreenUpdating = False

    ' Parameters
    SheetName1 = "IO"
    SheetName2 = "FileConfig"
    
    ' Initial Setup
    Set WS = ThisWorkbook.Worksheets(SheetName1)
    nfile = WS.Range("B1").End(xlDown).row - 1
    chk = WS.Range("C1").End(xlDown).row - 1
    
    ' Check 1
    If nfile > 100 Then
        MsgBox "No file to be converted" + vbCrLf + "セルにファイル名を入力してください"
        Exit Sub
    End If
            
    ' Check 2
    If chk <> nfile Then
        MsgBox "Error: Invalid number of I/O file" + vbCrLf + "I/Oファイルの数が一致していません"
        Exit Sub
    End If

    ' Check 3
    Dim ChkObj As Object
    On Error Resume Next
    Set ChkObj = GetObject(, "Word.Application")
    On Error GoTo 0
    If Not ChkObj Is Nothing Then
        MsgBox "Error: Close MS Word application before this procedure" + vbCrLf + "Wordアプリケーションが開いています．閉じてから再度実行してください．"
        Exit Sub
    End If
    ''''''Call CloseMSWordAll

    'Start Word Application
    Dim objWord As Word.Application
    Set objWord = CreateObject("Word.Application")
    With objWord
         ' Keep displayed
         '.Visible = True
         ' Or hidden
         .Visible = False
    End With
        
    ' Convert each file
    For i = 1 To nfile
        Call SetFileInfo(SheetName1, i + 1, 2, SheetName2, 2, 3)
        Call SetFileInfo(SheetName1, i + 1, 3, SheetName2, 3, 3)
        'Convert to PDF
        Call Convert(objWord)
    Next i
    
    'Activate Screen Updating
    'Application.ScreenUpdating = True
    
    'Close the application
    objWord.Quit
    
    MsgBox "Successfully completed"
End Sub
Sub SetFileInfo(ISheet, Irow, Icol, OSheet, Orow, Ocol)
    Dim IWS, OWS As Worksheet
    Set IWS = ThisWorkbook.Worksheets(ISheet)
    Set OWS = ThisWorkbook.Worksheets(OSheet)
    OWS.Cells(Orow, Ocol).Value = IWS.Cells(Irow, Icol).Value
End Sub
Sub GetFileName(str1, str2)
    Dim MainWS As Worksheet
    Set MainWS = ThisWorkbook.Worksheets("FileConfig")
    str1 = MainWS.Cells(2, 3).Value
    str2 = MainWS.Cells(3, 3).Value
End Sub
Sub Convert(objWord)
    Dim docname, pdfname As String
    Call GetFileName(docname, pdfname)
    Call MSWord2PDF(objWord, docname, pdfname)
End Sub
Sub MSWord2PDF(objWord, docname, pdfname)
    Dim openDoc As Word.Document
    
    ' In order to activate the relative path
    objWord.ChangeFileOpenDirectory ThisWorkbook.Path
    
    'Open MSWord Document (ReadOnly)
    Set openDoc = objWord.Documents.Open(Filename:=docname, ReadOnly:=True)
        
    '**** Convert the Document to PDF ****
    openDoc.ExportAsFixedFormat _
        OutputFileName:=pdfname, _
        ExportFormat:=wdExportFormatPDF, _
        OpenAfterExport:=False, _
        OptimizeFor:=wdExportOptimizeForPrint, _
        Range:=wdExportAllDocument, _
        Item:=wdExportDocumentContent, _
        IncludeDocProps:=True, _
        KeepIRM:=True, _
        CreateBookmarks:=wdExportCreateHeadingBookmarks, _
        DocStructureTags:=True, _
        BitmapMissingFonts:=True, _
        UseISO19005_1:=False

    'Close the document
    openDoc.Close SaveChanges:=False
End Sub
