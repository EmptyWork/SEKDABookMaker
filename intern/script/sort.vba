' Copyright (c) 2022 - EmptyWork
' github.com/EmptyWork
' emptywork.my.id

Dim objWord, objDoc As Object

' Tabel I01.xls
Dim tableRangesI As Variant
Dim tableLengthI As Integer
Dim tableHeadersI As Variant

Sub Start()
    tableRangesI = Array("A5:P80", "Q5:AD80")
    tableHeadersI = Array("I01a", "I01b")
    tableLengthI = 2
    
    ActiveWindow.View = xlNormalView
    
    Set objWord = CreateObject("Word.Application")
    Set objDoc = objWord.Documents.Open("D:\SEKDA\Template\SEKDA.docx")
    
    ' Exporting Phase
    '   Tabel I
    ExportDataExcel "Tabel I\i01.xls", tableRangesI, tableHeadersI, tableLengthI, objWord

    objDoc.SaveAs2 "Table I"
End Sub

Sub ExportDataExcel(fileName, tableRanges, tableHeaders, tableLength, objWord)
    Set ImportWorkbook = Workbooks.Open("D:\SEKDA\44. Januari 2022\" & fileName)
        
    For i = 0 To tableLength - 1
        SelectAndCopyFromRange ImportWorkbook, tableRanges(i)
        CopyImage objWord, tableHeaders(i)
    Next i

    ImportWorkbook.Application.DisplayAlerts = False
    ImportWorkbook.Close
End Sub

Sub SelectAndCopyFromRange(ImportWorkbook, selectedRange)
    ImportWorkbook.Worksheets(1).Range(selectedRange).Select
    ActiveWindow.DisplayGridlines = False
    Selection.CopyPicture Appearance:=xlScreen, Format:=xlPicture
End Sub

Sub CopyImage(objWord, tableHeader)
    objWord.Visible = True
    objWord.Selection.Find.Execute tableHeader
    objWord.Selection.Paragraphs.Alignment = 1
    objWord.Selection.Paste
    objWord.Selection.TypeParagraph
End Sub