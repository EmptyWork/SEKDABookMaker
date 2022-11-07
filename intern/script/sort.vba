' Copyright (c) 2022 - EmptyWork
' github.com/EmptyWork
' emptywork.my.id

Dim objWord, objDoc As Object

' Tabel I01.xls
Dim tableRangesI As Variant, tableLengthI As Integer, tableHeadersI As Variant

' Tabel II01.xls
Dim tabkeRangesIIa As Variant, tableLengthIIa As Variant, tableHeadersIIa As Variant

' Tabel II02.xls
Dim tabkeRangesIIb As Variant, tableLengthIIb As Variant, tableHeadersIIb As Variant

' Tabel II03.xls
Dim tabkeRangesIIc As Variant, tableLengthIIc As Variant, tableHeadersIIc As Variant

' Tabel II04.xls
Dim tabkeRangesIId As Variant, tableLengthIId As Variant, tableHeadersIId As Variant

' Tabel II05.xls
Dim tabkeRangesIIe As Variant, tableLengthIIe As Variant, tableHeadersIIe As Variant

' Tabel II06.xls
Dim tabkeRangesIIf As Variant, tableLengthIIf As Variant, tableHeadersIIf As Variant

' Tabel II07.xls
Dim tabkeRangesIIg As Variant, tableLengthIIg As Variant, tableHeadersIIg As Variant

' Tabel II08.xls
Dim tabkeRangesIIh As Variant, tableLengthIIh As Variant, tableHeadersIIh As Variant

' Tabel II09.xls
Dim tabkeRangesIIi As Variant, tableLengthIIi As Variant, tableHeadersIIi As Variant

' Tabel II10.xls
Dim tabkeRangesIIj As Variant, tableLengthIIj As Variant, tableHeadersIIj As Variant

' Tabel II11.xls
Dim tabkeRangesIIk As Variant, tableLengthIIk As Variant, tableHeadersIIk As Variant

' Tabel II12.xls
Dim tabkeRangesIIl As Variant, tableLengthIIl As Variant, tableHeadersIIl As Variant

Sub Start()
    tableRangesI = Array("A5:P80", "Q5:AD80")
    tableHeadersI = Array("I01a", "I01b")
    tableLengthI = 2
    
    tableRangesIIa = Array("A6:M42", "N6:Z42")
    tableHeadersIIa = Array("II01a", "II01b")
    tableLengthIIa = 2
    
    tableRangesIIb = Array("A5:P107", "Q5:AD107")
    tableHeadersIIb = Array("II02a", "II02b")
    tableLengthIIb = 2
    
    tableRangesIIc = Array("A5:J52", "K5:W52")
    tableHeadersIIc = Array("II03a", "II03b")
    tableLengthIIc = 2

    tableRangesIId = Array("A5:O63", "P5:AC63", "A64:O105", "P64:AC101")
    tableHeadersIId = Array("II04a", "II04b", "II04c", "II04d")
    tableLengthIId = 4

    tableRangesIIe = Array("A5:N89", "O5:AA85")
    tableHeadersIIe = Array("II05a", "II05b")
    tableLengthIIe = 2

    tableRangesIIf = Array("A6:N54", "O6:AA52")
    tableHeadersIIf = Array("II06a", "II06b")
    tableLengthIIf = 2

    tableRangesIIg = Array("A6:M89", "N6:Z89", "A90:M146", "N90:Z143")
    tableHeadersIIg = Array("II07a", "II07b", "II07b", "II07c")
    tableLengthIIg = 4
    
    tableRangesIIh = Array("A6:M82", "N6:Z82", "A83:M146", "N83:Z143")
    tableHeadersIIh = Array("II08a", "II08b", "II08c", "II08d")
    tableLengthIIh = 4

    tableRangesIIi = Array("A6:M62", "N6:Z62", "A63:M146", "N63:Z143")
    tableHeadersIIi = Array("II09a", "II09b", "II09c", "II09d")
    tableLengthIIi = 4

    tableRangesIIj = Array("A6:M89", "N6:Z89", "A90:M146", "N90:Z143")
    tableHeadersIIj = Array("II10a", "II10b", "II10c", "II10d")
    tableLengthIIj = 4

    tableRangesIIk = Array("A6:M89", "N6:Z89", "A90:M146", "N90:Z143")
    tableHeadersIIk = Array("II11a", "II11b", "II11c", "II11d")
    tableLengthIIk = 4

    tableRangesIIl = Array("A6:M89", "N6:Z89", "A90:M146", "N90:Z143")
    tableHeadersII = Array("II12a", "II12b", "II12c", "II12d")
    tableLengthIIl = 4

    ActiveWindow.View = xlNormalView
    
    Set objWord = CreateObject("Word.Application")
    Set objDoc = objWord.Documents.Open("D:\SEKDA\Template\SEKDA.docx")
    
    ' Exporting Phase
    '   Tabel I
    ExportDataExcel "Tabel I\i01.xls", tableRangesI, tableHeadersI, tableLengthI, objWord
    
    '   Tabel IIa
    ExportDataExcel "Tabel II\ii01.xls", tableRangesIIa, tableHeadersIIa, tableLengthIIa, objWord
    
    '   Tabel IIb
    ExportDataExcel "Tabel II\ii02.xls", tableRangesIIb, tableHeadersIIb, tableLengthIIb, objWord

    '   Tabel IIc
    ExportDataExcel "Tabel II\ii03b.xls", tableRangesIIc, tableHeadersIIc, tableLengthIIc, objWord
    
    '   Tabel IId
    ExportDataExcel "Tabel II\ii04.xls", tableRangesIId, tableHeadersIId, tableLengthIId, objWord

    '   Tabel IIe
    ExportDataExcel "Tabel II\ii05b.xls", tableRangesIIe, tableHeadersIIe, tableLengthIIe, objWord
    
    '   Tabel IIf
    ExportDataExcel "Tabel II\ii06.xls", tableRangesIIf, tableHeadersIIf, tableLengthIIf, objWord
    
    '   Tabel IIg
    ExportDataExcel "Tabel II\ii07.xls", tableRangesIIg, tableHeadersIIg, tableLengthIIg, objWord

    '   Tabel IIh
    ExportDataExcel "Tabel II\ii08.xls", tableRangesIIh, tableHeadersIIh, tableLengthIIh, objWord

    objDoc.SaveAs2 "Table I & IIh"
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
