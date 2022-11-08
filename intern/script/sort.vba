' Copyright (c) 2022 - EmptyWork
' github.com/EmptyWork
' emptywork.my.id

' SEKDA Book Maker
'
' Program sederhana untuk mengambil data dari excel
' dan menambahkan ke word, sesuai dengan template
' yang telah disediakan

' Body Of The Application
Dim objWord, objDoc As Object

' Tabel I
' Tabel I01.xls
Dim tableRangesI As Variant, tableLengthI As Integer, tableHeadersI As Variant

' Tabel II
' Tabel II01.xls
Dim tableRangesIIa As Variant, tableLengthIIa As Integer, tableHeadersIIa As Variant

' Tabel II02.xls
Dim tableRangesIIb As Variant, tableLengthIIb As Integer, tableHeadersIIb As Variant

' Tabel II03.xls
Dim tableRangesIIc As Variant, tableLengthIIc As Integer, tableHeadersIIc As Variant

' Tabel II04.xls
Dim tableRangesIId As Variant, tableLengthIId As Integer, tableHeadersIId As Variant

' Tabel II05.xls
Dim tableRangesIIe As Variant, tableLengthIIe As Integer, tableHeadersIIe As Variant

' Tabel II06.xls
Dim tableRangesIIf As Variant, tableLengthIIf As Integer, tableHeadersIIf As Variant

' Tabel II07.xls
Dim tableRangesIIg As Variant, tableLengthIIg As Integer, tableHeadersIIg As Variant

' Tabel II08.xls
Dim tableRangesIIh As Variant, tableLengthIIh As Integer, tableHeadersIIh As Variant

' Tabel II09.xls
Dim tableRangesIIi As Variant, tableLengthIIi As Integer, tableHeadersIIi As Variant

' Tabel II10.xls
Dim tableRangesIIj As Variant, tableLengthIIj As Integer, tableHeadersIIj As Variant

' Tabel II11.xls
Dim tableRangesIIk As Variant, tableLengthIIk As Integer, tableHeadersIIk As Variant

' Tabel II12.xls
Dim tableRangesIIl As Variant, tableLengthIIl As Integer, tableHeadersIIl As Variant

' Tabel II13.xls
Dim tableRangesIIm As Variant, tableLengthIIm As Integer, tableHeadersIIm As Variant

' Tabel II14.xls
Dim tableRangesIIn As Variant, tableLengthIIn As Variant, tableHeadersIIn As Variant

' Tabel II15.xls
Dim tableRangesIIo As Variant, tableLengthIIo As Integer, tableHeadersIIo As Variant

' Tabel II16.xls
Dim tableRangesIIp As Variant, tableLengthIIp As Integer, tableHeadersIIp As Variant

' Tabel II17.xls
Dim tableRangesIIq As Variant, tableLengthIIq As Integer, tableHeadersIIq As Variant

' Tabel II18.xls
Dim tableRangesIIr As Variant, tableLengthIIr As Integer, tableHeadersIIr As Variant

' Tabel II19.xls
Dim tableRangesIIs As Variant, tableLengthIIs As Integer, tableHeadersIIs As Variant

' Tabel II20.xls
Dim tableRangesIIt As Variant, tableLengthIIt As Integer, tableHeadersIIt As Variant

' Tabel II21.xls
Dim tableRangesIIu As Variant, tableLengthIIu As Integer, tableHeadersIIu As Variant

' Tabel II22.xls
Dim tableRangesIIv As Variant, tableLengthIIv As Integer, tableHeadersIIv As Variant

' Tabel II23.xls
Dim tableRangesIIw As Variant, tableLengthIIw As Integer, tableHeadersIIw As Variant

' Tabel II24.xls
Dim tableRangesIIx As Variant, tableLengthIIx As Integer, tableHeadersIIx As Variant

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

    tableRangesIId = Array("A5:O63", "P5:AC63", "A5:O6,A64:O105", "P5:AC6,P64:AC101")
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
    tableHeadersIIl = Array("II12a", "II12b", "II12c", "II12d")
    tableLengthIIl = 4

    tableRangesIIm = Array("A6:M28", "N6:Z26")
    tableHeadersIIm = Array("II13a", "II13b")
    tableLengthIIm = 2

    tableRangesIIn = Array("A5:J55", "K5:W53")
    tableHeadersIIn = Array("II14a", "II14b")
    tableLengthIIn = 2

    tableRangesIIo = Array("A6:N119", "O6:AB119", "A121:N231", "O121:AB231", "A233:N386", "N233:AB342")
    tableHeadersIIo = Array("II15a", "II15b", "II15c", "II15d", "II15e", "II15f")
    tableLengthIIo = 6

    tableRangesIIp = Array("A6:N58", "O6:AB58", "A59:N107", "O59:AB107", "A109:N164", "N109:AB158")
    tableHeadersIIp = Array("II16a", "II16b", "II16c", "II16d", "II16e", "II16f")
    tableLengthIIp = 6

    tableRangesIIq = Array("A6:N74", "O6:AC68")
    tableHeadersIIq = Array("II17a", "II17b")
    tableLengthIIq = 2

    tableRangesIIr = Array("A7:P53", "Q6:AD51")
    tableHeadersIIr = Array("II18a", "II18b")
    tableLengthIIr = 2

    tableRangesIIs = Array("A6:M119", "N6:Z116")
    tableHeadersIIs = Array("II19a", "II19b")
    tableLengthIIs = 2

    tableRangesIIt = Array("A6:N63", "O6:AB57")
    tableHeadersIIt = Array("II20a", "II20b")
    tableLengthIIt = 2

    tableRangesIIu = Array("A6:N116", "O6:AB116", "A117:N224", "O117:AB224", "A225:N338", "O225:AB332")
    tableHeadersIIu = Array("II21a", "II21b", "II21c", "II21d", "II21e", "II21f")
    tableLengthIIu = 6

    tableRangesIIv = Array("A4:N21", "O4:AA21")
    tableHeadersIIv = Array("II22a", "II22b")
    tableLengthIIv = 2

    tableRangesIIw = Array("A6:M89", "N6:Z89", "A90:M146", "N90:Z143")
    tableHeadersIIw = Array("II23a", "II23b", "II23c", "II23d")
    tableLengthIIw = 4

    tableRangesIIx = Array("A6:M25", "N6:Z23")
    tableHeadersIIx = Array("II24a", "II24b")
    tableLengthIIx = 2

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

    '   Tabel IIi
    ExportDataExcel "Tabel II\ii09.xls", tableRangesIIi, tableHeadersIIi, tableLengthIIi, objWord

    '   Tabel IIj
    ExportDataExcel "Tabel II\ii10.xls", tableRangesIIj, tableHeadersIIj, tableLengthIIj, objWord

    '   Tabel IIk
    ExportDataExcel "Tabel II\ii11.xls", tableRangesIIk, tableHeadersIIk, tableLengthIIk, objWord

    '   Tabel IIl
    ExportDataExcel "Tabel II\ii12.xls", tableRangesIIl, tableHeadersIIl, tableLengthIIl, objWord

    '   Tabel IIm
    ExportDataExcel "Tabel II\ii13.xls", tableRangesIIm, tableHeadersIIm, tableLengthIIm, objWord

    '   Tabel IIn
    ExportDataExcel "Tabel II\ii14.xls", tableRangesIIn, tableHeadersIIn, tableLengthIIn, objWord

    '   Tabel IIo
    ExportDataExcel "Tabel II\ii15.xls", tableRangesIIo, tableHeadersIIo, tableLengthIIo, objWord

    '   Tabel IIp
    ExportDataExcel "Tabel II\ii16.xls", tableRangesIIp, tableHeadersIIp, tableLengthIIp, objWord

    '   Tabel IIq
    ExportDataExcel "Tabel II\ii17.xls", tableRangesIIq, tableHeadersIIq, tableLengthIIq, objWord

    '   Tabel IIr
    ExportDataExcel "Tabel II\ii18.xls", tableRangesIIr, tableHeadersIIr, tableLengthIIr, objWord

    '   Tabel IIs
    ExportDataExcel "Tabel II\ii19.xls", tableRangesIIs, tableHeadersIIs, tableLengthIIs, objWord
    
    '   Tabel IIt
    ExportDataExcel "Tabel II\ii20.xls", tableRangesIIt, tableHeadersIIt, tableLengthIIt, objWord
    
    '   Tabel IIu
    ExportDataExcel "Tabel II\ii21.xls", tableRangesIIu, tableHeadersIIu, tableLengthIIu, objWord
    
    '   Tabel IIv
    ExportDataExcel "Tabel II\ii22.xls", tableRangesIIv, tableHeadersIIv, tableLengthIIv, objWord
    
    '   Tabel IIw
    ExportDataExcel "Tabel II\ii23.xls", tableRangesIIw, tableHeadersIIw, tableLengthIIw, objWord
    
    '   Tabel IIx
    ExportDataExcel "Tabel II\ii24.xls", tableRangesIIx, tableHeadersIIx, tableLengthIIx, objWord

    objDoc.SaveAs2 "Table I, II"
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
retry:
    ImportWorkbook.Worksheets(1).Range(selectedRange).Select
    ActiveWindow.DisplayGridlines = False
    On Error Resume Next
    ' Selection.CopyPicture Appearance:=xlScreen, Format:=xlPicture
    Selection.Copy
    On Error GoTo retry
End Sub

Sub CopyImage(objWord, tableHeader)
retry:
    objWord.Visible = True
    objWord.Selection.Find.Execute tableHeader
    objWord.Selection.Paragraphs.Alignment = 1
    On Error Resume Next
    objWord.Selection.Paste
    On Error GoTo retry
    objWord.Selection.TypeParagraph
End Sub
