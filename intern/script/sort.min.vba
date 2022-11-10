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
Dim tableRangesI As Variant, tableLengthI As Integer, tableHeadersI As Variant

' Tabel II
Dim tableRangesIIa As Variant, tableLengthIIa As Integer, tableHeadersIIa As Variant, _
    tableRangesIIb As Variant, tableLengthIIb As Integer, tableHeadersIIb As Variant, _
    tableRangesIIc As Variant, tableLengthIIc As Integer, tableHeadersIIc As Variant, _
    tableRangesIId As Variant, tableLengthIId As Integer, tableHeadersIId As Variant, _
    tableRangesIIe As Variant, tableLengthIIe As Integer, tableHeadersIIe As Variant, _
    tableRangesIIf As Variant, tableLengthIIf As Integer, tableHeadersIIf As Variant, _
    tableRangesIIg As Variant, tableLengthIIg As Integer, tableHeadersIIg As Variant, _
    tableRangesIIh As Variant, tableLengthIIh As Integer, tableHeadersIIh As Variant, _
    tableRangesIIi As Variant, tableLengthIIi As Integer, tableHeadersIIi As Variant, _
    tableRangesIIj As Variant, tableLengthIIj As Integer, tableHeadersIIj As Variant, _
    tableRangesIIk As Variant, tableLengthIIk As Integer, tableHeadersIIk As Variant, _
    tableRangesIIl As Variant, tableLengthIIl As Integer, tableHeadersIIl As Variant, _
    tableRangesIIm As Variant, tableLengthIIm As Integer, tableHeadersIIm As Variant, _
    tableRangesIIn As Variant, tableLengthIIn As Integer, tableHeadersIIn As Variant, _
    tableRangesIIo As Variant, tableLengthIIo As Integer, tableHeadersIIo As Variant, _
    tableRangesIIp As Variant, tableLengthIIp As Integer, tableHeadersIIp As Variant, _
    tableRangesIIq As Variant, tableLengthIIq As Integer, tableHeadersIIq As Variant, _
    tableRangesIIr As Variant, tableLengthIIr As Integer, tableHeadersIIr As Variant, _
    tableRangesIIs As Variant, tableLengthIIs As Integer, tableHeadersIIs As Variant, _
    tableRangesIIt As Variant, tableLengthIIt As Integer, tableHeadersIIt As Variant, _
    tableRangesIIu As Variant, tableLengthIIu As Integer, tableHeadersIIu As Variant, _
    tableRangesIIv As Variant, tableLengthIIv As Integer, tableHeadersIIv As Variant, _
    tableRangesIIw As Variant, tableLengthIIw As Integer, tableHeadersIIw As Variant, _
    tableRangesIIx As Variant, tableLengthIIx As Integer, tableHeadersIIx As Variant

' Tabel III
Dim tableRangesIIIa As Variant, tableLengthIIIa As Integer, tableHeadersIiIa As Variant, _
    tableRangesIIIb As Variant, tableLengthIIIb As Integer, tableHeadersIiIb As Variant, _
    tableRangesIIIc As Variant, tableLengthIIIc As Integer, tableHeadersIiIc As Variant, _
    tableRangesIIId As Variant, tableLengthIIId As Integer, tableHeadersIiId As Variant, _
    tableRangesIIIe As Variant, tableLengthIIIe As Integer, tableHeadersIiIe As Variant, _
    tableRangesIIIf As Variant, tableLengthIIIf As Integer, tableHeadersIiIf As Variant, _
    tableRangesIIIg As Variant, tableLengthIIIg As Integer, tableHeadersIiIg As Variant, _
    tableRangesIIIh As Variant, tableLengthIIIh As Integer, tableHeadersIiIh As Variant, _
    tableRangesIIIi As Variant, tableLengthIIIi As Integer, tableHeadersIiIi As Variant, _
    tableRangesIIIj As Variant, tableLengthIIIj As Integer, tableHeadersIiIj As Variant

' Tabel IV
Dim tableRangesIVa As Variant, tableLengthIVa As Integer, tableHeadersIVa As Variant, _
    tableRangesIVb As Variant, tableLengthIVb As Integer, tableHeadersIVb As Variant

' Tabel V
Dim tableRangesVa As Variant, tableLengthVa As Integer, tableHeadersVa As Variant, _
    tableRangesVb As Variant, tableLengthVb As Integer, tableHeadersVb As Variant, _
    tableRangesVc As Variant, tableLengthVc As Integer, tableHeadersVc As Variant, _
    tableRangesVd As Variant, tableLengthVd As Integer, tableHeadersVd As Variant, _
    tableRangesVe As Variant, tableLengthVe As Integer, tableHeadersVe As Variant, _
    tableRangesVf As Variant, tableLengthVf As Integer, tableHeadersVf As Variant

Dim FileLocation As String, FileTemplateLocation As String, FileExportName As String

Sub Main(isAutoSave)

    FileLocation = Sheets(1).Range("D4")
    FileTemplateLocation = Sheets(1).Range("D5")
    FileExportName = Sheets(1).Range("D6")

    ' Menjalankan Aplikasi
    SEKDABookMaker isAutoSave
End Sub

Sub SEKDABookMaker(isAutoSave)
    ActiveWindow.View = xlNormalView

    Set objWord = CreateObject("Word.Application")
    Set objDoc = objWord.Documents.Open(FileTemplateLocation)
    
    objDoc.SaveAs2 FileExportName

    Application.ScreenUpdating = False

    ' Exporting Phase
    ExportAllTables

    Application.ScreenUpdating = True
    If isAutoSave Then objDoc.Save
End Sub

Sub ExportAllTables()
    '   Tabel I
    ExportTablesI

    '   Tabel II
    ExportTablesII

    '   Tabel III
    ExportTablesIII

    '   Tabel IV
    ExportTablesIV

    '   Tabel V
    ExportTablesV
End Sub

Sub ExportDataExcel(fileName, tableRanges, tableHeaders, tableLength, objWord)
    On Error Resume Next
    Set ImportWorkbook = Workbooks.Open(FileLocation & "/" & fileName)
    On Error GoTo skipThis
    Dim previousCells As String

    For i = 0 To tableLength - 1
        If i > 1 And (i Mod 2) = 0 Then HideSelectedRows ImportWorkbook, previousCells
        If i < tableLength - 2 Then ImportWorkbook.Worksheets(1).Range(tableRanges(i)).Borders(xlEdgeBottom).Weight = xlMedium
        SelectAndCopyFromRange ImportWorkbook, tableRanges(i)
        CopyImage objWord, tableHeaders(i)
        previousCells = tableRanges(i)
        
        'Application.Wait (Now() + TimeValue("00:00:05"))
    Next i

    ImportWorkbook.Application.DisplayAlerts = False
    ImportWorkbook.Close
skipThis:
    Exit Sub
End Sub

Sub HideSelectedRows(ImportWorkbook, selectedRange)
    Dim needToHideCells As String

    needToHideCells = GetRowsNumberFromCells(selectedRange)

    ImportWorkbook.Worksheets(1).Range(needToHideCells).Select

    Selection.EntireRow.Hidden = True
End Sub

Function GetRowsNumberFromCells(selectedRange)
    Dim Cells As Variant
    Dim TopCells As Integer
    Dim BottomCells As Integer
    Dim NewSelectedCells As String
    Cells = Split(selectedRange, ":")

    TopCells = GetRowNumber(Cells(0)) + 2
    BottomCells = GetRowNumber(Cells(1))
    
    NewSelectedCells = TopCells & ":" & BottomCells
    GetRowsNumberFromCells = NewSelectedCells
End Function
    
Function GetRowNumber(Cell)
    Dim x As Integer
    Dim sCleanedStr As String
    For x = 1 To Len(Cell)
        If IsNumeric(Mid(Cell, x, 1)) Then sCleanedStr = sCleanedStr & Mid(Cell, x, 1)
    Next
    GetRowNumber = sCleanedStr
End Function

Sub SelectAndCopyFromRange(ImportWorkbook, selectedRange)
retry:
    ImportWorkbook.Worksheets(1).Range(selectedRange).Select
    ActiveWindow.DisplayGridlines = False
    On Error Resume Next
    Selection.CopyPicture Appearance:=xlScreen, Format:=xlPicture
    On Error GoTo showApplication
    'On Error Resume Next
    ' Selection.Copy
    'On Error GoTo retry
    Exit Sub
showApplication:
    Application.ScreenUpdating = True
    Exit Sub
End Sub

Sub CopyImage(objWord, tableHeader)
retry:
    objWord.Visible = True
    On Error Resume Next
    objWord.Selection.Find.Execute tableHeader
    On Error GoTo skipThis
    objWord.Selection.Paragraphs.Alignment = 1
    On Error Resume Next
    objWord.Selection.Paste
    On Error GoTo retry
    'objWord.Selection.TypeParagraph
skipThis:
    Exit Sub
End Sub

Sub AssignTablesI()
    ' Tabel 1
    tableRangesI = Array("A5:P80", "Q5:AD80")
    tableHeadersI = Array("T101a", "T101b")
    tableLengthI = 2
End Sub

Sub AssignTablesII()
    ' Tabel 1
    tableRangesIIa = Array("A6:M42", "N6:Z42")
    tableHeadersIIa = Array("T201a", "T201b")
    tableLengthIIa = 2

    ' Tabel 2
    tableRangesIIb = Array("A5:P107", "Q5:AD107")
    tableHeadersIIb = Array("T202a", "T202b")
    tableLengthIIb = 2

    ' Tabel 3
    tableRangesIIc = Array("A5:J52", "K5:W52")
    tableHeadersIIc = Array("T203a", "T203b")
    tableLengthIIc = 2

    ' Tabel 4
    tableRangesIId = Array("A5:O63", "P5:AC63", "A5:O105", "P5:AC101")
    tableHeadersIId = Array("T204a", "T204b", "T204c", "T204d")
    tableLengthIId = 4

    ' Tabel 5
    tableRangesIIe = Array("A5:N89", "O5:AA85")
    tableHeadersIIe = Array("T205a", "T205b")
    tableLengthIIe = 2

    ' Tabel 6
    tableRangesIIf = Array("A6:N54", "O6:AA52")
    tableHeadersIIf = Array("T206a", "T206b")
    tableLengthIIf = 2

    ' Tabel 7
    tableRangesIIg = Array("A6:M89", "N6:Z89", "A6:M146", "N6:Z143")
    tableHeadersIIg = Array("T207a", "T207b", "T207c", "T207d")
    tableLengthIIg = 4

    ' Tabel 8
    tableRangesIIh = Array("A6:M82", "N6:Z82", "A6:M146", "N6:Z143")
    tableHeadersIIh = Array("T208a", "T208b", "T208c", "T208d")
    tableLengthIIh = 4

    ' Tabel 9
    tableRangesIIi = Array("A6:M62", "N6:Z62", "A6:M146", "N6:Z143")
    tableHeadersIIi = Array("T209a", "T209b", "T209c", "T209d")
    tableLengthIIi = 4

    ' Tabel 10
    tableRangesIIj = Array("A6:M89", "N6:Z89", "A6:M146", "N6:Z143")
    tableHeadersIIj = Array("T210a", "T210b", "T210c", "T210d")
    tableLengthIIj = 4

    ' Tabel 11
    tableRangesIIk = Array("A6:M89", "N6:Z89", "A6:M146", "N6:Z143")
    tableHeadersIIk = Array("T211a", "T211b", "T211c", "T211d")
    tableLengthIIk = 4

    ' Tabel 12
    tableRangesIIl = Array("A6:M89", "N6:Z89", "A6:M146", "N6:Z143")
    tableHeadersIIl = Array("T212a", "T212b", "T212c", "T212d")
    tableLengthIIl = 4

    ' Tabel 13
    tableRangesIIm = Array("A6:M28", "N6:Z26")
    tableHeadersIIm = Array("T213a", "T213b")
    tableLengthIIm = 2

    ' Tabel 14
    tableRangesIIn = Array("A5:J55", "K5:W53")
    tableHeadersIIn = Array("T214a", "T214b")
    tableLengthIIn = 2

    ' Tabel 15
    tableRangesIIo = Array("A6:N119", "O6:AB119", "A6:N231", "O6:AB231", "A6:N348", "N6:AB342")
    tableHeadersIIo = Array("T215a", "T215b", "T215c", "T215d", "T215e", "T215f")
    tableLengthIIo = 6

    ' Tabel 16
    tableRangesIIp = Array("A6:N58", "O6:AB58", "A6:N107", "O6:AB107", "A6:N164", "N6:AB158")
    tableHeadersIIp = Array("T216a", "T216b", "T216c", "T216d", "T216e", "T216f")
    tableLengthIIp = 6

    ' Tabel 17
    tableRangesIIq = Array("A6:N74", "O6:AC68")
    tableHeadersIIq = Array("T217a", "T217b")
    tableLengthIIq = 2

    ' Tabel 18
    tableRangesIIr = Array("A7:P53", "Q6:AD51")
    tableHeadersIIr = Array("T218a", "T218b")
    tableLengthIIr = 2

    ' Tabel 19
    tableRangesIIs = Array("A6:M119", "N6:Z116")
    tableHeadersIIs = Array("T219a", "T219b")
    tableLengthIIs = 2

    ' Tabel 20
    tableRangesIIt = Array("A6:N63", "O6:AB57")
    tableHeadersIIt = Array("T220a", "T220b")
    tableLengthIIt = 2

    ' Tabel 21
    tableRangesIIu = Array("A6:N116", "O6:AB116", "A6:N224", "O6:AB224", "A6:N338", "O6:AB332")
    tableHeadersIIu = Array("T221a", "T221b", "T221c", "T221d", "T221e", "T221f")
    tableLengthIIu = 6

    ' Tabel 22
    tableRangesIIv = Array("A4:N21", "O4:AA21")
    tableHeadersIIv = Array("T222a", "T222b")
    tableLengthIIv = 2

    ' Tabel 23
    tableRangesIIw = Array("A6:M89", "N6:Z89", "A6:M146", "N6:Z143")
    tableHeadersIIw = Array("T223a", "T223b", "T223c", "T223d")
    tableLengthIIw = 4

    ' Tabel 24
    tableRangesIIx = Array("A6:M25", "N6:Z23")
    tableHeadersIIx = Array("T224a", "T224b")
    tableLengthIIx = 2
End Sub

Sub AssignTablesIII()
    ' Tabel 1
    tableRangesIIIa = Array("B5:M85", "N5:AD85")
    tableHeadersIiIa = Array("T301a", "T301b")
    tableLengthIIIa = 2
    
    ' Tabel 2
    tableRangesIIIb = Array("B5:O85", "P5:AD85")
    tableHeadersIiIb = Array("T302a", "T302b")
    tableLengthIIIb = 2
    
    ' Tabel 3
    tableRangesIIIc = Array("B6:O64", "P6:AC64")
    tableHeadersIiIc = Array("T303a", "T303b")
    tableLengthIIIc = 2
    
    ' Tabel 4
    tableRangesIIId = Array("B5:M86", "N5:AD86")
    tableHeadersIiId = Array("T304a", "T304b")
    tableLengthIIId = 2
    
    ' Tabel 5
    tableRangesIIIe = Array("B5:O86", "P5:AC86")
    tableHeadersIiIe = Array("T305a", "T305b")
    tableLengthIIIe = 2
    
    ' Tabel 6
    tableRangesIIIf = Array("B6:O65", "P6:AC65")
    tableHeadersIiIf = Array("T306a", "T306b")
    tableLengthIIIf = 2
    
    ' Tabel 7
    tableRangesIIIg = Array("B6:O64", "P6:AC64")
    tableHeadersIiIg = Array("T307a", "T307b")
    tableLengthIIIg = 2

    ' Tabel 8
    tableRangesIIIh = Array("B6:O65", "P6:AC65")
    tableHeadersIiIh = Array("T308a", "T308b")
    tableLengthIIIh = 2

    ' Tabel 9
    tableRangesIIIi = Array("B5:O85", "P5:AD85")
    tableHeadersIiIi = Array("T309a", "T309b")
    tableLengthIIIi = 2

    ' Tabel 10
    tableRangesIIIj = Array("B5:O82", "P5:AC82")
    tableHeadersIiIj = Array("T310a", "T310b")
    tableLengthIIIj = 2
End Sub

Sub AssignTablesIV()
    ' Tabel 1
    tableRangesIVa = Array("A5:P80", "Q5:AD80")
    tableHeadersIVa = Array("IV01a", "IV01b")
    tableLengthIVa = 2

    ' Tabel 2
    tableRangesIVb = Array("A5:P80", "Q5:AD80")
    tableHeadersIVb = Array("IV02a", "IV02b")
    tableLengthIVb = 2
End Sub

Sub AssignTablesV()
    ' Tabel 1
    tableRangesVa = Array("B6:BB30", "BC6:BK30")
    tableHeadersVa = Array("T501a", "T501b")
    tableLengthVa = 2

    ' Tabel 2
    tableRangesVb = Array("B5:BB30", "BC5:BK30")
    tableHeadersVb = Array("T502a", "T502b")
    tableLengthVb = 2

    ' Tabel 3
    tableRangesVc = Array("B6:BL38", "BM6:BV38")
    tableHeadersVc = Array("T503a", "T503b")
    tableLengthVc = 2

    ' Tabel 4
    tableRangesVd = Array("B6:BL38", "BM6:BV38")
    tableHeadersVd = Array("T504a", "T504b")
    tableLengthVd = 2

    ' Tabel 5
    tableRangesVe = Array("B6:BB31", "BC6:BK31")
    tableHeadersVe = Array("T505a", "T505b")
    tableLengthVe = 2

    ' Tabel 6
    tableRangesVf = Array("B6:BB30", "BC6:BK30")
    tableHeadersVf = Array("T506a", "T506b")
    tableLengthVf = 2
    
End Sub

Sub ExportTablesI()
    '   Assign Data to Tables
    AssignTablesI

    '   Tabel I
    ExportDataExcel "Tabel I\i01.xls", tableRangesI, tableHeadersI, tableLengthI, objWord
End Sub

Sub ExportTablesII()
    '   Assign Data To Tables
    AssignTablesII

    '   Tabel IIa
    ExportDataExcel "Tabel II\ii01.xls", tableRangesIIa, tableHeadersIIa, tableLengthIIa, objWord

    '   Tabel IIb
    ExportDataExcel "Tabel II\ii02.xls", tableRangesIIb, tableHeadersIIb, tableLengthIIb, objWord

    '   Tabel IIc
    ExportDataExcel "Tabel II\ii03.xls", tableRangesIIc, tableHeadersIIc, tableLengthIIc, objWord

    '   Tabel IId
    ExportDataExcel "Tabel II\ii04.xls", tableRangesIId, tableHeadersIId, tableLengthIId, objWord

    '   Tabel IIe
    ExportDataExcel "Tabel II\ii05.xls", tableRangesIIe, tableHeadersIIe, tableLengthIIe, objWord

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
End Sub

Sub ExportTablesIII()
    '   Assign Data to Tables
    AssignTablesIII

    '   Tabel IIIa
    ExportDataExcel "Tabel III\iii01.xls", tableRangesIIIa, tableHeadersIiIa, tableLengthIIIa, objWord

    '   Tabel IIIb
    ExportDataExcel "Tabel III\iii02.xls", tableRangesIIIb, tableHeadersIiIb, tableLengthIIIb, objWord

    '   Tabel IIIc
    ExportDataExcel "Tabel III\iii03.xls", tableRangesIIIc, tableHeadersIiIc, tableLengthIIIc, objWord

    '   Tabel IIId
    ExportDataExcel "Tabel III\iii04.xls", tableRangesIIId, tableHeadersIiId, tableLengthIIId, objWord

    '   Tabel IIIe
    ExportDataExcel "Tabel III\iii05.xls", tableRangesIIIe, tableHeadersIiIe, tableLengthIIIe, objWord

    '   Tabel IIIf
    ExportDataExcel "Tabel III\iii06.xls", tableRangesIIIf, tableHeadersIiIf, tableLengthIIIf, objWord

    '   Tabel IIIg
    ExportDataExcel "Tabel III\iii07.xls", tableRangesIIIg, tableHeadersIiIg, tableLengthIIIg, objWord

    '   Tabel IIIh
    ExportDataExcel "Tabel III\iii08.xls", tableRangesIIIh, tableHeadersIiIh, tableLengthIIIh, objWord

    '   Tabel IIIi
    ExportDataExcel "Tabel III\iii09.xls", tableRangesIIIi, tableHeadersIiIi, tableLengthIIIi, objWord

    '   Tabel IIIj
    ExportDataExcel "Tabel III\iii10.xls", tableRangesIIIj, tableHeadersIiIj, tableLengthIIIj, objWord
End Sub

Sub ExportTablesIV()
    '   Assign Data to Tables
    AssignTablesIV

    '   Tabel IVa
    ExportDataExcel "Tabel IV\iv01.xls", tableRangesIVa, tableHeadersIVa, tableLengthIVa, objWord

    '   Tabel IVb
    ExportDataExcel "Tabel IV\iv02.xls", tableRangesIVb, tableHeadersIVb, tableLengthIVb, objWord
End Sub

Sub ExportTablesV()
    '   Assign Data to Tables
    AssignTablesV

    '   Tabel Va
    ExportDataExcel "Tabel V\v01.xls", tableRangesVa, tableHeadersVa, tableLengthVa, objWord

    '   Tabel Vb
    ExportDataExcel "Tabel V\v02.xls", tableRangesVb, tableHeadersVb, tableLengthVb, objWord

    '   Tabel Vc
    ExportDataExcel "Tabel V\v03.xls", tableRangesVc, tableHeadersVc, tableLengthVc, objWord

    '   Tabel Vd
    ExportDataExcel "Tabel V\v04.xls", tableRangesVd, tableHeadersVd, tableLengthVd, objWord

    '   Tabel Ve
    ExportDataExcel "Tabel V\v05.xls", tableRangesVe, tableHeadersVe, tableLengthVe, objWord

    '   Tabel Vf
    ExportDataExcel "Tabel V\v06.xls", tableRangesVf, tableHeadersVf, tableLengthVf, objWord

End Sub
