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
Dim FileLocation As String, FileTemplateLocation As String, FileExportName As String

Sub InitilizationOfVariants()
    Dim RangesCollection As Collection, FileNameCollection As Collection, TableIDCollection As Collection

    Set RangesCollection = New Collection
    Set FileNameCollection = New Collection
    Set TableIDCollection = New Collection

    InitilizationOfTableOne RangesCollection, FileNameCollection, TableIDCollection
    InitilizationOfTableTwo RangesCollection, FileNameCollection, TableIDCollection
    InitilizationOfTableThree RangesCollection, FileNameCollection, TableIDCollection
    
    For i = 1 To RangesCollection.Count
        Debug.Print RangesCollection.Item(i), FileNameCollection.Item(i), TableIDCollection.Item(i)
    Next i
End Sub

Sub InitilizationOfTableOne(RCollection As Collection, FNCollection As Collection, TIDCollection As Collection)
    Dim tableRanges As Variant, tableTableCount As Variant, tableFileName As Variant, tableID As Variant
    Dim tempCurrentTable As String, tempCurrentSheet As Integer
    
    tableRanges = Array("B8")
    tableTableCount = Array("C5")
    tableFileName = Array("C6")
    tableID = Array("D8")
    tempCurrentSheet = 2
    tempCurrentTable = Sheet1.Range("D3")

    RangesToCollection RCollection, FNCollection, TIDCollection, tempCurrentTable, _
                        tableRanges, tableTableCount, tableFileName, tableID, tempCurrentSheet
End Sub

Sub InitilizationOfTableTwo(RCollection As Collection, FNCollection As Collection, TIDCollection As Collection)
    Dim tableRanges As Variant, tableTableCount As Variant, tableFileName As Variant, tableID As Variant
    Dim tempCurrentTable As String, tempCurrentSheet As Integer
    
    tableRanges = Array("B8", "F8", "J8", "N8", "B19", "F19", "J19", "N19", _
                        "B30", "F30", "J30", "N30", "B41", "F41", "J41", "N41", _
                        "B52", "F52", "J52", "N52", "B63", "F63", "J63", "N63")
    tableTableCount = Array("C5", "G5", "K5", "O5", "C16", "G16", "K16", "O16", _
                            "C27", "G27", "K27", "O27", "C38", "G38", "K38", "O38", _
                            "C49", "G49", "K49", "O49", "C60", "G60", "K60", "O60")
    tableFileName = Array("C6", "G6", "K6", "O6", "C17", "G17", "K17", "O17", _
                            "C28", "G28", "K28", "O28", "C39", "G39", "K39", "O39", _
                            "C50", "G50", "K50", "O50", "C61", "G61", "K61", "O61")
    tableID = Array("D8", "H8", "L8", "P8", "D19", "H19", "L19", "P19", _
                    "D30", "H30", "L30", "P30", "D41", "H41", "L41", "P41", _
                    "D52", "H52", "L52", "P52", "D63", "H63", "L63", "P63")
    tempCurrentSheet = 3
    tempCurrentTable = Sheet2.Range("D3")

    RangesToCollection RCollection, FNCollection, TIDCollection, tempCurrentTable, _
                        tableRanges, tableTableCount, tableFileName, tableID, tempCurrentSheet
End Sub

Sub InitilizationOfTableThree(RCollection As Collection, FNCollection As Collection, TIDCollection As Collection)
    Dim tableRanges As Variant, tableTableCount As Variant, tableFileName As Variant, tableID As Variant
    Dim tempCurrentTable As String, tempCurrentSheet As Integer
    
    tableRanges = Array("B8", "F8", "J8", "N8", "B19", "F19", "J19", "N19", _
                        "B30", "F30")
    tableTableCount = Array("C5", "G5", "K5", "O5", "C16", "G16", "K16", "O16", _
                            "C27", "G27")
    tableFileName = Array("C6", "G6", "K6", "O6", "C17", "G17", "K17", "O17", _
                            "C28", "G28")
    tableID = Array("D8", "H8", "L8", "P8", "D19", "H19", "L19", "P19", _
                    "D30", "H30")
    tempCurrentSheet = 4
    tempCurrentTable = Sheet3.Range("D3")

    RangesToCollection RCollection, FNCollection, TIDCollection, tempCurrentTable, _
                        tableRanges, tableTableCount, tableFileName, tableID, tempCurrentSheet
End Sub

Sub RangesToCollection(RCollection As Collection, FNCollection As Collection, TIDCollection As Collection, _
                        tFolder, tRanges As Variant, tTableCount As Variant, tFileName As Variant, _
                        tID As Variant, tSheet)
    Dim tempCurrentTableFileName As String, tempRanges As Range, tempTableSize As Integer
    Dim tempValueOfTheRange As String, tempValueOfID As String
    
    Sheets(tSheet).Select
    
    For i = 0 To UBound(tRanges)
        Set tempRanges = ActiveSheet.Range(tRanges(i)).CurrentRegion
        Set tempID = ActiveSheet.Range(tID(i)).CurrentRegion
        tempTableSize = ActiveSheet.Range(tTableCount(i))
        tempCurrentTableFileName = tFolder & "\" & ActiveSheet.Range(tFileName(i))

        For j = 0 To tempTableSize
        tempValueOfTheRange = tempRanges(j + 1, 1).Value
        tempValueOfID = tempID(j + 1, 1).Value
            If Not IsEmpty(tempValueOfTheRange) And tempValueOfTheRange <> "Ranges" Then
                RCollection.Add tempValueOfTheRange
                FNCollection.Add tempCurrentTableFileName
                TIDCollection.Add tempValueOfID
            End If
        Next j
    Next i
End Sub



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
