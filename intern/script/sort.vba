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

Sub Main(isAutoSave)

    FileLocation = Sheets(1).Range("D4")
    FileTemplateLocation = Sheets(1).Range("D5")
    FileExportName = Sheets(1).Range("D6")

    ' Menjalankan Aplikasi
    SEKDABookMaker isAutoSave
End Sub

Sub SEKDABookMaker(isAutoSave)
    Dim RangesCollection As Collection, FileNameCollection As Collection, TableIDCollection As Collection, _
        TableLengthCollection As Collection

    Dim previousFileName As String, currentTableIndicator As Integer, currentIndicatorLength As Integer
    Dim ImportWorkbook As Object

    Set RangesCollection = New Collection
    Set FileNameCollection = New Collection
    Set TableIDCollection = New Collection
    Set TableLengthCollection = New Collection

    InitilizationOfVariants RangesCollection, FileNameCollection, TableIDCollection, TableLengthCollection

    ActiveWindow.View = xlNormalView

    Set objWord = CreateObject("Word.Application")
    Set objDoc = objWord.Documents.Open(FileTemplateLocation)
    
    If isAutoSave = True Then objDoc.SaveAs2 FileExportName

    Application.ScreenUpdating = False
    
    currentTableIndicator = 0
    currentLengthIndicator = 0

    For i = 1 To RangesCollection.Count
        Dim previousCells As String
        If previousFileName <> FileNameCollection.Item(i) Then
            If Not ImportWorkbook Is Nothing Then
                ImportWorkbook.Application.DisplayAlerts = False
                ImportWorkbook.Close
                currentLengthIndicator = 0
            End If
            On Error Resume Next
            Set ImportWorkbook = Workbooks.Open(FileLocation & "\" & FileNameCollection.Item(i))
            On Error GoTo tonext
            previousFileName = FileNameCollection.Item(i)
            currentTableIndicator = currentTableIndicator + 1
        End If

        If TableLengthCollection.Item(currentTableIndicator) > 2 Then
            If currentLengthIndicator > 1 And (currentLengthIndicator Mod 2) = 0 Then
                HideSelectedRows ImportWorkbook, previousCells
            End If
            If currentLengthIndicator < TableLengthCollection.Item(currentTableIndicator) - 2 Then
                ImportWorkbook.Worksheets(1).Range(RangesCollection.Item(i)). _
                Borders(xlEdgeBottom).Weight = xlMedium
            End If
        End If
        
        SelectAndCopyFromRange ImportWorkbook, RangesCollection.Item(i)
        CopyImage objWord, TableIDCollection.Item(i)
        previousCells = RangesCollection.Item(i)
        currentLengthIndicator = currentLengthIndicator + 1
tonext:
    Next i

    ImportWorkbook.Application.DisplayAlerts = False
    ImportWorkbook.Close

    Application.ScreenUpdating = True
    If isAutoSave = True Then objDoc.Save
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

Sub InitilizationOfVariants(RCollection As Collection, FNCollection As Collection, _
                            TIDCollection As Collection, TLCollection As Collection)
    'InitilizationOfTableOne RCollection, FNCollection, TIDCollection, TLCollection
    'InitilizationOfTableTwo RCollection, FNCollection, TIDCollection, TLCollection
    'InitilizationOfTableThree RCollection, FNCollection, TIDCollection, TLCollection
    InitilizationOfTableFour RCollection, FNCollection, TIDCollection, TLCollection
    InitilizationOfTableFive RCollection, FNCollection, TIDCollection, TLCollection
End Sub

Sub InitilizationOfTableOne(RCollection As Collection, FNCollection As Collection, _
                            TIDCollection As Collection, TLCollection As Collection)
    Dim tableRanges As Variant, tableTableCount As Variant, tableFileName As Variant, tableID As Variant
    Dim tempCurrentTable As String, tempCurrentSheet As Integer
    
    tableRanges = Array("B8")
    tableTableCount = Array("C5")
    tableFileName = Array("C6")
    tableID = Array("D8")
    tempCurrentSheet = 2
    tempCurrentTable = Sheet1.Range("D3")

    RangesToCollection RCollection, FNCollection, TIDCollection, TLCollection, tempCurrentTable, _
                        tableRanges, tableTableCount, tableFileName, tableID, tempCurrentSheet
End Sub

Sub InitilizationOfTableTwo(RCollection As Collection, FNCollection As Collection, _
                            TIDCollection As Collection, TLCollection As Collection)
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

    RangesToCollection RCollection, FNCollection, TIDCollection, TLCollection, tempCurrentTable, _
                        tableRanges, tableTableCount, tableFileName, tableID, tempCurrentSheet
End Sub

Sub InitilizationOfTableThree(RCollection As Collection, FNCollection As Collection, _
    TIDCollection As Collection, TLCollection As Collection)
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

    RangesToCollection RCollection, FNCollection, TIDCollection, TLCollection, tempCurrentTable, _
                        tableRanges, tableTableCount, tableFileName, tableID, tempCurrentSheet
End Sub

Sub InitilizationOfTableFour(RCollection As Collection, FNCollection As Collection, _
        TIDCollection As Collection, TLCollection As Collection)
    Dim tableRanges As Variant, tableTableCount As Variant, tableFileName As Variant, tableID As Variant
    Dim tempCurrentTable As String, tempCurrentSheet As Integer

    tableRanges = Array("B8", "F8")
    tableTableCount = Array("C5", "G5")
    tableFileName = Array("C6", "G6")
    tableID = Array("D8", "H8")
    tempCurrentSheet = 5
    tempCurrentTable = Sheet4.Range("D3")

    RangesToCollection RCollection, FNCollection, TIDCollection, TLCollection, tempCurrentTable, _
                        tableRanges, tableTableCount, tableFileName, tableID, tempCurrentSheet
End Sub

Sub InitilizationOfTableFive(RCollection As Collection, FNCollection As Collection, _
    TIDCollection As Collection, TLCollection As Collection)
    Dim tableRanges As Variant, tableTableCount As Variant, tableFileName As Variant, tableID As Variant
    Dim tempCurrentTable As String, tempCurrentSheet As Integer
    
    tableRanges = Array("B8", "F8", "J8", "N8", "B19", "F19")
    tableTableCount = Array("C5", "G5", "K5", "O5", "C16", "G16")
    tableFileName = Array("C6", "G6", "K6", "O6", "C17", "G17")
    tableID = Array("D8", "H8", "L8", "P8", "D19", "H19")
    tempCurrentSheet = 6
    tempCurrentTable = Sheet5.Range("D3")

    RangesToCollection RCollection, FNCollection, TIDCollection, TLCollection, tempCurrentTable, _
                        tableRanges, tableTableCount, tableFileName, tableID, tempCurrentSheet
End Sub

Sub RangesToCollection(RCollection As Collection, FNCollection As Collection, TIDCollection As Collection, _
                        TLCollection As Collection, tFolder, tRanges As Variant, tTableCount As Variant, _
                        tFileName As Variant, tID As Variant, tSheet)
    Dim tempCurrentTableFileName As String, tempRanges As Range, tempTableSize As Integer
    Dim tempValueOfTheRange As String, tempValueOfID As String
    

    Sheets(tSheet).Select
    
    For i = 0 To UBound(tRanges)
        Set tempRanges = ActiveSheet.Range(tRanges(i)).CurrentRegion
        Set tempID = ActiveSheet.Range(tID(i)).CurrentRegion
        tempTableSize = ActiveSheet.Range(tTableCount(i))
        tempCurrentTableFileName = tFolder & "\" & ActiveSheet.Range(tFileName(i))
        TLCollection.Add tempTableSize

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