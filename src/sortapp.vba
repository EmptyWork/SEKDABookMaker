Private K9OldVal As Variant
Dim isAutoSave, isNotAutoSave

Sub StartButton_Click()
    Dim mbResult As Integer
    Dim autoSaveCell As Range
    isNotAutoSave = RGB("255", "137", "137")

    Set autoSaveCell = Range("K10")
    Dim currentAutoSaveStatus

    currentAutoSaveStatus = autoSaveCell <> isNotAutoSave

    mbResult = MsgBox("Program akan berjalan dan memakan waktu beberapa menit, " & _
    "kemungkinan besar akan mengganggu pekerjaan anda, harap pastikan pekerjaan anda sudah selesai" & _
    "dan sudah disimpan. Tetap jalankan programnya?", vbYesNo)

    Select Case mbResult
        Case vbYes
            Main currentAutoSaveStatus
        Case vbNo
            Exit Sub
    End Select
End Sub

Public Function K9AlertOnChange(val)
    'If val <> K9OldVal Then MsgBox "Value changed!"
    K9OldVal = val
    K9AlertOnChange = val
End Function

Sub Folder_Path()
    Dim Folder_Picker As FileDialog
    Dim my_path As String
    
    Set Folder_Picker = Application.FileDialog(msoFileDialogFolderPicker)
    Folder_Picker.Title = "Select a Folder"
    Folder_Picker.Filters.Clear
    Folder_Picker.Show
    
    If Folder_Picker.SelectedItems.Count = 1 Then my_path = Folder_Picker.SelectedItems(1)
    If Not IsEmpty(my_path) Then ActiveSheet.Range("D4").Value = my_path
End Sub

Sub File_Path()
    Dim File_Picker As FileDialog
    Dim my_path As String
    
    Set File_Picker = Application.FileDialog(msoFileDialogFilePicker)
    File_Picker.Title = "Select a File" & FileType
    File_Picker.Filters.Clear
    File_Picker.Show
    
    If File_Picker.SelectedItems.Count = 1 Then my_path = File_Picker.SelectedItems(1)
    If Not IsEmpty(my_path) Then ActiveSheet.Range("D5").Value = my_path
End Sub

Sub Auto_Save()
    Dim autoSaveCell As Range
    isNotAutoSave = RGB("255", "137", "137")
    isAutoSave = RGB("193", "252", "140")
    
    Set autoSaveCell = Range("K10")
    
    If autoSaveCell.Interior.Color = isNotAutoSave Then
        autoSaveCell.Interior.Color = isAutoSave
        autoSaveCell.Value = "y"
    Else:
        autoSaveCell.Interior.Color = isNotAutoSave
        autoSaveCell.Value = "n"
    End If
End Sub

Sub GetTableTwoData()
Dim Ranges As Variant

Ranges = Array("B9:B14", "F9:F14")

For i = 0 To 1
   Sheets("Table 2").Range(Ranges(i)).Select
   Application.Wait (Now() + TimeValue("00:00:05"))
Next i
End Sub
