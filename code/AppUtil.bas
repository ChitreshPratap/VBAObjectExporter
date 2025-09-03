Attribute VB_Name = "AppUtil"

Sub main()
    formVBAObjectTransporter.Show
End Sub


Function getSelectedFolder() As String
    Dim selectedFolderStatus As Variant
    Dim selectedPath As String
    selectedFolderStatus = Application.FileDialog(msoFileDialogFolderPicker).Show
    If selectedFolderStatus = 0 Then
        selectedPath = ""
    Else
        selectedPath = Application.FileDialog(msoFileDialogFolderPicker).SelectedItems(1)
    End If
    getSelectedFolder = selectedPath
End Function

Function getSelectedFile() As String
    Dim sourceFileName As String
    sourceFileName = Application.GetOpenFilename("Excel Files (*.xls; *.xlsx; *.xlsm; *.xlsb), *.xls", 1, "Select File to export object", "Select File", False)
    If LCase(sourceFileName) = "false" Then
        sourceFileName = ""
    End If
    getSelectedFile = sourceFileName
End Function

Sub moveAllItems(sourceList As MSForms.listBox, targetList As MSForms.listBox)
    Dim eItm As Variant
    Dim itmNumber As Long
    
    For itmNumber = 0 To sourceList.ListCount - 1
        eItm = UtilListBox.removeItemByIndex(sourceList, 0)
        targetList.AddItem eItm
    Next itmNumber
End Sub

