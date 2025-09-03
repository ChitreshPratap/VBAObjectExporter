VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} formVBAObjectImporter 
   Caption         =   "formVBAObjectImporter"
   ClientHeight    =   9315.001
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   13965
   OleObjectBlob   =   "formVBAObjectImporter.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "formVBAObjectImporter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnImport_Click()
    Dim liBox As Variant
    Dim itemCounter As Integer
    Dim itmName As String
    Dim targetFile As String
    Dim sourceModulePath As String
    Dim destinationPath As String
    Dim wbSource As Workbook
    
    
    targetFile = formVBAObjectImporter.txtTargetFile.Value
    sourceModulePath = Me.txtSourceFolder.Value
    
    If targetFile <> "" And sourceModulePath <> "" Then
        Set wbSource = Workbooks.Open(targetFile)
        Set liBox = formVBAObjectImporter.listBoxModuleToImort
        
        For itemCounter = 1 To liBox.ListCount
            itmName = liBox.List(itemCounter - 1)
            If checkBoxReplaceExistingObjects.Value Then
                UtilVBComponents.removeComponent wbSource.VBProject.VBComponents, CStr(Split(itmName, ".")(0))
            End If
                wbSource.VBProject.VBComponents.Import sourceModulePath & "\" & itmName
            
        Next itemCounter
    
        Set liBox = Me.listBoxClassModuleToImport
        For itemCounter = 1 To liBox.ListCount
            itmName = liBox.List(itemCounter - 1)
            If checkBoxReplaceExistingObjects.Value Then
                UtilVBComponents.removeComponent wbSource.VBProject.VBComponents, CStr(Split(itmName, ".")(0))
            End If

            wbSource.VBProject.VBComponents.Import sourceModulePath & "\" & itmName
        Next itemCounter
    
        Set liBox = Me.listBoxUserFormToImport
        For itemCounter = 1 To liBox.ListCount
            itmName = liBox.List(itemCounter - 1)
            If checkBoxReplaceExistingObjects.Value Then
                UtilVBComponents.removeComponent wbSource.VBProject.VBComponents, CStr(Split(itmName, ".")(0))
            End If
            wbSource.VBProject.VBComponents.Import sourceModulePath & "\" & itmName
        Next itemCounter
        MsgBox "Modules Imported Successfully.", vbOKOnly + vbInformation, "Import Status"
        GoTo finalizeResources
    Else
        MsgBox "Please select targetFile or source folder", vbOKOnly + vbExclamation, "Required Field"
    End If
    
finalizeResources:
    If Not (wbSource Is Nothing) Then
        wbSource.Close True
    End If
End Sub

Private Sub btnSelectFile_Click()
    Dim selectedFileName As String
    selectedFileName = Application.GetOpenFilename("Excel Files (*.xls; *.xlsx; *.xlsm; *.xlsb), *.xls", 1, "Select File to export object", "Select File", False)
    If LCase(selectedFileName) = "false" Then
        selectedFileName = ""
    End If
    formVBAObjectImporter.txtTargetFile.Value = selectedFileName

End Sub

Private Sub btnSelectFolder_Click()
    Dim selectedFolderStatus As Variant
    Dim selectedPath As String
    selectedFolderStatus = Application.FileDialog(msoFileDialogFolderPicker).Show
    If selectedFolderStatus = 0 Then
        formVBAObjectImporter.txtSourceFolder = ""
    Else
            selectedPath = Application.FileDialog(msoFileDialogFolderPicker).SelectedItems(1)
            formVBAObjectImporter.txtSourceFolder.Value = selectedPath
    End If

End Sub

Private Sub CommandButton1_Click()

End Sub

Private Sub listBoxClassModuleName_Click()
    UtilListBox.removeAddItem formVBAObjectImporter.listBoxClassModuleName, formVBAObjectImporter.listBoxClassModuleToImport
End Sub

Private Sub listBoxClassModuleToImport_Click()
    UtilListBox.removeAddItem formVBAObjectImporter.listBoxClassModuleToImport, formVBAObjectImporter.listBoxClassModuleName
End Sub

Private Sub listBoxModuleName_Click()
    UtilListBox.removeAddItem formVBAObjectImporter.listBoxModuleName, formVBAObjectImporter.listBoxModuleToImort
End Sub

Private Sub listBoxModuleToImort_Click()
    UtilListBox.removeAddItem formVBAObjectImporter.listBoxModuleToImort, formVBAObjectImporter.listBoxModuleName
End Sub

Private Sub listBoxUserFormName_Click()
    UtilListBox.removeAddItem formVBAObjectImporter.listBoxUserFormName, formVBAObjectImporter.listBoxUserFormToImport
End Sub

Private Sub listBoxUserFormToImport_Click()
    UtilListBox.removeAddItem formVBAObjectImporter.listBoxUserFormToImport, formVBAObjectImporter.listBoxUserFormName
End Sub

Private Sub txtSourceFolder_Change()
    Dim sourceFolderPath As String
    Dim eachFile As File
    Dim sourceFolder As Folder
    Dim ttype As Variant
    Dim fSplit As Variant
    Dim tFileName As String
    Dim fso As New FileSystemObject
    
    sourceFolderPath = formVBAObjectImporter.txtSourceFolder.Value
    If sourceFolderPath <> "" Then
        Set sourceFolder = fso.GetFolder(sourceFolderPath)
        For Each eachFile In sourceFolder.Files
           Debug.Print eachFile.Name
            fSplit = Split(eachFile, ".")
            tFileName = fSplit(0)
            ttype = fSplit(1)
            If LCase(ttype) = "cls" Then
                formVBAObjectImporter.listBoxClassModuleName.AddItem eachFile.Name
            ElseIf LCase(ttype) = "bas" Then
                formVBAObjectImporter.listBoxModuleName.AddItem eachFile.Name
            ElseIf LCase(ttype) = "frm" Then
                formVBAObjectImporter.listBoxUserFormName.AddItem eachFile.Name
            End If
        Next eachFile
    Else
        MsgBox "Please select source folder path", vbOKOnly + vbExclamation, "Field Required"
    End If
    
End Sub

Private Sub UserForm_Click()

End Sub
