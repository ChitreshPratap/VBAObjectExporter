VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} formVBAObjectExporter 
   Caption         =   "VBA Objects Exporter"
   ClientHeight    =   9300.001
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9600.001
   OleObjectBlob   =   "formVBAObjectExporter.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "formVBAObjectExporter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim wbSource As Workbook

Private Sub btnExportAll_Click()

    AppUtil.moveAllItems Me.listBoxSourceClassModuleName, Me.listBoxTargetClassModuleName
    AppUtil.moveAllItems Me.listBoxSourceModuleName, Me.listBoxTargetModuleName
    AppUtil.moveAllItems Me.listBoxSourceUserFormName, Me.listBoxTargetUserFormName
    
End Sub

Private Sub btnRemoveAll_Click()

    AppUtil.moveAllItems Me.listBoxTargetClassModuleName, Me.listBoxSourceClassModuleName
    AppUtil.moveAllItems Me.listBoxTargetModuleName, Me.listBoxSourceModuleName
    AppUtil.moveAllItems Me.listBoxTargetUserFormName, Me.listBoxSourceUserFormName
    
End Sub

Private Sub btnSelectFolder_Click()

    Dim selectedPath As String
    selectedPath = AppUtil.getSelectedFolder()
    formVBAObjectExporter.txtDestinationFolder = selectedPath

End Sub

Private Sub btnSelectSourceFile_Click()
    Dim sourceFileName As String
    sourceFileName = AppUtil.getSelectedFile()
    formVBAObjectExporter.txtSelectedFileName = sourceFileName
End Sub

Private Sub formVBAObjectExporter_Click()

End Sub

Private Sub listBoxSourceClassModuleName_Click()
    UtilListBox.removeAddItem Me.listBoxSourceClassModuleName, Me.listBoxTargetClassModuleName
End Sub

Private Sub listBoxSourceModuleName_Click()
   UtilListBox.removeAddItem Me.listBoxSourceModuleName, Me.listBoxTargetModuleName
End Sub

Private Sub listBoxSourceObjName_Click()
    UtilListBox.removeAddItem Me.listBoxSourceObjName, Me.listBoxTargetObjectName
End Sub

Private Sub listBoxSourceUserFormName_Click()
    UtilListBox.removeAddItem Me.listBoxSourceUserFormName, Me.listBoxTargetUserFormName
End Sub

Private Sub listBoxTargetClassModuleName_Click()
    UtilListBox.removeAddItem Me.listBoxTargetClassModuleName, Me.listBoxSourceClassModuleName
End Sub

Private Sub listBoxTargetModuleName_Click()
    UtilListBox.removeAddItem Me.listBoxTargetModuleName, Me.listBoxSourceModuleName
End Sub

Private Sub listBoxTargetUserFormName_Click()
    UtilListBox.removeAddItem Me.listBoxTargetUserFormName, Me.listBoxSourceUserFormName
End Sub

Private Sub UserForm_Initialize()
        
    UtilListBox.CreateListBoxHeader Me.listBoxSourceModuleName, Me.listBoxSourceModuleNameHeading, Array("Module")
    UtilListBox.CreateListBoxHeader Me.listBoxSourceClassModuleName, Me.listBoxSourceClassModuleNameHeading, Array("Class Module")
    UtilListBox.CreateListBoxHeader Me.listBoxSourceUserFormName, Me.listBoxSourceUserFormNameHeading, Array("Forms")
    UtilListBox.CreateListBoxHeader Me.listBoxSourceObjName, Me.listBoxSourceObjNameHeading, Array("Excel Objects")
    
    UtilListBox.CreateListBoxHeader Me.listBoxTargetModuleName, Me.listBoxTargetModuleNameHeading, Array("Module")
    UtilListBox.CreateListBoxHeader Me.listBoxTargetClassModuleName, Me.listBoxTargetClassModuleNameHeading, Array("Class Module")
    UtilListBox.CreateListBoxHeader Me.listBoxTargetUserFormName, Me.listBoxTargetUserFormNameHeading, Array("Forms")
    
End Sub



Private Sub btnExport_Click()
    Dim liBox As Variant
    Dim itemCounter As Integer
    Dim itmName As String
    Dim destinationPath As String
    destinationPath = Me.txtDestinationFolder.Value
    
    If destinationPath <> "" Then
        Set liBox = Me.listBoxTargetModuleName
        For itemCounter = 1 To liBox.ListCount
            itmName = liBox.List(itemCounter - 1)
            wbSource.VBProject.VBComponents(itmName).Export destinationPath & "\" & itmName & ".bas"
        Next itemCounter
    
        Set liBox = Me.listBoxTargetClassModuleName
        For itemCounter = 1 To liBox.ListCount
            itmName = liBox.List(itemCounter - 1)
            wbSource.VBProject.VBComponents(itmName).Export destinationPath & "\" & itmName & ".cls"
        Next itemCounter
    
        Set liBox = Me.listBoxTargetUserFormName
        For itemCounter = 1 To liBox.ListCount
            itmName = liBox.List(itemCounter - 1)
            wbSource.VBProject.VBComponents(itmName).Export destinationPath & "\" & itmName & ".frm"
        Next itemCounter
        MsgBox "Modules Exported Successfully", vbOKOnly + vbInformation, "Export Status"
        Exit Sub
    Else
        MsgBox "Please select destination path", vbOKOnly + vbExclamation, "Required Field"
    End If
    
End Sub


Private Sub txtSelectedFileName_Change()
    Dim sourceFileName As String
    Dim wbSourceFile As Workbook
    Dim comp As Variant
    
    sourceFileName = formVBAObjectExporter.txtSelectedFileName.Text
    If sourceFileName <> "" Then
        Set wbSourceFile = Workbooks.Open(sourceFileName, False, True)
        Set wbSource = wbSourceFile
        For Each comp In wbSourceFile.VBProject.VBComponents
            If comp.Type = 1 Then
                formVBAObjectExporter.listBoxSourceModuleName.AddItem comp.Name
            ElseIf comp.Type = 2 Then
                formVBAObjectExporter.listBoxSourceClassModuleName.AddItem comp.Name
            ElseIf comp.Type = 3 Then
                formVBAObjectExporter.listBoxSourceUserFormName.AddItem comp.Name
            ElseIf comp.Type = 100 Then
                formVBAObjectExporter.listBoxSourceObjName.AddItem comp.Name
            End If
        Next comp
    End If
    
End Sub



Private Sub UserForm_Terminate()
    On Error Resume Next
    wbSource.Close False
End Sub
