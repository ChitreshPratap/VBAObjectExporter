VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} formVBAObjectTransporter 
   Caption         =   "Application"
   ClientHeight    =   9900.001
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   15705
   OleObjectBlob   =   "formVBAObjectTransporter.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "formVBAObjectTransporter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnExportAll_Click()
    
    AppUtil.moveAllItems Me.listBoxSourceClassModuleName, Me.listBoxTargetClassModuleName
    AppUtil.moveAllItems Me.listBoxSourceModuleName, Me.listBoxTargetModuleName
    AppUtil.moveAllItems Me.listBoxSourceUserFormName, Me.listBoxTargetUserFormName
    AppUtil.moveAllItems Me.listBoxSourcePowerQueryName, Me.listBoxTargetPowerQueryName
    
End Sub

Private Sub btnOIImport_Click()
    
    Dim liBox As Variant
    Dim itemCounter As Integer
    Dim itmName As String
    Dim targetFile As String
    Dim sourceModulePath As String
    Dim destinationPath As String
    Dim wbSource As Workbook
    
    
    targetFile = Me.textBoxOITargetFile.Value
    sourceModulePath = Me.textBoxOISourceFolder.Value
    
    If targetFile <> "" And sourceModulePath <> "" Then
        Set wbSource = Workbooks.Open(targetFile)
        Set liBox = Me.listBoxOITargetModuleName
        
        For itemCounter = 1 To liBox.ListCount
            itmName = liBox.List(itemCounter - 1)
            If checkBoxReplaceExistingObjects.Value Then
                UtilVBComponents.removeComponent wbSource.VBProject.VBComponents, CStr(Split(itmName, ".")(0))
            End If
                wbSource.VBProject.VBComponents.Import sourceModulePath & "\" & itmName
        Next itemCounter
    
        Set liBox = Me.listBoxOITargetClassModuleName
        For itemCounter = 1 To liBox.ListCount
            itmName = liBox.List(itemCounter - 1)
            If checkBoxReplaceExistingObjects.Value Then
                UtilVBComponents.removeComponent wbSource.VBProject.VBComponents, CStr(Split(itmName, ".")(0))
            End If
            wbSource.VBProject.VBComponents.Import sourceModulePath & "\" & itmName
        Next itemCounter
    
        Set liBox = Me.listBoxOITargetFormsName
        For itemCounter = 1 To liBox.ListCount
            itmName = liBox.List(itemCounter - 1)
            If checkBoxReplaceExistingObjects.Value Then
                UtilVBComponents.removeComponent wbSource.VBProject.VBComponents, CStr(Split(itmName, ".")(0))
            End If
            wbSource.VBProject.VBComponents.Import sourceModulePath & "\" & itmName
        Next itemCounter
        
        Set liBox = Me.listBoxOITargetPowerQueryName
        For itemCounter = 1 To liBox.ListCount
            Dim qryName As String
            Dim qryFormula As String
            Dim fileNum As Integer
            Dim line As String
            Dim itmFullPath As String
            Dim tPQ As WorkbookQuery
            
            qryFormula = ""
            itmName = liBox.List(itemCounter - 1)
            qryName = CStr(Split(itmName, ".")(0))
            
            fileNum = FreeFile
            itmFullPath = sourceModulePath & "\" & itmName
            Open itmFullPath For Input As #fileNum
                Do Until EOF(fileNum)
                    Line Input #fileNum, line
                    qryFormula = qryFormula & line & vbCrLf
                Loop
            Close #fileNum
            
            If checkBoxReplaceExistingObjects.Value Then
                On Error Resume Next
                Set tPQ = wbSource.Queries(qryName)
                If Not tPQ Is Nothing Then tPQ.Delete
            Else
                On Error Resume Next
                Set tPQ = wbSource.Queries(qryName)
                If Not tPQ Is Nothing Then
                    qryName = qryName & "_" & Format(Now(), "yyyyMMddhhmmss")
                End If
            End If
            
            wbSource.Queries.Add Name:=qryName, Formula:=qryFormula
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

Private Sub btnOIImportAll_Click()

    AppUtil.moveAllItems Me.listBoxOISourceClassModuleName, Me.listBoxOITargetClassModuleName
    AppUtil.moveAllItems Me.listBoxOISourceFormsName, Me.listBoxOITargetFormsName
    AppUtil.moveAllItems Me.listBoxOISourceModuleName, Me.listBoxOITargetModuleName
    AppUtil.moveAllItems Me.listBoxOISourcePowerQueryName, Me.listBoxOITargetPowerQueryName

End Sub

Private Sub btnOIRemoveAll_Click()

    AppUtil.moveAllItems Me.listBoxOITargetClassModuleName, Me.listBoxOISourceClassModuleName
    AppUtil.moveAllItems Me.listBoxOITargetFormsName, Me.listBoxOISourceFormsName
    AppUtil.moveAllItems Me.listBoxOITargetModuleName, Me.listBoxOISourceModuleName
    AppUtil.moveAllItems Me.listBoxOITargetPowerQueryName, Me.listBoxOISourcePowerQueryName

End Sub

Private Sub btnOISelectFolder_Click()

    Dim selectedPath As String
    selectedPath = AppUtil.getSelectedFolder()
    Me.textBoxOISourceFolder = selectedPath
    
End Sub

Private Sub btnOISelectTargetFile_Click()
    Dim targetFileName As String
    targetFileName = AppUtil.getSelectedFile()
    Me.textBoxOITargetFile = targetFileName

End Sub

Private Sub btnRemoveAll_Click()

    AppUtil.moveAllItems Me.listBoxTargetClassModuleName, Me.listBoxSourceClassModuleName
    AppUtil.moveAllItems Me.listBoxTargetModuleName, Me.listBoxSourceModuleName
    AppUtil.moveAllItems Me.listBoxTargetUserFormName, Me.listBoxSourceUserFormName
    AppUtil.moveAllItems Me.listBoxTargetPowerQueryName, Me.listBoxSourcePowerQueryName

End Sub


Private Sub btnSelectFileOI_Click()

End Sub

Private Sub btnSelectFolder_Click()
    Dim selectedPath As String
    selectedPath = AppUtil.getSelectedFolder()
    Me.txtDestinationFolder = selectedPath
End Sub

Private Sub btnSelectSourceFile_Click()
    Dim sourceFileName As String
    sourceFileName = AppUtil.getSelectedFile()
    Me.txtSelectedFileName = sourceFileName
End Sub

Private Sub Frame1_Click()

End Sub

Private Sub frameObjectExporter_Click()

End Sub

Private Sub lblUserName_Click()

End Sub

Private Sub lblWelcome_Click()

End Sub

Private Sub listBoxOISourceClassModuleName_Click()
    UtilListBox.removeAddItem Me.listBoxOISourceClassModuleName, Me.listBoxOITargetClassModuleName
End Sub

Private Sub listBoxOISourceFormsName_Click()
    UtilListBox.removeAddItem Me.listBoxOISourceFormsName, Me.listBoxOITargetFormsName
End Sub

Private Sub listBoxOISourceModuleName_Click()
    UtilListBox.removeAddItem Me.listBoxOISourceModuleName, Me.listBoxOITargetModuleName
End Sub

Private Sub listBoxOISourcePowerQueryName_Click()
    UtilListBox.removeAddItem Me.listBoxOISourcePowerQueryName, Me.listBoxOITargetPowerQueryName
End Sub

Private Sub listBoxOITargetClassModuleName_Click()
    UtilListBox.removeAddItem Me.listBoxOITargetClassModuleName, Me.listBoxOISourceClassModuleName
End Sub

Private Sub listBoxOITargetFormsName_Click()
    UtilListBox.removeAddItem Me.listBoxOITargetFormsName, Me.listBoxOISourceFormsName
End Sub

Private Sub listBoxOITargetModuleName_Click()
    UtilListBox.removeAddItem Me.listBoxOITargetModuleName, Me.listBoxOISourceModuleName
End Sub

Private Sub listBoxOITargetPowerQueryName_Click()
    UtilListBox.removeAddItem Me.listBoxOITargetPowerQueryName, Me.listBoxOISourcePowerQueryName
End Sub

Private Sub listBoxSourceClassModuleName_Click()
    UtilListBox.removeAddItem Me.listBoxSourceClassModuleName, Me.listBoxTargetClassModuleName
End Sub

Private Sub listBoxSourcePowerQueryName_Click()
    UtilListBox.removeAddItem Me.listBoxSourcePowerQueryName, Me.listBoxTargetPowerQueryName
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

Private Sub listBoxTargetPowerQueryName_Click()
    UtilListBox.removeAddItem Me.listBoxTargetPowerQueryName, Me.listBoxSourcePowerQueryName
End Sub

Private Sub listBoxTargetModuleName_Click()
    UtilListBox.removeAddItem Me.listBoxTargetModuleName, Me.listBoxSourceModuleName
End Sub

Private Sub listBoxTargetUserFormName_Click()
    UtilListBox.removeAddItem Me.listBoxTargetUserFormName, Me.listBoxSourceUserFormName
End Sub

Private Sub MultiPage1_Change()

End Sub

Private Sub textBoxOISourceFolder_Change()

    Dim sourceFolderPath As String
    Dim eachFile As File
    Dim sourceFolder As Folder
    Dim ttype As Variant
    Dim fSplit As Variant
    Dim tFileName As String
    Dim fso As New FileSystemObject
    
    Me.clearAllObjectImportLists
    sourceFolderPath = Me.textBoxOISourceFolder.Value
    If sourceFolderPath <> "" Then
        Set sourceFolder = fso.GetFolder(sourceFolderPath)
        For Each eachFile In sourceFolder.Files
            fSplit = Split(eachFile, ".")
            tFileName = fSplit(0)
            ttype = fSplit(1)
            If LCase(ttype) = "cls" Then
                Me.listBoxOISourceClassModuleName.AddItem eachFile.Name
            ElseIf LCase(ttype) = "bas" Then
                Me.listBoxOISourceModuleName.AddItem eachFile.Name
            ElseIf LCase(ttype) = "frm" Then
                Me.listBoxOISourceFormsName.AddItem eachFile.Name
            ElseIf LCase(ttype) = "pq" Then
                Me.listBoxOISourcePowerQueryName.AddItem eachFile.Name
            End If
        Next eachFile
    Else
        MsgBox "Please select source folder path", vbOKOnly + vbExclamation, "Field Required"
    End If

End Sub

Private Sub UserForm_Initialize()
    UtilListBox.CreateListBoxHeader Me.listBoxSourceModuleName, Me.listBoxSourceModuleNameHeading, Array("Module")
    UtilListBox.CreateListBoxHeader Me.listBoxSourceClassModuleName, Me.listBoxSourceClassModuleNameHeading, Array("Class Module")
    UtilListBox.CreateListBoxHeader Me.listBoxSourceUserFormName, Me.listBoxSourceUserFormNameHeading, Array("Forms")
    UtilListBox.CreateListBoxHeader Me.listBoxSourcePowerQueryName, Me.listBoxSourcePowerQueryNameHeading, Array("Query")
    UtilListBox.CreateListBoxHeader Me.listBoxSourceObjName, Me.listBoxSourceObjNameHeading, Array("Excel Objects")
    
    UtilListBox.CreateListBoxHeader Me.listBoxTargetModuleName, Me.listBoxTargetModuleNameHeading, Array("Module")
    UtilListBox.CreateListBoxHeader Me.listBoxTargetClassModuleName, Me.listBoxTargetClassModuleNameHeading, Array("Class Module")
    UtilListBox.CreateListBoxHeader Me.listBoxTargetPowerQueryName, Me.listBoxTargetPowerQueryHeading, Array("Query")
    UtilListBox.CreateListBoxHeader Me.listBoxTargetUserFormName, Me.listBoxTargetUserFormNameHeading, Array("Forms")


    UtilListBox.CreateListBoxHeader Me.listBoxOISourceModuleName, Me.listBoxOISourceModuleNameHeading, Array("Module")
    UtilListBox.CreateListBoxHeader Me.listBoxOISourceClassModuleName, Me.listBoxOISourceClassModuleNameHeading, Array("Class Module")
    UtilListBox.CreateListBoxHeader Me.listBoxOISourceFormsName, Me.listBoxOISourceFormsNameHeading, Array("Forms")
    UtilListBox.CreateListBoxHeader Me.listBoxOISourcePowerQueryName, Me.listBoxOISourcePowerQueryHeading, Array("Query")
    
    UtilListBox.CreateListBoxHeader Me.listBoxOITargetModuleName, Me.listBoxOITargetModuleNameHeading, Array("Module")
    UtilListBox.CreateListBoxHeader Me.listBoxOITargetClassModuleName, Me.listBoxOITargetClassModuleNameHeading, Array("Class Module")
    UtilListBox.CreateListBoxHeader Me.listBoxOITargetFormsName, Me.listBoxOITargetFormsNameHeading, Array("Forms")
    UtilListBox.CreateListBoxHeader Me.listBoxOITargetPowerQueryName, Me.listBoxOITargetPowerQueryHeading, Array("Query")
    Me.lblUserName = "Welcome " & Application.UserName
    Me.MultiPage1.Value = 0

End Sub

Private Sub txtSelectedFileName_Change()
    
    Dim sourceFileName As String
    Dim wbSource As Workbook
    Dim comp As Variant
    Dim qry As WorkbookQuery
                
    Me.clearAllObjectExportLists
    sourceFileName = Me.txtSelectedFileName.Value
    
    If sourceFileName <> "" Then
        Set wbSource = Workbooks.Open(sourceFileName, False, True)
        
        For Each qry In wbSource.Queries
            Me.listBoxSourcePowerQueryName.AddItem qry.Name
        Next qry
        
        For Each comp In wbSource.VBProject.VBComponents
            If comp.Type = 1 Then
                Me.listBoxSourceModuleName.AddItem comp.Name
            ElseIf comp.Type = 2 Then
                Me.listBoxSourceClassModuleName.AddItem comp.Name
            ElseIf comp.Type = 3 Then
                Me.listBoxSourceUserFormName.AddItem comp.Name
            ElseIf comp.Type = 100 Then
                Me.listBoxSourceObjName.AddItem comp.Name
            End If
        Next comp
        
    End If
    
finalizeResource:
    If Not (wbSource Is Nothing) Then
        wbSource.Close False
    End If
    
End Sub


Sub clearAllObjectExportLists()

    Me.listBoxSourceModuleName.Clear
    Me.listBoxSourceClassModuleName.Clear
    Me.listBoxSourceUserFormName.Clear
    Me.listBoxSourceObjName.Clear
    Me.listBoxTargetClassModuleName.Clear
    Me.listBoxTargetModuleName.Clear
    Me.listBoxTargetUserFormName.Clear
    
End Sub

Sub clearAllObjectImportLists()

    Me.listBoxOISourceClassModuleName.Clear
    Me.listBoxOISourceFormsName.Clear
    Me.listBoxOISourceModuleName.Clear
    
    Me.listBoxOITargetClassModuleName.Clear
    Me.listBoxOITargetFormsName.Clear
    Me.listBoxOITargetModuleName.Clear
    
End Sub


Private Sub btnExport_Click()
    
    Dim liBox As Variant
    Dim itemCounter As Integer
    Dim itmName As String
    Dim destinationPath As String
    Dim sourceFileName As String
    Dim wbSource As Workbook
    Dim fso As FileSystemObject
    Set fso = New FileSystemObject
    
    sourceFileName = Me.txtSelectedFileName.Value
    destinationPath = Me.txtDestinationFolder.Value
    
    Set wbSource = Workbooks.Open(sourceFileName)
        
        
    
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
        
        Set liBox = Me.listBoxTargetPowerQueryName
        For itemCounter = 1 To liBox.ListCount
            Dim qry As WorkbookQuery
            Dim exportQryPath As String
            Dim ts As TextStream
            
            itmName = liBox.List(itemCounter - 1)
            exportQryPath = destinationPath & "\" & itmName & ".pq"
            Set qry = wbSource.Queries(itmName)
            Set ts = fso.CreateTextFile(exportQryPath, True)
            'ts.WriteLine "Query Name : " & itmName
            ts.WriteLine qry.Formula
            ts.Close
        Next itemCounter
                
        MsgBox "Modules Exported Successfully", vbOKOnly + vbInformation, "Export Status"
        Exit Sub
    Else
        MsgBox "Please select destination path", vbOKOnly + vbExclamation, "Required Field"
    End If
    
    
    
finalizeResource:
    If Not (wbSource Is Nothing) Then
        wbSource.Close False
    End If
    
End Sub

