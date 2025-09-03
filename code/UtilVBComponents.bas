Attribute VB_Name = "UtilVBComponents"
Option Explicit

Function removeComponent(compList As Variant, componentName As String)
    
    Dim isCompFound As Boolean
    isCompFound = componentExists2(compList, componentName)
    
    If isCompFound Then
        compList.Remove compList(componentName)
    End If
End Function

Function componentExists1(componentList As Variant, componentName As String) As Boolean
    Dim tempComp As VBComponent
    Dim compName As String
    Dim isAvailable As Boolean
    compName = componentName
    isAvailable = False
    For Each tempComp In componentList
        If LCase(tempComp.Name) = LCase(compName) Then
            isAvailable = True
            Exit For
        End If
    Next tempComp
    componentExists = isAvailable
End Function

Function componentExists2(componentList As Variant, componentName As String) As Boolean
    Dim checkComp As Variant
    On Error GoTo notFound
    Set checkComp = componentList(componentName)
    componentExists2 = True
    Exit Function
notFound:
    componentExists2 = False
End Function

