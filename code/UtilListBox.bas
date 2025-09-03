Attribute VB_Name = "UtilListBox"
Option Explicit

Public Function removeItemByIndex(sourceList As MSForms.listBox, indexToRemove As Long, Optional ignoreErrorIfOutOfIndex As Boolean = True) As Variant
    
    Dim selectedIndex As Integer
    Dim targetItem As Variant
                
    If indexToRemove > (sourceList.ListCount - 1) Or indexToRemove < 0 Then
        If ignoreErrorIfOutOfIndex Then
            Exit Function
        Else
            Err.Raise 1380, "UtilListBox.removeItemByIndex", "Error : (380 + 1000)" & vbNewLine & "IndexOutOfRangeError : Index is out of range. Could not set the selected property. Invalid property value."
        End If
    End If
    targetItem = sourceList.List(indexToRemove)
    sourceList.Selected(indexToRemove) = False
    sourceList.RemoveItem indexToRemove
    removeItemByIndex = targetItem
    
End Function


Public Function removeSelectedItem(sourceList As MSForms.listBox) As Variant

    Dim selectedIndex As Integer
    Dim targetItem As Variant
    selectedIndex = sourceList.ListIndex
    targetItem = UtilListBox.removeItemByIndex(selectedIndex)
    removeSelectedItem = targetItem
    
End Function

Public Sub removeAddItem(sourceList As Variant, targetList As Variant)
    Dim selectedIndex As Integer
    targetList.AddItem sourceList.List(sourceList.ListIndex)
    selectedIndex = sourceList.ListIndex
    sourceList.Selected(selectedIndex) = False
    sourceList.RemoveItem selectedIndex
End Sub

Public Sub CreateListBoxHeader(body As MSForms.listBox, header As MSForms.listBox, arrHeaders)
    ' make column count match
    header.ColumnCount = body.ColumnCount
    header.ColumnWidths = body.ColumnWidths

    ' add header elements
    header.Clear
    header.AddItem
    Dim i As Integer
    For i = 0 To UBound(arrHeaders)
        header.List(0, i) = arrHeaders(i)
    Next i

    ' make it pretty
    body.ZOrder (1)
    header.ZOrder (0)
    header.SpecialEffect = fmSpecialEffectFlat
    header.BackColor = RGB(200, 200, 200)

    ' align header to body (should be done last!)
    header.Width = body.Width
    header.Left = body.Left
    header.Top = body.Top - (header.Height - 1)
End Sub
