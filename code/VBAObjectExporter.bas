Attribute VB_Name = "VBAObjectExporter"
Option Explicit

Sub showFormVBAExporter()
    formVBAObjectExporter.Show
End Sub

Sub showFormVBAImporter()
    formVBAObjectImporter.Show
End Sub


Sub test1()
    Dim wbPath As String
    Dim wb As Workbook
    Dim comp As Variant
    wbPath = "C:\Users\pc\Downloads\Example1.xlsx"
    Set wb = Workbooks.Open(wbPath)
    
    Debug.Print wb.Queries.Count

End Sub
