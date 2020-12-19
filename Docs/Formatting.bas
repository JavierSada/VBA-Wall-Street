Attribute VB_Name = "Formatting"
Option Explicit

Sub Formatting()
Attribute Formatting.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro1 Macro
Dim ws As Worksheet
For Each ws In ActiveWorkbook.Sheets
    ws.Activate
    Rows("1:1").Select
    Selection.Font.Bold = False
    Selection.Font.Bold = True
    Columns("G:G").Select
    Selection.Style = "Comma"
    Selection.NumberFormat = "_(* #,##0.0_);_(* (#,##0.0);_(* ""-""??_);_(@_)"
    Selection.NumberFormat = "_(* #,##0_);_(* (#,##0);_(* ""-""??_);_(@_)"
    Columns("L:L").Select
    Selection.Style = "Comma"
    Selection.NumberFormat = "_(* #,##0.0_);_(* (#,##0.0);_(* ""-""??_);_(@_)"
    Selection.NumberFormat = "_(* #,##0_);_(* (#,##0);_(* ""-""??_);_(@_)"
    Range("Q4").Select
    Selection.Style = "Comma"
    Selection.NumberFormat = "_(* #,##0.0_);_(* (#,##0.0);_(* ""-""??_);_(@_)"
    Selection.NumberFormat = "_(* #,##0_);_(* (#,##0);_(* ""-""??_);_(@_)"
    Cells.Select
    Cells.EntireColumn.AutoFit
    Range("A1").Select

Next ws

End Sub
