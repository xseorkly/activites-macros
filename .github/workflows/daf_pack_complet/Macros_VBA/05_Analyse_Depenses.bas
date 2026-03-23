Attribute VB_Name = "modDepenses"
Option Explicit

Sub AnalyserDepenses()
    Dim ws As Worksheet, wsOut As Worksheet
    Dim lastRow As Long, outRow As Long, i As Long
    Dim d As Object, k As Variant

    Set ws = ThisWorkbook.Worksheets("Depenses")
    On Error Resume Next
    Set wsOut = ThisWorkbook.Worksheets("Synthese_Depenses")
    On Error GoTo 0

    If wsOut Is Nothing Then
        Set wsOut = ThisWorkbook.Worksheets.Add(After:=ws)
        wsOut.Name = "Synthese_Depenses"
    Else
        wsOut.Cells.Clear
    End If

    Set d = CreateObject("Scripting.Dictionary")
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    For i = 2 To lastRow
        If d.Exists(CStr(ws.Cells(i, 3).Value)) Then
            d(CStr(ws.Cells(i, 3).Value)) = d(CStr(ws.Cells(i, 3).Value)) + CDbl(ws.Cells(i, 5).Value)
        Else
            d.Add CStr(ws.Cells(i, 3).Value), CDbl(ws.Cells(i, 5).Value)
        End If
    Next i

    wsOut.Range("A1:B1").Value = Array("Catégorie", "Total HT (€)")
    outRow = 2
    For Each k In d.Keys
        wsOut.Cells(outRow, 1).Value = k
        wsOut.Cells(outRow, 2).Value = d(k)
        outRow = outRow + 1
    Next k

    wsOut.Range("D1:E4").Value = Array( _
        Array("Top 3 postes", "Montant"), _
        Array("", ""), _
        Array("", ""), _
        Array("", ""))

    wsOut.Range("A1:B" & outRow - 1).Sort Key1:=wsOut.Range("B2"), Order1:=xlDescending, Header:=xlYes
    wsOut.Range("D2:E4").Value = wsOut.Range("A2:B4").Value
    wsOut.Columns.AutoFit

    MsgBox "Synthèse des dépenses générée.", vbInformation
End Sub