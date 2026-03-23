Attribute VB_Name = "modRentabilite"
Option Explicit

Sub AnalyserRentabilite()
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long
    Dim margeUnitaire As Double, caNet As Double, margeTotale As Double

    Set ws = ThisWorkbook.Worksheets("Produits")
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    ws.Range("H1:K1").Value = Array("CA net (€)", "Marge unitaire (€)", "Marge totale (€)", "Diagnostic")
    ws.Range("A2:K" & lastRow).Interior.Pattern = xlNone

    For i = 2 To lastRow
        caNet = ws.Cells(i, 3).Value * ws.Cells(i, 4).Value * (1 - ws.Cells(i, 7).Value)
        margeUnitaire = ws.Cells(i, 4).Value - ws.Cells(i, 5).Value
        margeTotale = (ws.Cells(i, 3).Value * margeUnitaire * (1 - ws.Cells(i, 7).Value)) - ws.Cells(i, 6).Value

        ws.Cells(i, 8).Value = caNet
        ws.Cells(i, 9).Value = margeUnitaire
        ws.Cells(i, 10).Value = margeTotale

        If margeTotale < 0 Then
            ws.Cells(i, 11).Value = "Destructeur de valeur"
            ws.Range("A" & i & ":K" & i).Interior.Color = RGB(255, 230, 230)
        ElseIf margeTotale < 25000 Then
            ws.Cells(i, 11).Value = "Sous surveillance"
            ws.Range("A" & i & ":K" & i).Interior.Color = RGB(255, 245, 204)
        Else
            ws.Cells(i, 11).Value = "Rentable"
            ws.Range("A" & i & ":K" & i).Interior.Color = RGB(230, 255, 230)
        End If
    Next i

    ws.Columns.AutoFit
    MsgBox "Analyse de rentabilité terminée.", vbInformation
End Sub