Attribute VB_Name = "modTresorerie"
Option Explicit

Sub CalculerTresorerie()
    Dim wsFlux As Worksheet, wsParam As Worksheet
    Dim lastRow As Long, i As Long
    Dim solde As Double, seuil As Double
    Dim ch As ChartObject

    Set wsFlux = ThisWorkbook.Worksheets("Flux_Journaliers")
    Set wsParam = ThisWorkbook.Worksheets("Parametres")

    solde = wsParam.Range("B2").Value
    seuil = wsParam.Range("B3").Value
    lastRow = wsFlux.Cells(wsFlux.Rows.Count, "A").End(xlUp).Row

    wsFlux.Range("E1:F1").Value = Array("Solde quotidien (€)", "Alerte")
    wsFlux.Range("E2:F" & wsFlux.Rows.Count).ClearContents
    wsFlux.Range("A2:F" & lastRow).Interior.Pattern = xlNone

    For i = 2 To lastRow
        solde = solde + Nz(wsFlux.Cells(i, 2).Value) - Nz(wsFlux.Cells(i, 3).Value)
        wsFlux.Cells(i, 5).Value = solde

        If solde < seuil Then
            wsFlux.Cells(i, 6).Value = "Alerte"
            wsFlux.Range("A" & i & ":F" & i).Interior.Color = RGB(255, 235, 235)
        Else
            wsFlux.Cells(i, 6).Value = "OK"
        End If
    Next i

    For Each ch In wsFlux.ChartObjects
        ch.Delete
    Next ch

    Set ch = wsFlux.ChartObjects.Add(Left:=450, Top:=20, Width:=460, Height:=260)
    ch.Chart.ChartType = xlLineMarkers
    ch.Chart.SetSourceData Source:=wsFlux.Range("A1:E" & lastRow)
    ch.Chart.SeriesCollection(1).Delete
    ch.Chart.SeriesCollection(1).Delete
    ch.Chart.SeriesCollection(1).Delete
    ch.Chart.HasTitle = True
    ch.Chart.ChartTitle.Text = "Évolution du solde quotidien"

    wsFlux.Columns("A:F").AutoFit
    MsgBox "Calcul et graphique de trésorerie terminés.", vbInformation
End Sub

Private Function Nz(valeur As Variant) As Double
    If IsNumeric(valeur) Then
        Nz = CDbl(valeur)
    Else
        Nz = 0
    End If
End Function