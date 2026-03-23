Attribute VB_Name = "modPrevisions"
Option Explicit

Sub PrevoirCA6Mois()
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long
    Dim pente As Double, moyenneInc As Double
    Dim dernierCA As Double, prochaineDate As Date

    Set ws = ThisWorkbook.Worksheets("Historique_CA")
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    ws.Range("F1:G1").Value = Array("Mois prévision", "CA prévisionnel (€)")

    moyenneInc = (ws.Cells(lastRow, 2).Value - ws.Cells(2, 2).Value) / (lastRow - 2)
    dernierCA = ws.Cells(lastRow, 2).Value
    prochaineDate = DateSerial(Year(ws.Cells(lastRow, 1).Value), Month(ws.Cells(lastRow, 1).Value) + 1, 1)

    For i = 1 To 6
        ws.Cells(i + 1, 6).Value = prochaineDate
        ws.Cells(i + 1, 7).Value = dernierCA + moyenneInc * i
        prochaineDate = DateSerial(Year(prochaineDate), Month(prochaineDate) + 1, 1)
    Next i

    ws.Columns.AutoFit
    MsgBox "Prévisions à 6 mois générées.", vbInformation
End Sub