Attribute VB_Name = "modImpayes"
Option Explicit

Sub SuivreImpayes()
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long
    Dim dateRef As Date, retard As Long

    Set ws = ThisWorkbook.Worksheets("Suivi_Factures")
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    dateRef = Date

    ws.Range("I1:J1").Value = Array("Retard (jours)", "Priorité")
    ws.Range("A2:J" & lastRow).Interior.Pattern = xlNone

    For i = 2 To lastRow
        If ws.Cells(i, 6).Value = "Impayée" Then
            retard = DateDiff("d", ws.Cells(i, 4).Value, dateRef)
            If retard < 0 Then retard = 0
            ws.Cells(i, 9).Value = retard

            If retard > 30 Then
                ws.Cells(i, 10).Value = "Relance urgente"
                ws.Range("A" & i & ":J" & i).Interior.Color = RGB(255, 230, 230)
            ElseIf retard > 0 Then
                ws.Cells(i, 10).Value = "Relance"
                ws.Range("A" & i & ":J" & i).Interior.Color = RGB(255, 245, 204)
            Else
                ws.Cells(i, 10).Value = "À venir"
            End If
        Else
            ws.Cells(i, 9).Value = 0
            ws.Cells(i, 10).Value = "Soldée"
        End If
    Next i

    ws.Columns.AutoFit
    MsgBox "Suivi des impayés mis à jour.", vbInformation
End Sub