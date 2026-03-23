Attribute VB_Name = "modAnomalies"
Option Explicit

Sub DetecterAnomalies()
    Dim ws As Worksheet, wsOut As Worksheet
    Dim lastRow As Long, outRow As Long, i As Long
    Dim cle As String
    Dim d As Object

    Set ws = ThisWorkbook.Worksheets("Ecritures")
    On Error Resume Next
    Set wsOut = ThisWorkbook.Worksheets("Anomalies")
    On Error GoTo 0

    If wsOut Is Nothing Then
        Set wsOut = ThisWorkbook.Worksheets.Add(After:=ws)
        wsOut.Name = "Anomalies"
    Else
        wsOut.Cells.Clear
    End If

    wsOut.Range("A1:F1").Value = Array("Ligne source", "Type anomalie", "Date", "Compte", "NumPiece", "Montant (€)")
    outRow = 2
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    Set d = CreateObject("Scripting.Dictionary")

    For i = 2 To lastRow
        cle = CStr(ws.Cells(i, 1).Value) & "|" & CStr(ws.Cells(i, 3).Value) & "|" & CStr(ws.Cells(i, 5).Value) & "|" & CStr(ws.Cells(i, 6).Value)

        If d.Exists(cle) Then
            AjouterAnomalie ws, wsOut, outRow, i, "Doublon détecté"
        Else
            d.Add cle, i
        End If

        If ws.Cells(i, 5).Value = "" Then
            AjouterAnomalie ws, wsOut, outRow, i, "Numéro de pièce manquant"
        End If

        If Abs(CDbl(Nz(ws.Cells(i, 6).Value))) >= 50000 Then
            AjouterAnomalie ws, wsOut, outRow, i, "Montant élevé à contrôler"
        End If

        If ws.Cells(i, 2).Value = "ACH" And CDbl(Nz(ws.Cells(i, 6).Value)) < 0 Then
            AjouterAnomalie ws, wsOut, outRow, i, "Charge d'achat négative"
        End If
    Next i

    wsOut.Columns.AutoFit
    MsgBox "Contrôle terminé. Consultez l'onglet Anomalies.", vbInformation
End Sub

Private Sub AjouterAnomalie(ws As Worksheet, wsOut As Worksheet, ByRef outRow As Long, ByVal i As Long, ByVal motif As String)
    wsOut.Cells(outRow, 1).Value = i
    wsOut.Cells(outRow, 2).Value = motif
    wsOut.Cells(outRow, 3).Value = ws.Cells(i, 1).Value
    wsOut.Cells(outRow, 4).Value = ws.Cells(i, 3).Value
    wsOut.Cells(outRow, 5).Value = ws.Cells(i, 5).Value
    wsOut.Cells(outRow, 6).Value = ws.Cells(i, 6).Value
    outRow = outRow + 1
End Sub

Private Function Nz(valeur As Variant) As Double
    If IsNumeric(valeur) Then
        Nz = CDbl(valeur)
    Else
        Nz = 0
    End If
End Function