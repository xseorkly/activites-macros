Attribute VB_Name = "modAuditDonnees"
Option Explicit

Sub AuditerBaseClients()
    Dim ws As Worksheet, wsOut As Worksheet
    Dim lastRow As Long, outRow As Long, i As Long
    Dim d As Object

    Set ws = ThisWorkbook.Worksheets("Base_Clients")
    On Error Resume Next
    Set wsOut = ThisWorkbook.Worksheets("Audit")
    On Error GoTo 0

    If wsOut Is Nothing Then
        Set wsOut = ThisWorkbook.Worksheets.Add(After:=ws)
        wsOut.Name = "Audit"
    Else
        wsOut.Cells.Clear
    End If

    wsOut.Range("A1:D1").Value = Array("Ligne", "Anomalie", "Client", "Détail")
    outRow = 2
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    Set d = CreateObject("Scripting.Dictionary")

    For i = 2 To lastRow
        If ws.Cells(i, 3).Value = "" Or InStr(1, ws.Cells(i, 3).Value, "@") = 0 Then
            AuditLigne ws, wsOut, outRow, i, "Email invalide", ws.Cells(i, 3).Value
        End If

        If Len(CStr(ws.Cells(i, 4).Value)) <> 14 Then
            AuditLigne ws, wsOut, outRow, i, "SIRET invalide", ws.Cells(i, 4).Value
        End If

        If CDbl(Nz(ws.Cells(i, 5).Value)) < 0 Then
            AuditLigne ws, wsOut, outRow, i, "Encours négatif", ws.Cells(i, 5).Value
        End If

        If d.Exists(CStr(ws.Cells(i, 1).Value) & "|" & CStr(ws.Cells(i, 4).Value)) Then
            AuditLigne ws, wsOut, outRow, i, "Doublon client", ws.Cells(i, 1).Value
        Else
            d.Add CStr(ws.Cells(i, 1).Value) & "|" & CStr(ws.Cells(i, 4).Value), i
        End If
    Next i

    wsOut.Columns.AutoFit
    MsgBox "Audit terminé. Consultez l'onglet Audit.", vbInformation
End Sub

Private Sub AuditLigne(ws As Worksheet, wsOut As Worksheet, ByRef outRow As Long, ByVal i As Long, ByVal anomalie As String, ByVal detail As Variant)
    wsOut.Cells(outRow, 1).Value = i
    wsOut.Cells(outRow, 2).Value = anomalie
    wsOut.Cells(outRow, 3).Value = ws.Cells(i, 1).Value
    wsOut.Cells(outRow, 4).Value = detail
    outRow = outRow + 1
End Sub

Private Function Nz(valeur As Variant) As Double
    If IsNumeric(valeur) Then
        Nz = CDbl(valeur)
    Else
        Nz = 0
    End If
End Function