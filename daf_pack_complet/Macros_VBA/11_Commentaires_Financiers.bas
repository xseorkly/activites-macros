Attribute VB_Name = "modCommentaires"
Option Explicit

Sub GenererCommentaires()
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long
    Dim commentaire As String

    Set ws = ThisWorkbook.Worksheets("KPI_Mensuels")
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    ws.Range("H1").Value = "Commentaire automatique"

    For i = 2 To lastRow
        commentaire = ""

        If ws.Cells(i, 2).Value > ws.Cells(i - 1 + IIf(i = 2, 1, 0), 2).Value And i > 2 Then
            commentaire = commentaire & "CA en progression. "
        ElseIf i > 2 And ws.Cells(i, 2).Value < ws.Cells(i - 1, 2).Value Then
            commentaire = commentaire & "CA en retrait. "
        End If

        If ws.Cells(i, 3).Value >= 0.6 Then
            commentaire = commentaire & "Marge brute solide. "
        ElseIf ws.Cells(i, 3).Value < 0.58 Then
            commentaire = commentaire & "Marge sous pression. "
        End If

        If ws.Cells(i, 6).Value < 43000 Then
            commentaire = commentaire & "Niveau de trésorerie à surveiller. "
        End If

        If ws.Cells(i, 7).Value > 47 Then
            commentaire = commentaire & "DSO élevé, renforcer les relances."
        End If

        If commentaire = "" Then commentaire = "Situation stable sur le mois."
        ws.Cells(i, 8).Value = Trim(commentaire)
    Next i

    ws.Columns.AutoFit
    MsgBox "Commentaires financiers générés.", vbInformation
End Sub