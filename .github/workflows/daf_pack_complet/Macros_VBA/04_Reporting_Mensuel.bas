Attribute VB_Name = "modReportingMensuel"
Option Explicit

Sub GenererReportingMensuel()
    Dim wsData As Worksheet, wsRep As Worksheet
    Dim lastRow As Long, i As Long
    Dim totalCA As Double, totalCV As Double, totalCF As Double
    Dim ebitda As Double, marge As Double, budgetCA As Double, budgetEBITDA As Double
    Dim ch As ChartObject

    Set wsData = ThisWorkbook.Worksheets("P&L_Mensuel")
    On Error Resume Next
    Set wsRep = ThisWorkbook.Worksheets("Reporting")
    On Error GoTo 0

    If wsRep Is Nothing Then
        Set wsRep = ThisWorkbook.Worksheets.Add(After:=wsData)
        wsRep.Name = "Reporting"
    Else
        wsRep.Cells.Clear
    End If

    lastRow = wsData.Cells(wsData.Rows.Count, "A").End(xlUp).Row

    For i = 2 To lastRow
        totalCA = totalCA + wsData.Cells(i, 2).Value
        totalCV = totalCV + wsData.Cells(i, 3).Value
        totalCF = totalCF + wsData.Cells(i, 4).Value
        budgetCA = budgetCA + wsData.Cells(i, 5).Value
        budgetEBITDA = budgetEBITDA + wsData.Cells(i, 6).Value
    Next i

    marge = totalCA - totalCV
    ebitda = marge - totalCF

    wsRep.Range("A1:B8").Value = _
        Array( _
        Array("Indicateur", "Valeur"), _
        Array("CA total (€)", totalCA), _
        Array("Coûts variables (€)", totalCV), _
        Array("Marge brute (€)", marge), _
        Array("Charges fixes (€)", totalCF), _
        Array("EBITDA (€)", ebitda), _
        Array("Écart CA vs budget (€)", totalCA - budgetCA), _
        Array("Écart EBITDA vs budget (€)", ebitda - budgetEBITDA))

    wsRep.Range("D1:F" & lastRow).ClearContents
    wsRep.Range("D1:F1").Value = Array("Mois", "CA réel", "EBITDA")
    For i = 2 To lastRow
        wsRep.Cells(i, 4).Value = wsData.Cells(i, 1).Value
        wsRep.Cells(i, 5).Value = wsData.Cells(i, 2).Value
        wsRep.Cells(i, 6).Value = wsData.Cells(i, 2).Value - wsData.Cells(i, 3).Value - wsData.Cells(i, 4).Value
    Next i

    For Each ch In wsRep.ChartObjects
        ch.Delete
    Next ch
    Set ch = wsRep.ChartObjects.Add(Left:=320, Top:=20, Width:=470, Height:=260)
    ch.Chart.ChartType = xlColumnClustered
    ch.Chart.SetSourceData wsRep.Range("D1:F" & lastRow)
    ch.Chart.HasTitle = True
    ch.Chart.ChartTitle.Text = "CA et EBITDA mensuels"

    wsRep.Columns("A:F").AutoFit
    MsgBox "Reporting mensuel généré.", vbInformation
End Sub