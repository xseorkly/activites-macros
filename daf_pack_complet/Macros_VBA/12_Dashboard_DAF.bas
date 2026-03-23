Attribute VB_Name = "modDashboard"
Option Explicit

Sub ConstruireDashboard()
    Dim wsKPI As Worksheet, wsBU As Worksheet, wsDash As Worksheet
    Dim lastRow As Long, lastRowBU As Long
    Dim ca As Double, marge As Double, ebitda As Double, cash As Double, dso As Double

    Set wsKPI = ThisWorkbook.Worksheets("Donnees_KPI")
    Set wsBU = ThisWorkbook.Worksheets("Donnees_BU")

    On Error Resume Next
    Set wsDash = ThisWorkbook.Worksheets("Dashboard")
    On Error GoTo 0

    If wsDash Is Nothing Then
        Set wsDash = ThisWorkbook.Worksheets.Add(Before:=wsKPI)
        wsDash.Name = "Dashboard"
    Else
        wsDash.Cells.Clear
    End If

    lastRow = wsKPI.Cells(wsKPI.Rows.Count, "A").End(xlUp).Row
    ca = wsKPI.Cells(lastRow, 2).Value
    marge = wsKPI.Cells(lastRow, 2).Value - wsKPI.Cells(lastRow, 3).Value
    ebitda = marge - wsKPI.Cells(lastRow, 4).Value
    cash = wsKPI.Cells(lastRow, 5).Value
    dso = wsKPI.Cells(lastRow, 6).Value

    wsDash.Range("A1:B6").Value = Array( _
        Array("KPI", "Valeur"), _
        Array("CA dernier mois (€)", ca), _
        Array("Marge brute (€)", marge), _
        Array("EBITDA (€)", ebitda), _
        Array("Cash (€)", cash), _
        Array("DSO", dso))

    wsDash.Range("D1:F1").Value = Array("BU", "CA total (€)", "Marge brute totale (€)")
    lastRowBU = wsBU.Cells(wsBU.Rows.Count, "A").End(xlUp).Row
    GenererSyntheseBU wsBU, wsDash, lastRowBU

    wsDash.Columns.AutoFit
    MsgBox "Dashboard généré.", vbInformation
End Sub

Private Sub GenererSyntheseBU(wsBU As Worksheet, wsDash As Worksheet, lastRowBU As Long)
    Dim dCA As Object, dMarge As Object, i As Long, k As Variant, outRow As Long
    Set dCA = CreateObject("Scripting.Dictionary")
    Set dMarge = CreateObject("Scripting.Dictionary")

    For i = 2 To lastRowBU
        If dCA.Exists(CStr(wsBU.Cells(i, 1).Value)) Then
            dCA(CStr(wsBU.Cells(i, 1).Value)) = dCA(CStr(wsBU.Cells(i, 1).Value)) + wsBU.Cells(i, 3).Value
            dMarge(CStr(wsBU.Cells(i, 1).Value)) = dMarge(CStr(wsBU.Cells(i, 1).Value)) + wsBU.Cells(i, 4).Value
        Else
            dCA.Add CStr(wsBU.Cells(i, 1).Value), wsBU.Cells(i, 3).Value
            dMarge.Add CStr(wsBU.Cells(i, 1).Value), wsBU.Cells(i, 4).Value
        End If
    Next i

    outRow = 2
    For Each k In dCA.Keys
        wsDash.Cells(outRow, 4).Value = k
        wsDash.Cells(outRow, 5).Value = dCA(k)
        wsDash.Cells(outRow, 6).Value = dMarge(k)
        outRow = outRow + 1
    Next k
End Sub