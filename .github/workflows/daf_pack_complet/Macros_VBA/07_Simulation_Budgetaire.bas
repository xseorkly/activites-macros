Attribute VB_Name = "modSimulationBudget"
Option Explicit

Sub SimulerScenarios()
    Dim ws As Worksheet, wsOut As Worksheet
    Dim volume As Double, prix As Double, cvu As Double
    Dim chargesFixes As Double
    Dim scenarios, i As Long
    Dim facteurVol As Double, facteurPrix As Double

    Set ws = ThisWorkbook.Worksheets("Hypotheses")
    On Error Resume Next
    Set wsOut = ThisWorkbook.Worksheets("Scenarios")
    On Error GoTo 0

    If wsOut Is Nothing Then
        Set wsOut = ThisWorkbook.Worksheets.Add(After:=ws)
        wsOut.Name = "Scenarios"
    Else
        wsOut.Cells.Clear
    End If

    volume = ws.Range("B2").Value
    prix = ws.Range("B3").Value
    cvu = ws.Range("B4").Value
    chargesFixes = Application.WorksheetFunction.Sum(ws.Range("B5:B8"))

    scenarios = Array("Prudent", "Central", "Optimiste")

    wsOut.Range("A1:E1").Value = Array("Scénario", "CA (€)", "Marge brute (€)", "Charges fixes (€)", "Résultat (€)")

    For i = 0 To 2
        Select Case scenarios(i)
            Case "Prudent": facteurVol = 0.92: facteurPrix = 0.98
            Case "Central": facteurVol = 1: facteurPrix = 1
            Case "Optimiste": facteurVol = 1.08: facteurPrix = 1.03
        End Select

        wsOut.Cells(i + 2, 1).Value = scenarios(i)
        wsOut.Cells(i + 2, 2).Value = volume * facteurVol * prix * facteurPrix
        wsOut.Cells(i + 2, 3).Value = volume * facteurVol * (prix * facteurPrix - cvu)
        wsOut.Cells(i + 2, 4).Value = chargesFixes
        wsOut.Cells(i + 2, 5).Value = wsOut.Cells(i + 2, 3).Value - chargesFixes
    Next i

    wsOut.Columns.AutoFit
    MsgBox "Scénarios budgétaires calculés.", vbInformation
End Sub