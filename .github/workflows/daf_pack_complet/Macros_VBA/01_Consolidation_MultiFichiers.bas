Attribute VB_Name = "modConsolidation"
Option Explicit

Sub ConsoliderFiliales()
    Dim wsDest As Worksheet
    Dim wbSource As Workbook
    Dim wsSource As Worksheet
    Dim dossier As String
    Dim fichier As String
    Dim lastRowDest As Long
    Dim lastRowSrc As Long
    Dim i As Long

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    Set wsDest = ThisWorkbook.Worksheets("Consolidation")
    wsDest.Rows("2:" & wsDest.Rows.Count).ClearContents

    dossier = ThisWorkbook.Path & Application.PathSeparator
    fichier = Dir(dossier & "*.xlsx")

    Do While fichier <> ""
        If fichier <> ThisWorkbook.Name And InStr(1, fichier, "Filiale_", vbTextCompare) > 0 Then
            Set wbSource = Workbooks.Open(dossier & fichier, ReadOnly:=True)
            Set wsSource = wbSource.Worksheets(1)

            lastRowSrc = wsSource.Cells(wsSource.Rows.Count, "A").End(xlUp).Row
            If lastRowSrc >= 2 Then
                For i = 2 To lastRowSrc
                    If Application.WorksheetFunction.CountA(wsSource.Rows(i)) > 0 Then
                        lastRowDest = wsDest.Cells(wsDest.Rows.Count, "A").End(xlUp).Row + 1
                        wsDest.Cells(lastRowDest, 1).Value = wsSource.Cells(i, 1).Value
                        wsDest.Cells(lastRowDest, 2).Value = wsSource.Cells(i, 2).Value
                        wsDest.Cells(lastRowDest, 3).Value = wsSource.Cells(i, 3).Value
                        wsDest.Cells(lastRowDest, 4).Value = wsSource.Cells(i, 4).Value
                        wsDest.Cells(lastRowDest, 5).Value = wsSource.Cells(i, 5).Value
                        wsDest.Cells(lastRowDest, 6).Value = Replace(Replace(fichier, ".xlsx", ""), ".xlsm", "")
                    End If
                Next i
            End If

            wbSource.Close SaveChanges:=False
        End If
        fichier = Dir
    Loop

    lastRowDest = wsDest.Cells(wsDest.Rows.Count, "A").End(xlUp).Row
    If lastRowDest >= 2 Then
        wsDest.Range("A1:F" & lastRowDest).RemoveDuplicates Columns:=Array(1, 2, 3, 4, 5, 6), Header:=xlYes
        wsDest.Range("A1:F" & lastRowDest).Sort Key1:=wsDest.Range("A2"), Order1:=xlAscending, Header:=xlYes
        wsDest.Columns("A:F").AutoFit
    End If

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

    MsgBox "Consolidation terminée.", vbInformation
End Sub