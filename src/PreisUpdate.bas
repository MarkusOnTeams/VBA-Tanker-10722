Attribute VB_Name = "PreisUpdate"
Option Explicit
    
Private SourceSheet As Worksheet
'

Sub PreisUpdateAusAndererMappe()
    Dim TargetSheet As Worksheet
    
    Set TargetSheet = getTarget
    If TargetSheet Is Nothing Then GoTo SubExit
    Set SourceSheet = getSource
    If SourceSheet Is Nothing Then GoTo SubExit
    
    UpdatePriceTable TargetSheet
    SortSheet TargetSheet
    CloseSourceBySheet SourceSheet
    

SubExit:

End Sub

Private Sub UpdatePriceTable(TargetSheet As Worksheet)
    Const StartRow = 2
    Dim LastRow As Long
    Dim ActualRow As Long
    
    Dim rngTreffer As Range
    
    AppFunktions False
    
    LastRow = getLastRow(SourceSheet)
    For ActualRow = StartRow To LastRow

        Set rngTreffer = FindNextValue(TargetSheet, ActualRow)
        If rngTreffer Is Nothing Then
            AddValue TargetSheet, ActualRow
        Else
            UpdateValue rngTreffer, ActualRow
        End If

    Next ActualRow
    
 
SubExit:
    AppFunktions True
    
End Sub

Private Function FindNextValue(TargetSheet As Worksheet, ActualRow As Long) As Range
    Set FindNextValue = TargetSheet.Range("A:A").Find _
            (what:=SourceSheet.Range("A" & ActualRow).Value, _
            lookat:=xlWhole)
End Function

Private Sub AddValue(TargetSheet As Worksheet, ActualRow As Long)
    Dim NextRow As Long
    
    NextRow = getLastRow(TargetSheet) + 1
    TargetSheet.Range("A" & NextRow).Value = SourceSheet.Range("A" & ActualRow).Value
    TargetSheet.Range("B" & NextRow).Value = SourceSheet.Range("B" & ActualRow).Value
    
    TargetSheet.Range("A" & NextRow).Interior.ColorIndex = 6
End Sub

Private Sub UpdateValue(TargetRange As Range, UpdateRow As Long)
    TargetRange.Offset(0, 1).Value = SourceSheet.Range("B" & UpdateRow).Value
    TargetRange.Offset(0, 1).Interior.ColorIndex = 4
End Sub

Private Sub CloseSourceBySheet(SourceSheet As Worksheet)
    SourceSheet.Parent.Close savechanges:=False
End Sub

Private Sub SortSheet(TargetSheet As Worksheet)
        
        TargetSheet.Range("A:D").Sort _
                Key1:=TargetSheet.Range("A1"), _
                order1:=xlAscending, _
                Header:=xlYes

End Sub

Private Function getTarget() As Worksheet
    Set getTarget = tbl_Bestand
End Function

Private Function getSource() As Worksheet
    Dim result As Worksheet
    
    On Error GoTo ErrorExit
    With Workbooks.Open(ThisWorkbook.Path & "\Preise.xlsx")
        Set result = .Worksheets(1)
    End With
    
FuncExit:
    Set getSource = result
    Exit Function
    
ErrorExit:
    Set result = Nothing
    Resume FuncExit
End Function

Private Function getLastRow(SourceSheet As Worksheet) As Long
    getLastRow = SourceSheet.Range("A" & SourceSheet.Rows.Count).End(xlUp).Row
End Function

Private Sub AppFunktions(TurnOn As Boolean)

    Application.ScreenUpdating = TurnOn
    Application.DisplayAlerts = TurnOn
    If TurnOn Then
        Application.Calculation = xlCalculationAutomatic
    Else
        Application.Calculation = xlCalculationManual
    End If
    
End Sub





