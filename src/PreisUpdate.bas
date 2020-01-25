Attribute VB_Name = "PreisUpdate"
Option Explicit

Sub PreisUpdateAusAndererMappe()
    Dim TargetSheet As Worksheet
    Dim SourceSheet As Worksheet
    
    Set TargetSheet = getTarget
    If TargetSheet Is Nothing Then GoTo SubExit
    Set SourceSheet = getSource
    If SourceSheet Is Nothing Then GoTo SubExit
    
    UpdatePrice getTarget, getSource
    SortSheet TargetSheet
    CloseSourceBySheet SourceSheet
    

SubExit:

End Sub

Private Sub CloseSourceBySheet(SourceSheet As Worksheet)
    SourceSheet.Parent.Close savechanges:=False
End Sub

Private Sub UpdatePrice(TargetSheet As Worksheet, SourceSheet As Worksheet)
    Const StartRow = 2
    Dim LastRow As Long
    Dim actualRow As Long
    
    Dim rngTreffer As Range
    Dim lngZeileFrei As Long
    
    AppFunktions False
    
    LastRow = getLastRow(SourceSheet)
    For actualRow = StartRow To LastRow

        Set rngTreffer = TargetSheet.Range("A:A").Find _
            (what:=SourceSheet.Range("A" & actualRow).Value, lookat:=xlWhole)
        If rngTreffer Is Nothing Then
            lngZeileFrei = TargetSheet.Range("A" & _
                TargetSheet.Rows.Count).End(xlUp).Row + 1
            TargetSheet.Range("A" & lngZeileFrei).Value = SourceSheet.Range("A" & actualRow).Value
            TargetSheet.Range("B" & lngZeileFrei).Value = SourceSheet.Range("B" & actualRow).Value
            TargetSheet.Range("A" & lngZeileFrei).Interior.ColorIndex = 6
        Else
            rngTreffer.Offset(0, 1).Value = SourceSheet.Range("B" & actualRow).Value
            rngTreffer.Offset(0, 1).BorderAround ColorIndex:=4
        End If

    Next actualRow
    
 
SubExit:
    AppFunktions True
    
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




