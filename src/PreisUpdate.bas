Attribute VB_Name = "PreisUpdate"
Option Explicit

Sub PreisUpdateAusAndererMappe()
    Const StartRow = 2
    Dim LastRow As Long
    Dim TargetSheet As Worksheet
    Dim SourceSheet As Worksheet
    
    Dim lngZeile As Long
    Dim rngTreffer As Range
    Dim lngZeileFrei As Long
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
 
    Set TargetSheet = tbl_Bestand
    Set SourceSheet = getSource
    If SourceSheet Is Nothing Then GoTo SubExit

    LastRow = SourceSheet.Range("A" & SourceSheet.Rows.Count).End(xlUp).Row

    For lngZeile = StartRow To LastRow

        Set rngTreffer = TargetSheet.Range("A:A").Find _
            (what:=SourceSheet.Range("A" & lngZeile).Value, lookat:=xlWhole)
        If rngTreffer Is Nothing Then
            lngZeileFrei = TargetSheet.Range("A" & _
                TargetSheet.Rows.Count).End(xlUp).Row + 1
            TargetSheet.Range("A" & lngZeileFrei).Value = SourceSheet.Range("A" & lngZeile).Value
            TargetSheet.Range("B" & lngZeileFrei).Value = SourceSheet.Range("B" & lngZeile).Value
            TargetSheet.Range("A" & lngZeileFrei).Interior.ColorIndex = 6
        Else
            rngTreffer.Offset(0, 1).Value = SourceSheet.Range("B" & lngZeile).Value
            rngTreffer.Offset(0, 1).BorderAround ColorIndex:=4
        End If

    Next lngZeile
    TargetSheet.Range("A:D").Sort Key1:=TargetSheet.Range("A1"), _
        order1:=xlAscending, Header:=xlYes

    SourceSheet.Parent.Close savechanges:=False
 
SubExit:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
End Sub

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





