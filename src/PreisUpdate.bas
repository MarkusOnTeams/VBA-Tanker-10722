Attribute VB_Name = "PreisUpdate"
Option Explicit

Sub PreisUpdateAusAndererMappe()
    Dim lngZeile As Long
    Dim lngZeileMax As Long
    Dim rngTreffer As Range
    Dim lngZeileFrei As Long
    
    Dim TargetSheet As Worksheet
    Dim SourceSheet As Worksheet
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
 
    Set SourceSheet = getSource
    If SourceSheet Is Nothing Then GoTo SubExit
    
    
'    With SourceSheet
'
'        lngZeileMax = .Range("A" & .Rows.Count).End(xlUp).Row
'
'        For lngZeile = 2 To lngZeileMax
'
'            Set rngTreffer = tbl_Bestand.Range("A:A").Find _
'                (what:=.Range("A" & lngZeile).Value, lookat:=xlWhole)
'            If rngTreffer Is Nothing Then
'                lngZeileFrei = tbl_Bestand.Range("A" & _
'                    tbl_Bestand.Rows.Count).End(xlUp).Row + 1
'                tbl_Bestand.Range("A" & lngZeileFrei).Value = .Range("A" & lngZeile).Value
'                tbl_Bestand.Range("B" & lngZeileFrei).Value = .Range("B" & lngZeile).Value
'                tbl_Bestand.Range("A" & lngZeileFrei).Interior.ColorIndex = 6
'            Else
'                rngTreffer.Offset(0, 1).Value = .Range("B" & lngZeile).Value
'                rngTreffer.Offset(0, 1).BorderAround ColorIndex:=4
'            End If
'
'        Next lngZeile
'        tbl_Bestand.Range("A:D").Sort Key1:=tbl_Bestand.Range("A1"), _
'            order1:=xlAscending, Header:=xlYes
'
'    End With
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

