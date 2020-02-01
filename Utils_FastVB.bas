Attribute VB_Name = "Utils_FastVB"
Option Explicit

Public Function FastVB(Optional TurnOn As Boolean = True, Optional access As Boolean = False) As Boolean
'* beginAlreadyOn, is temporary variable, to save the static AlreadyOn value for the debugmessage at the end
'* delete for release
'******************************************************
    Static AlreadyOn As Boolean
    
    Dim beginAlreadyOn As Boolean
    beginAlreadyOn = AlreadyOn
    
    If TurnOn Then
        If Not AlreadyOn Then
            FastVB = True
            AlreadyOn = True
            FastVB_On
            GoTo Exit_Sub
        End If
    Else
        If access Then
            AlreadyOn = False
            FastVB_Off
        End If
    End If
    FastVB = False
    
Exit_Sub:
    Debug.Print "FastVB: Turnon=" & TurnOn & " / Access=" & access & " / FastVB=" & FastVB & " / AlreadyOn=" & beginAlreadyOn & "-->" & AlreadyOn
    
End Function

Private Function FastVB_On() As Boolean
    
    With Application
        .EnableEvents = False
        .Calculation = xlCalculationManual
        .ScreenUpdating = False
        .DisplayAlerts = False
    End With
    
End Function

Private Function FastVB_Off() As Boolean
    
    With Application
        .EnableEvents = True
        .Calculation = xlCalculationAutomatic
        .ScreenUpdating = True
        .DisplayAlerts = True
        
        .Calculate
    End With
    
End Function

