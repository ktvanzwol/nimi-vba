Attribute VB_Name = "examples_ni568x"
Option Explicit

Sub MeasurePowerMeter()
    Dim cPM As ni568x_Session
    Dim sResourceName As String
    Dim lUnit As Long
    Dim dPower As Double
    Dim sUnit As String
    
    On Error GoTo Error
    
    sResourceName = "COM1"
    lUnit = NI568X_VAL_DBM
    
    Set cPM = ni568x_CreateSession(sResourceName)
    With cPM
        .ConfigureUnits lUnit
        .Read dPower
    End With
    
    Select Case lUnit
        Case NI568X_VAL_DBM
            sUnit = " dBm"
        Case NI568X_VAL_WATTS
            sUnit = " Watts"
        Case NI568X_VAL_MWATTS
            sUnit = " mWatts"
        Case NI568X_VAL_UWATTS
            sUnit = " uWatts"
        Case Else
            sUnit = ""
    End Select
    
    Debug.Print "Measured Power = "; dPower; sUnit
    
Error:
    If Err Then niTools_ErrorMsgBox Err
    
End Sub





