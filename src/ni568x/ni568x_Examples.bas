Attribute VB_Name = "ni568x_Examples"
Option Explicit

Sub MeasurePowerMeter()
    Dim cPM As ni568x_Session
    Dim sResourceName As String
    Dim eUnit As ni568x_Units
    Dim dFrequency As Double
    Dim dOffset As Double
    Dim dPower As Double
    Dim sUnit As String
    Dim eZeroStatus As ni568x_ZeroStatus
    
    On Error GoTo Error
    
    sResourceName = "COM1"
    eUnit = NI568X_VAL_DBM
    dFrequency = 1000000#
    dOffset = -3#
    
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
    
    Set cPM = ni568x_CreateSession(sResourceName)
    With cPM
        .SetAttributeViInt32 "", NI568X_ATTR_UNITS, lUnit
        .SetAttributeViReal64 "", NI568X_ATTR_CORRECTION_FREQUENCY, dFrequency
        .SetAttributeViReal64 "", NI568X_ATTR_OFFSET, dOffset
    End With
    
    cPM.DisableOffset
    cPM.Read dPower
    Debug.Print "Measured Power = "; dPower; sUnit
    
    cPM.EnableOffset
    cPM.Read dPower
    Debug.Print "Measured Power = "; dPower; sUnit; " +Offset"
    
    cPM.Zero
    Do
        DoEvents
        cPM.IsZeroCompleted eZeroStatus
        
    Loop While eZeroStatus = NI568X_VAL_ZERO_IN_PROGRESS

    cPM.Read dPower
    Debug.Print "Measured Power = "; dPower; sUnit; " +Zero +Offset"
    
Error:
    If Err Then niTools_ErrorMsgBox Err
    
End Sub


