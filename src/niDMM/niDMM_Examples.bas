Attribute VB_Name = "niDMM_Examples"
Option Explicit

Sub MeasureDMM()
    Dim cDMM As niDMM_Session
    Dim sResourceName As String
    Dim lFunction As Long
    Dim dRange As Double
    Dim dResolutionDigits As Double
    Dim dReading As Double
    
    On Error GoTo Error
    
    sResourceName = "DMM_4081_C2_S02"
    lFunction = NIDMM_VAL_DC_VOLTS
    dRange = 5
    dResolutionDigits = 5.5
        
    Set cDMM = niDMM_CreateSession(sResourceName)
    With cDMM
        .ConfigureMeasurementDigits lFunction, dRange, dResolutionDigits
        .Powerline_Freq = NIDMM_VAL_50_HERTZ
        .Read dReading
    End With
    
    cDMM.Read dReading
    
    Debug.Print "Reading    = "; dReading
    Debug.Print "Resolution = "; cDMM.Resolution_Absolute
    
Error:
    If Err Then niTools_ErrorMsgBox Err
    
End Sub

