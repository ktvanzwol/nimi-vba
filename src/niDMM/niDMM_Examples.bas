Attribute VB_Name = "niDMM_Examples"
Option Explicit

Sub MeasureDMM()
    Dim cDMM As niDMM_Session
    Dim sResourceName As String
    Dim lFunction As Long
    Dim dRange As Double
    Dim dResolutionDigits As Double
    Dim dResolutionAbsolute As Double
    Dim dReading As Double
    
    On Error GoTo Error
    
    sResourceName = "DMM_4081_C2_S02"
    lFunction = NIDMM_VAL_DC_VOLTS
    dRange = NIDMM_VAL_AUTO_RANGE_ON
    dResolutionDigits = 5.5
        
    Set cDMM = niDMM_CreateSession(sResourceName)
    With cDMM
        .ConfigureMeasurementDigits lFunction, dRange, dResolutionDigits
        .SetAttributeViReal64 "", NIDMM_ATTR_POWERLINE_FREQ, NIDMM_VAL_50_HERTZ
        .Read dReading
    End With
    
    cDMM.Read dReading
    cDMM.GetAttributeViReal64 "", NIDMM_ATTR_RESOLUTION_ABSOLUTE, dResolutionAbsolute
    
    Debug.Print "Reading    = "; dReading
    Debug.Print "Resolution = "; dResolutionAbsolute
    
Error:
    If Err Then niTools_ErrorMsgBox Err
    
End Sub

