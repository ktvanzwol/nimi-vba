Attribute VB_Name = "RFmx_Examples"
Option Explicit

Sub RFmxSpecAnTXP()
    Dim cRFmx As RFmx_Session
    Dim sResourceName As String
    Dim dCenterFrequency As Double
    Dim dReferenceLevel As Double
    Dim dExternalAttenuation As Double
    Dim dMeasInterval As Double
    Dim eRBWFilterType As RFmxSpecAn_TXPRBWFilterTypes
    Dim dRBW As Double
    Dim dRRCAlpha As Double
    Dim eVBWAuto As RFmxSpecAn_TXPVBWAutoEnabled
    Dim dVBW As Double
    Dim dVBWToRBWRatio As Double
    Dim bIQPowerEdgeEnabled As Boolean
    Dim dIQPowerEdgeLevel As Double
    Dim dTriggerDelay As Double
    Dim dMinQuietTime As Double
    Dim eAveragingEnabled As RFmxSpecAn_TXPAveragingEnabled
    Dim lAveragingCount As Long
    Dim eAveragingType As RFmxSpecAn_TXPAveragingTypes
    Dim eThresholdEnabled As RFmxSpecAn_TXPThresholdEnabled
    Dim eThresholdType As RFmxSpecAn_TXPThresholdType
    Dim dThresholdLevel As Double
    Dim dMinimumPower As Double
    Dim dAverageMeanPower As Double
    Dim dPeakToAveragePower As Double
    Dim dMaximumPower As Double
    Dim dX0 As Double, dDx As Double
    Dim dPowerTrace() As Single
    
    On Error GoTo Error
    
    sResourceName = "VST_5841_C1_S13"
    dCenterFrequency = 1000000000#
    dReferenceLevel = -10
    dExternalAttenuation = 3
    dMeasInterval = 0.001
    
    'RBW Filter
    eRBWFilterType = RFMXSPECAN_VAL_TXP_RBW_FILTER_TYPE_GAUSSIAN
    dRBW = 100000#
    dRRCAlpha = 0.01
    
    ' VBW
    eVBWAuto = RFMXSPECAN_VAL_TXP_VBW_FILTER_AUTO_BANDWIDTH_TRUE
    dVBW = 30000#
    dVBWToRBWRatio = 3
    
    ' Triggering
    bIQPowerEdgeEnabled = False
    dIQPowerEdgeLevel = -20#
    dTriggerDelay = 0#
    dMinQuietTime = 0#
    
    ' Averaging
    eAveragingEnabled = RFMXSPECAN_VAL_TXP_AVERAGING_ENABLED_FALSE
    lAveragingCount = 10
    eAveragingType = RFMXSPECAN_VAL_TXP_AVERAGING_TYPE_RMS
    
    ' Threshold
    eThresholdEnabled = RFMXSPECAN_VAL_TXP_THRESHOLD_ENABLED_FALSE
    eThresholdType = RFMXSPECAN_VAL_TXP_THRESHOLD_TYPE_RELATIVE
    dThresholdLevel = -20
    
    ' Setup
    Set cRFmx = RFmx_CreateSession(sResourceName, optionString:="Simulate=1, RFmxSetup=Model:5841")
    With cRFmx
        .CfgFrequencyReference "", "OnboardClock", 10000000#
        .SetAttributeString "", RFMXSPECAN_ATTR_SELECTED_PORTS, "" ' Only needed for Multi Port Devices
        .SpecAn.CfgFrequency "", dCenterFrequency
        .SpecAn.CfgReferenceLevel "", dReferenceLevel
        .SpecAn.CfgExternalAttenuation "", dExternalAttenuation
    End With
    
    'Trigger
    If bIQPowerEdgeEnabled Then
        cRFmx.SpecAn.CfgIQPowerEdgeTrigger "", "0", dIQPowerEdgeLevel, RFMXSPECAN_VAL_IQ_POWER_EDGE_RISING_SLOPE, _
                                           dTriggerDelay, RFMXSPECAN_VAL_TRIGGER_MINIMUM_QUIET_TIME_MODE_MANUAL, _
                                           dMinQuietTime, RFMX_VAL_TRUE
    Else
        cRFmx.SpecAn.DisableTrigger ""
    End If
    
    ' TXP Measurement
    With cRFmx.SpecAn
        .SelectMeasurements "", RFMXSPECAN_VAL_TXP, RFMX_VAL_TRUE
        .TXPCfgMeasurementInterval "", dMeasInterval
        .TXPCfgRBWFilter "", dRBW, eRBWFilterType, dRRCAlpha
        .TXPCfgAveraging "", eAveragingEnabled, lAveragingCount, eAveragingType
        .TXPCfgVBWFilter "", eVBWAuto, dVBW, dVBWToRBWRatio
        .TXPCfgThreshold "", eThresholdEnabled, dThresholdLevel, eThresholdType
    End With
    
    ' Initiate Measurement
    cRFmx.SpecAn.Initiate "", ""
    
    'Fetch Results
    cRFmx.SpecAn.TXPFetchMeasurement "", 10#, dAverageMeanPower, dPeakToAveragePower, dMaximumPower, dMinimumPower
    
    Debug.Print "TXP Measurement Results"
    Debug.Print "  Average Mean Power    :"; dAverageMeanPower; " dBm"
    Debug.Print "  Peak to Average Ratio :"; dPeakToAveragePower; " dB"
    Debug.Print "  Maximum Power         :"; dMaximumPower; " dBm"
    Debug.Print "  Minimum Power         :"; dMinimumPower; " dBm"
    
    cRFmx.SpecAn.TXPFetchPowerTrace "", 10#, dX0, dDx, dPowerTrace
    
    ' plot Power trace
    Dim ws As Worksheet
    Dim cs As Chart
    Dim index As Long
    Set ws = Example_GetNewOutputSheet("RFmaSpecAn TXP")
      
    ws.range("A1").Value2 = "Time"
    ws.range("B1").Value2 = "Power (dBM)"

    For index = 0 To UBound(dPowerTrace)
        ws.Cells(2 + index, 1).Value2 = dX0 + (index * dDx)
        ws.Cells(2 + index, 2).Value2 = dPowerTrace(index)
    Next
    
    ws.Shapes.AddChart2(240, xlXYScatterSmooth, 100, 10, 600, 400).Select
    ActiveChart.SetSourceData Source:=ws.UsedRange
    
Error:
    If Err Then niTools_ErrorMsgBox Err

End Sub

Sub GetDeviceTemperature()
    Dim cRFmx As RFmx_Session
    Dim sResourceName As String
    Dim dTemp As Double
    
    On Error GoTo Error
    
    sResourceName = "VST_5841_C1_S13"
    
    Set cRFmx = RFmx_CreateSession(sResourceName, optionString:="Simulate=1, RFmxSetup=Model:5841")
    cRFmx.GetAttributeF64 "", RFMXINSTR_ATTR_DEVICE_TEMPERATURE, dTemp

    Debug.Print "Device Temperature: "; dTemp
    
Error:
    If Err Then niTools_ErrorMsgBox Err

End Sub

Sub CreateRFmxFromRFSASession()
    Dim cRFmx As RFmx_Session
    Dim cRFSA As niRFSA_Session
    Dim cRFmxRFSA As niRFSA_Session
    Dim sResourceName As String
    Dim dTemp As Double
    Dim sMatch As String
    
    On Error GoTo Error
    
    sResourceName = "VST_5841_C1_S13"
    
    Set cRFSA = niRFSA_CreateSession(sResourceName, optionString:="Simulate=1,DriverSetup=Model:5841")
    Set cRFmx = RFmx_CreateSessionFromNIRFSASession(cRFSA)
     
    ' check internal sessions
    Set cRFmxRFSA = cRFmx.GetNIRFSASession
    sMatch = IIf(cRFmxRFSA.InternalSession = cRFSA.InternalSession, "Yes", "No")
    
    Debug.Print "RFSA Internal Session        :"; cRFSA.InternalSession
    Debug.Print "RFmx's RFSA Internal Session :"; cRFmxRFSA.InternalSession
    Debug.Print "Sessions Match?              : "; sMatch
    
    'Close GetNIRFSASession returned session, should not close the one used by RFmx
    Set cRFmxRFSA = Nothing
    
    cRFmx.GetAttributeF64 "", RFMXINSTR_ATTR_DEVICE_TEMPERATURE, dTemp
    Debug.Print "Device Temperature           :"; dTemp
    
    'Close the main RFSA session, this should invalidate the RFmx used session.
    Set cRFSA = Nothing
    
    ' this line should raise a session invalid error as the above line closed the RFSA session used by RFmx.
    cRFmx.GetAttributeF64 "", RFMXINSTR_ATTR_DEVICE_TEMPERATURE, dTemp
    Debug.Print "Device Temperature           :"; dTemp

Error:
    If Err Then niTools_ErrorMsgBox Err

End Sub

