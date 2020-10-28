Attribute VB_Name = "niDCPower_Examples"
Option Explicit

Sub SinglePointForceMeasureDCVoltage()
    Dim cDCP As niDCPower_Session
    Dim sResourceName As String
    Dim sChannel As String
    Dim lNumChannels As Long
    Dim dLevel As Double
    Dim dLevelRange As Double
    Dim dLimit As Double
    Dim dLimitRange As Double
    Dim dSourceDelay As Double
    Dim dVoltMeasurements() As Double
    Dim dCurrentMeasurements() As Double
    Dim bInCompliance As Boolean
            
    On Error GoTo Error
    
    sResourceName = "SMU_4143_C1_S06"
    sChannel = "0"
    lNumChannels = 1 'Needs to match the number of channels specified in sChannel
    dLevel = 3.3 'V
    dLevelRange = dLevel 'Select closest range based on source level
    dLimit = 0.1 'A
    dLimitRange = dLimit 'Select closest range based on source level
    dSourceDelay = 0.01 'sec
    
    ' Measurement Array sizes needs to match the number of channels
    ReDim dVoltMeasurements(lNumChannels - 1) As Double
    ReDim dCurrentMeasurements(lNumChannels - 1) As Double
    
    Set cDCP = niDCPower_CreateSession(sResourceName, sChannel)
    With cDCP
        .ConfigureSourceMode NIDCPOWER_VAL_SINGLE_POINT
        .ConfigureOutputFunction sChannel, NIDCPOWER_VAL_DC_VOLTAGE
        .ConfigureVoltageLevel sChannel, dLevel
        .ConfigureVoltageLevelRange sChannel, dLevelRange
        .ConfigureCurrentLimit sChannel, dLimit
        .ConfigureCurrentLimitRange sChannel, dLimitRange
        .SetAttributeDouble sChannel, NIDCPOWER_ATTR_SOURCE_DELAY, dSourceDelay
    End With
    
    ' Initiate Sourcing
    cDCP.Initiate
    
    ' Wait for Source complete (settling time defined by Source Delay)
    cDCP.WaitForEvent NIDCPOWER_VAL_SOURCE_COMPLETE_EVENT
    
    ' Measure Channel Voltage and Currents
    cDCP.MeasureMultiple sChannel, dVoltMeasurements, dCurrentMeasurements
    
    ' Check if channel went into compliance
    cDCP.QueryInCompliance sChannel, bInCompliance
    
    Debug.Print "Measurement Results:"
    Debug.Print "    Voltage       : "; dVoltMeasurements(0)
    Debug.Print "    Current       : "; dCurrentMeasurements(0)
    Debug.Print "    In Compliance : "; bInCompliance
    
Error:
    If Err Then niTools_ErrorMsgBox Err
    cDCP.Reset
    
End Sub

