VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RFmx_SpecAn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Enum RFmxSpecAn_IQPowerEdgeTriggerSlope
    RFMXSPECAN_VAL_IQ_POWER_EDGE_RISING_SLOPE = 0
    RFMXSPECAN_VAL_IQ_POWER_EDGE_FALLING_SLOPE = 1
End Enum

Enum RFmxSpecAn_TriggerMinimumQuietTimeMode
    RFMXSPECAN_VAL_TRIGGER_MINIMUM_QUIET_TIME_MODE_MANUAL = 0
    RFMXSPECAN_VAL_TRIGGER_MINIMUM_QUIET_TIME_MODE_AUTO = 1
End Enum

' Values for MeasurementTypes
Enum RFmxSpecAn_MeasurementTypes
    RFMXSPECAN_VAL_ACP = 1            ' 1 << 0
    RFMXSPECAN_VAL_CCDF = 2           ' 1 << 1
    RFMXSPECAN_VAL_CHP = 4            ' 1 << 2
    RFMXSPECAN_VAL_FCNT = 8           ' 1 << 3
    RFMXSPECAN_VAL_HARMONICS = 16     ' 1 << 4
    RFMXSPECAN_VAL_OBW = 32           ' 1 << 5
    RFMXSPECAN_VAL_SEM = 64           ' 1 << 6
    RFMXSPECAN_VAL_SPECTRUM = 128     ' 1 << 7
    RFMXSPECAN_VAL_SPUR = 256         ' 1 << 8
    RFMXSPECAN_VAL_TXP = 512          ' 1 << 9
    RFMXSPECAN_VAL_AMPM = 1024        ' 1 << 10
    RFMXSPECAN_VAL_DPD = 2048         ' 1 << 11
    RFMXSPECAN_VAL_IQ = 4096          ' 1 << 12
    RFMXSPECAN_VAL_IM = 8192          ' 1 << 13
    RFMXSPECAN_VAL_NF = 16384         ' 1 << 14
    RFMXSPECAN_VAL_PHASENOISE = 32768 ' 1 << 15
    RFMXSPECAN_VAL_PAVT = 65536       ' 1 << 16
End Enum

Enum RFmxSpecAn_TXPRBWFilterTypes
    RFMXSPECAN_VAL_TXP_RBW_FILTER_TYPE_NONE = 5             ' The measurement does not use any RBW filtering.
    RFMXSPECAN_VAL_TXP_RBW_FILTER_TYPE_GAUSSIAN = 1         ' The RBW filter has a Gaussian response.
    RFMXSPECAN_VAL_TXP_RBW_FILTER_TYPE_FLAT = 2             ' The RBW filter has a flat response.
    RFMXSPECAN_VAL_TXP_RBW_FILTER_TYPE_SYNCH_TUNED_4 = 3    ' The RBW filter has a response of a 4-pole synchronously-tuned filter.
    RFMXSPECAN_VAL_TXP_RBW_FILTER_TYPE_SYNCH_TUNED_5 = 4    ' The RBW filter has a response of a 5-pole synchronously-tuned filter.
    RFMXSPECAN_VAL_TXP_RBW_FILTER_TYPE_RRC = 6              ' The RRC filter with the roll-off specified by RRCAlpha parameter is used as the RBW filter.
End Enum

' Values for RFMXSPECAN_ATTR_TXP_THRESHOLD_ENABLED
Enum RFmxSpecAn_TXPThresholdEnabled
    RFMXSPECAN_VAL_TXP_THRESHOLD_ENABLED_FALSE = RFMX_VAL_FALSE
    RFMXSPECAN_VAL_TXP_THRESHOLD_ENABLED_TRUE = RFMX_VAL_TRUE
End Enum

' Values for RFMXSPECAN_ATTR_TXP_THRESHOLD_TYPE
Enum RFmxSpecAn_TXPThresholdType
    RFMXSPECAN_VAL_TXP_THRESHOLD_TYPE_RELATIVE = 0
    RFMXSPECAN_VAL_TXP_THRESHOLD_TYPE_ABSOLUTE = 1
End Enum

Enum RFmxSpecAn_TXPAveragingEnabled
    RFMXSPECAN_VAL_TXP_AVERAGING_ENABLED_FALSE = RFMX_VAL_FALSE
    RFMXSPECAN_VAL_TXP_AVERAGING_ENABLED_TRUE = RFMX_VAL_TRUE
End Enum
 
' Values for RFMXSPECAN_ATTR_TXP_AVERAGING_TYPE
Enum RFmxSpecAn_TXPAveragingTypes
    RFMXSPECAN_VAL_TXP_AVERAGING_TYPE_RMS = 0
    RFMXSPECAN_VAL_TXP_AVERAGING_TYPE_LOG = 1
    RFMXSPECAN_VAL_TXP_AVERAGING_TYPE_SCALAR = 2
    RFMXSPECAN_VAL_TXP_AVERAGING_TYPE_MAXIMUM = 3
    RFMXSPECAN_VAL_TXP_AVERAGING_TYPE_MINIMUM = 4
End Enum

Enum RFmxSpecAn_TXPVBWAutoEnabled
    RFMXSPECAN_VAL_TXP_VBW_FILTER_AUTO_BANDWIDTH_FALSE = RFMX_VAL_FALSE
    RFMXSPECAN_VAL_TXP_VBW_FILTER_AUTO_BANDWIDTH_TRUE = RFMX_VAL_TRUE
End Enum

'int32 __stdcall RFmxSpecAn_GetError (niRFmxInstrHandle instrumentHandle, int32* errorCode, int32 errorDescriptionBufferSize, char errorDescription[]);
Private Declare PtrSafe Function RFmxSpecAn_GetError Lib "niRFmxSpecAn" ( _
    ByVal instrumentHandle As LongPtr, ByRef errorCode As Long, ByVal errorDescriptionBufferSize As Long, ByVal errorDescription As LongPtr) As Long
    
'int32 __stdcall RFmxSpecAn_CfgFrequency (niRFmxInstrHandle instrumentHandle, char selectorString[], float64 centerFrequency);
Private Declare PtrSafe Function RFmxSpecAn_CfgFrequency Lib "niRFmxSpecAn" ( _
    ByVal instrumentHandle As LongPtr, ByVal selectorString As String, ByVal centerFrequency As Double) As Long

'int32 __stdcall RFmxSpecAn_CfgReferenceLevel (niRFmxInstrHandle instrumentHandle, char selectorString[], float64 referenceLevel);
Private Declare PtrSafe Function RFmxSpecAn_CfgReferenceLevel Lib "niRFmxSpecAn" ( _
    ByVal instrumentHandle As LongPtr, ByVal selectorString As String, ByVal referenceLevel As Double) As Long

'int32 __stdcall RFmxSpecAn_CfgExternalAttenuation (niRFmxInstrHandle instrumentHandle, char selectorString[], float64 externalAttenuation);
Private Declare PtrSafe Function RFmxSpecAn_CfgExternalAttenuation Lib "niRFmxSpecAn" ( _
    ByVal instrumentHandle As LongPtr, ByVal selectorString As String, ByVal externalAttenuation As Double) As Long

'int32 __stdcall RFmxSpecAn_CfgIQPowerEdgeTrigger (niRFmxInstrHandle instrumentHandle, char selectorString[], char IQPowerEdgeSource[], float64 IQPowerEdgeLevel,
'                                                  int32 IQPowerEdgeSlope, float64 triggerDelay, int32 triggerMinQuietTimeMode, float64 triggerMinQuietTimeDuration,
'                                                  int32 enableTrigger);
Private Declare PtrSafe Function RFmxSpecAn_CfgIQPowerEdgeTrigger Lib "niRFmxSpecAn" ( _
    ByVal instrumentHandle As LongPtr, ByVal selectorString As String, ByVal IQPowerEdgeSource As String, _
    ByVal IQPowerEdgeLevel As Double, ByVal IQPowerEdgeSlope As Long, ByVal triggerDelay As Double, ByVal triggerMinQuietTimeMode As Long, _
    ByVal triggerMinQuietTimeDuration As Double, ByVal enableTrigger As Long) As Long

'int32 __stdcall RFmxSpecAn_DisableTrigger (niRFmxInstrHandle instrumentHandle, char selectorString[]);
Private Declare PtrSafe Function RFmxSpecAn_DisableTrigger Lib "niRFmxSpecAn" ( _
    ByVal instrumentHandle As LongPtr, ByVal selectorString As String) As Long

'int32 __stdcall RFmxSpecAn_SelectMeasurements (niRFmxInstrHandle instrumentHandle, char selectorString[], uInt32 measurements, int32 enableAllTraces);
Private Declare PtrSafe Function RFmxSpecAn_SelectMeasurements Lib "niRFmxSpecAn" ( _
    ByVal instrumentHandle As LongPtr, ByVal selectorString As String, ByVal measurements As Long, ByVal enableAllTraces As Long) As Long

'int32 __stdcall RFmxSpecAn_TXPCfgMeasurementInterval (niRFmxInstrHandle instrumentHandle, char selectorString[], float64 measurementInterval);
Private Declare PtrSafe Function RFmxSpecAn_TXPCfgMeasurementInterval Lib "niRFmxSpecAn" ( _
    ByVal instrumentHandle As LongPtr, ByVal selectorString As String, ByVal measurementInterval As Double) As Long

'int32 __stdcall RFmxSpecAn_TXPCfgRBWFilter (niRFmxInstrHandle instrumentHandle, char selectorString[], float64 RBW, int32 RBWFilterType, float64 RRCAlpha);
Private Declare PtrSafe Function RFmxSpecAn_TXPCfgRBWFilter Lib "niRFmxSpecAn" ( _
    ByVal instrumentHandle As LongPtr, ByVal selectorString As String, ByVal RBW As Double, ByVal RBWFilterType As Long, ByVal RRCAlpha As Double) As Long

'int32 __stdcall RFmxSpecAn_TXPCfgThreshold (niRFmxInstrHandle instrumentHandle, char selectorString[], int32 thresholdEnabled, float64 thresholdLevel, int32 thresholdType);
Private Declare PtrSafe Function RFmxSpecAn_TXPCfgThreshold Lib "niRFmxSpecAn" ( _
    ByVal instrumentHandle As LongPtr, ByVal selectorString As String, ByVal thresholdEnabled As Long, ByVal thresholdLevel As Double, ByVal thresholdType As Long) As Long

'int32 __stdcall RFmxSpecAn_TXPCfgAveraging (niRFmxInstrHandle instrumentHandle, char selectorString[], int32 averagingEnabled, int32 averagingCount, int32 averagingType);
Private Declare PtrSafe Function RFmxSpecAn_TXPCfgAveraging Lib "niRFmxSpecAn" ( _
    ByVal instrumentHandle As LongPtr, ByVal selectorString As String, ByVal averagingEnabled As Long, ByVal averagingCount As Long, ByVal averagingType As Long) As Long

'int32 __stdcall RFmxSpecAn_TXPCfgVBWFilter (niRFmxInstrHandle instrumentHandle, char selectorString[], int32 VBWAuto, float64 VBW, float64 VBWToRBWRatio);
Private Declare PtrSafe Function RFmxSpecAn_TXPCfgVBWFilter Lib "niRFmxSpecAn" ( _
    ByVal instrumentHandle As LongPtr, ByVal selectorString As String, ByVal VBWAuto As Long, ByVal VBW As Double, ByVal VBWToRBWRatio As Double) As Long

'int32 __stdcall RFmxSpecAn_Initiate (niRFmxInstrHandle instrumentHandle, char selectorString[], char resultName[]);
Private Declare PtrSafe Function RFmxSpecAn_Initiate Lib "niRFmxSpecAn" ( _
    ByVal instrumentHandle As LongPtr, ByVal selectorString As String, ByVal resultName As String) As Long

'int32 __stdcall RFmxSpecAn_TXPFetchMeasurement (niRFmxInstrHandle instrumentHandle, char selectorString[], float64 timeout, float64* averageMeanPower,
'                                                float64* peakToAverageRatio, float64* maximumPower, float64* minimumPower);
Private Declare PtrSafe Function RFmxSpecAn_TXPFetchMeasurement Lib "niRFmxSpecAn" ( _
    ByVal instrumentHandle As LongPtr, ByVal selectorString As String, ByVal timeout As Double, ByRef averageMeanPower As Double, _
    ByRef peakToAverageRatio As Double, ByRef maximumPower As Double, ByRef minimumPower As Double) As Long

'int32 __stdcall RFmxSpecAn_TXPFetchPowerTrace (niRFmxInstrHandle instrumentHandle, char selectorString[], float64 timeout, float64* x0, float64* dx, float32 power[],
'                                               int32 arraySize, int32* actualArraySize);
Private Declare PtrSafe Function RFmxSpecAn_TXPFetchPowerTrace Lib "niRFmxSpecAn" ( _
    ByVal instrumentHandle As LongPtr, ByVal selectorString As String, ByVal timeout As Double, ByRef x0 As Double, _
    ByRef dx As Double, ByVal power As LongPtr, ByVal arraySize As Long, ByRef actualArraySize As Long) As Long

Private m_Handle As LongPtr

' initialize internal variables, call Init first to create a valid session
Private Sub Class_Initialize()
    m_Handle = 0
End Sub

' Automatically clear session when object gets destroyed
Private Sub Class_Terminate()
    m_Handle = 0
End Sub

' Error Checker
Private Sub CheckError(status As Long)
    If status < 0 Then
        ErrorHandler status
    End If
End Sub

' Error Handler
Private Sub ErrorHandler(errorCode As Long)
    Dim status As Long
    Dim size As Long
    Dim buffer() As Byte
    Dim errorMsg As String
    
    size = RFmxSpecAn_GetError(m_Handle, errorCode, 0, 0)
    ReDim buffer(size - 1) As Byte
 
    status = RFmxSpecAn_GetError(m_Handle, errorCode, size, VarPtr(buffer(0)))
    errorMsg = StrConv(LeftB(buffer(), size - 1), vbUnicode) 'Remove \0 character and convert to Unicode
    
    niTools_RaiseError errorCode, errorMsg, "NI-RFmxSpecAn"
End Sub

Public Sub InitSpecAn(handle As LongPtr)
    m_Handle = handle
End Sub

Public Sub CfgFrequency(selectorString As String, centerFrequency As Double)
    CheckError RFmxSpecAn_CfgFrequency(m_Handle, selectorString, centerFrequency)
End Sub

Public Sub CfgReferenceLevel(selectorString As String, referenceLevel As Double)
    CheckError RFmxSpecAn_CfgReferenceLevel(m_Handle, selectorString, referenceLevel)
End Sub

Public Sub CfgExternalAttenuation(selectorString As String, ByVal externalAttenuation As Double)
    CheckError RFmxSpecAn_CfgExternalAttenuation(m_Handle, selectorString, externalAttenuation)
End Sub

Public Sub CfgIQPowerEdgeTrigger(selectorString As String, IQPowerEdgeSource As String, _
    IQPowerEdgeLevel As Double, IQPowerEdgeSlope As RFmxSpecAn_IQPowerEdgeTriggerSlope, triggerDelay As Double, _
    triggerMinQuietTimeMode As RFmxSpecAn_TriggerMinimumQuietTimeMode, triggerMinQuietTimeDuration As Double, enableTrigger As RFmx_Binary)
    
    CheckError RFmxSpecAn_CfgIQPowerEdgeTrigger(m_Handle, selectorString, IQPowerEdgeSource, IQPowerEdgeLevel, _
                                                IQPowerEdgeSlope, triggerDelay, triggerMinQuietTimeMode, triggerMinQuietTimeDuration, _
                                                enableTrigger)
End Sub

Public Sub DisableTrigger(selectorString As String)
    CheckError RFmxSpecAn_DisableTrigger(m_Handle, selectorString)
End Sub

Public Sub SelectMeasurements(selectorString As String, measurements As RFmxSpecAn_MeasurementTypes, enableAllTraces As RFmx_Binary)
    CheckError RFmxSpecAn_SelectMeasurements(m_Handle, selectorString, measurements, enableAllTraces)
End Sub

Public Sub TXPCfgMeasurementInterval(selectorString As String, measurementInterval As Double)
    CheckError RFmxSpecAn_TXPCfgMeasurementInterval(m_Handle, selectorString, measurementInterval)
End Sub

Public Sub TXPCfgRBWFilter(selectorString As String, RBW As Double, RBWFilterType As RFmxSpecAn_TXPRBWFilterTypes, RRCAlpha As Double)
    CheckError RFmxSpecAn_TXPCfgRBWFilter(m_Handle, selectorString, RBW, RBWFilterType, RRCAlpha)
End Sub

Public Sub TXPCfgThreshold(selectorString As String, thresholdEnabled As RFmxSpecAn_TXPThresholdEnabled, thresholdLevel As Double, thresholdType As RFmxSpecAn_TXPThresholdType)
    CheckError RFmxSpecAn_TXPCfgThreshold(m_Handle, selectorString, thresholdEnabled, thresholdLevel, thresholdType)
End Sub

Public Sub TXPCfgAveraging(selectorString As String, averagingEnabled As RFmxSpecAn_TXPAveragingEnabled, averagingCount As Long, averagingType As RFmxSpecAn_TXPAveragingTypes)
    CheckError RFmxSpecAn_TXPCfgAveraging(m_Handle, selectorString, averagingEnabled, averagingCount, averagingType)
End Sub

Public Sub TXPCfgVBWFilter(selectorString As String, VBWAuto As RFmxSpecAn_TXPVBWAutoEnabled, VBW As Double, VBWToRBWRatio As Double)
    CheckError RFmxSpecAn_TXPCfgVBWFilter(m_Handle, selectorString, VBWAuto, VBW, VBWToRBWRatio)
End Sub

Public Sub Initiate(selectorString As String, resultName As String)
    CheckError RFmxSpecAn_Initiate(m_Handle, selectorString, resultName)
End Sub

Public Sub TXPFetchMeasurement(selectorString As String, timeout As Double, ByRef averageMeanPower As Double, _
    ByRef peakToAverageRatio As Double, ByRef maximumPower As Double, ByRef minimumPower As Double)
    
    CheckError RFmxSpecAn_TXPFetchMeasurement(m_Handle, selectorString, timeout, averageMeanPower, peakToAverageRatio, maximumPower, minimumPower)
End Sub

Public Sub TXPFetchPowerTrace(selectorString As String, timeout As Double, ByRef x0 As Double, ByRef dx As Double, ByRef power() As Single)
    Dim size As Long
    
    CheckError RFmxSpecAn_TXPFetchPowerTrace(m_Handle, selectorString, timeout, x0, dx, 0, 0, size)
    ReDim power(size - 1) As Single

    CheckError RFmxSpecAn_TXPFetchPowerTrace(m_Handle, selectorString, timeout, x0, dx, VarPtr(power(0)), size, size)
End Sub


