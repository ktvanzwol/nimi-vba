VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "niDCPower_Session"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Attribute IDs
' Note: in header files
' IVI_SPECIFIC_PUBLIC_ATTR_BASE = 1150000
' IVI_CLASS_PUBLIC_ATTR_BASE = 1250000
Public Enum niDCPower_AttributeIDs
    NIDCPOWER_ATTR_SOURCE_DELAY = (1150000 + 51)
End Enum

' Measurement Functions
Public Enum niDCPower_SourceMode
    NIDCPOWER_VAL_SINGLE_POINT = 1020
    NIDCPOWER_VAL_SEQUENCE = 1021
End Enum

Public Enum niDCPower_OutputFunction
    NIDCPOWER_VAL_DC_VOLTAGE = 1006     'Sets the output function to DC voltage.
    NIDCPOWER_VAL_DC_CURRENT = 1007     'Sets the output function to DC current.
    NIDCPOWER_VAL_PULSE_VOLTAGE = 1049  'Sets the output function to pulse voltage.
    NIDCPOWER_VAL_PULSE_CURRENT = 1050  'Sets the output function to pulse current.
End Enum

'/*- Defined values for attributes -*/
'/*-   NIDCPOWER_ATTR_CURRENT_LIMIT_BEHAVIOR -*/
Public Enum niDCPower_CurrentLimitBehavior
    NIDCPOWER_VAL_CURRENT_REGULATE = 0
    NIDCPOWER_VAL_CURRENT_TRIP = 1
End Enum

Public Enum niDCPower_Events
    NIDCPOWER_VAL_SOURCE_COMPLETE_EVENT = 1030              'Waits for the Source Complete event.
    NIDCPOWER_VAL_MEASURE_COMPLETE_EVENT = 1031             'Waits for the Measure Complete event.
    NIDCPOWER_VAL_SEQUENCE_ITERATION_COMPLETE_EVENT = 1032  'Waits for the Sequence Iteration Complete event.
    NIDCPOWER_VAL_SEQUENCE_ENGINE_DONE_EVENT = 1033         'Waits for the Sequence Engine Done event.
    NIDCPOWER_VAL_PULSE_COMPLETE_EVENT = 1051               'Waits for the Pulse Complete event.
    NIDCPOWER_VAL_READY_FOR_PULSE_TRIGGER_EVENT = 1052      'Waits for the Ready for Pulse Trigger event.
End Enum

'ViStatus niDCPower_InitializeWithChannels(ViRsrc resourceName, ViConstString channels, ViBoolean reset, ViConstString optionString, ViSession *vi);
Private Declare PtrSafe Function niDCPower_InitializeWithChannels Lib "niDCPower_64" ( _
    ByVal resourceName As String, ByVal channels As String, ByVal Reset As Boolean, ByVal optionString As String, ByRef vi As Long) As Long

'ViStatus _VI_FUNC niDCPower_close(ViSession vi);
Private Declare PtrSafe Function niDCPower_close Lib "niDCPower_64" (ByVal vi As Long) As Long

'ViStatus _VI_FUNC niDCPower_reset(ViSession vi);
Private Declare PtrSafe Function niDCPower_reset Lib "niDCPower_64" (ByVal vi As Long) As Long

'ViStatus _VI_FUNC niDCPower_GetAttributeViInt32(ViSession vi, ViConstString channelName, ViAttr attributeId, ViInt32 *value);
Private Declare PtrSafe Function niDCPower_GetAttributeViInt32 Lib "niDCPower_64" ( _
    ByVal vi As Long, ByVal channelName As String, ByVal attributeID As Long, ByRef value As Long) As Long

'ViStatus _VI_FUNC niDCPower_SetAttributeViInt32(ViSession vi, ViConstString channelName, ViAttr attributeId, ViInt32 value);
Private Declare PtrSafe Function niDCPower_SetAttributeViInt32 Lib "niDCPower_64" ( _
    ByVal vi As Long, ByVal channelName As String, ByVal attributeID As Long, ByVal value As Long) As Long

'ViStatus _VI_FUNC niDCPower_GetAttributeViReal64(ViSession vi, ViConstString channelName, ViAttr attributeId, ViReal64 *value);
Private Declare PtrSafe Function niDCPower_GetAttributeViReal64 Lib "niDCPower_64" ( _
    ByVal vi As Long, ByVal channelName As String, ByVal attributeID As Long, ByRef value As Double) As Long

'ViStatus _VI_FUNC niDCPower_SetAttributeViReal64(ViSession vi, ViConstString channelName, ViAttr attributeId, ViReal64 value);
Private Declare PtrSafe Function niDCPower_SetAttributeViReal64 Lib "niDCPower_64" ( _
    ByVal vi As Long, ByVal channelName As String, ByVal attributeID As Long, ByVal value As Double) As Long

'ViStatus _VI_FUNC niDCPower_GetAttributeViString(ViSession vi, ViConstString channelName, ViAttr attributeId, ViInt32 bufSize, ViChar value[]);
Private Declare PtrSafe Function niDCPower_GetAttributeViString Lib "niDCPower_64" ( _
    ByVal vi As Long, ByVal channelName As String, ByVal attributeID As Long, ByVal bufSize As Long, ByVal value As LongPtr) As Long

'ViStatus _VI_FUNC niDCPower_SetAttributeViString(ViSession vi, ViConstString channelName, ViAttr attributeId, ViChar value[]);
Private Declare PtrSafe Function niDCPower_SetAttributeViString Lib "niDCPower_64" ( _
    ByVal vi As Long, ByVal channelName As String, ByVal attributeID As Long, ByVal value As String) As Long

'ViStatus _VI_FUNC niDCPower_GetAttributeViBoolean(ViSession vi, ViConstString channelName, ViAttr attributeId, ViBoolean *value);
Private Declare PtrSafe Function niDCPower_GetAttributeViBoolean Lib "niDCPower_64" ( _
    ByVal vi As Long, ByVal channelName As String, ByVal attributeID As Long, ByRef value As Boolean) As Long

'ViStatus _VI_FUNC niDCPower_SetAttributeViBoolean(ViSession vi, ViConstString channelName, ViAttr attributeId, ViBoolean value);
Private Declare PtrSafe Function niDCPower_SetAttributeViBoolean Lib "niDCPower_64" ( _
    ByVal vi As Long, ByVal channelName As String, ByVal attributeID As Long, ByVal value As Boolean) As Long

'ViStatus _VI_FUNC niDCPower_GetError(ViSession vi, ViStatus *errorCode, ViInt32 bufferSize, ViChar description[]);
Private Declare PtrSafe Function niDCPower_GetError Lib "niDCPower_64" ( _
    ByVal vi As Long, ByRef errorCode As Long, ByVal bufferSize As Long, ByVal errMessage As LongPtr) As Long

'ViStatus niDCPower_ConfigureSourceMode(ViSession vi, ViInt32 sourceMode);
Private Declare PtrSafe Function niDCPower_ConfigureSourceMode Lib "niDCPower_64" ( _
    ByVal vi As Long, ByVal sourceMode As Long) As Long
    
'ViStatus niDCPower_ConfigureOutputFunction(ViSession vi, ViConstString channelName, ViInt32 function);
Private Declare PtrSafe Function niDCPower_ConfigureOutputFunction Lib "niDCPower_64" ( _
    ByVal vi As Long, ByVal channelName As String, ByVal outputFunction As Long) As Long

'ViStatus niDCPower_ConfigureVoltageLevel(ViSession vi, ViConstString channelName, ViReal64 level);
Private Declare PtrSafe Function niDCPower_ConfigureVoltageLevel Lib "niDCPower_64" ( _
    ByVal vi As Long, ByVal channelName As String, ByVal level As Double) As Long

'ViStatus niDCPower_ConfigureVoltageLevelRange(ViSession vi, ViConstString channelName, ViReal64 range);
Private Declare PtrSafe Function niDCPower_ConfigureVoltageLevelRange Lib "niDCPower_64" ( _
    ByVal vi As Long, ByVal channelName As String, ByVal range As Double) As Long

'ViStatus niDCPower_ConfigureCurrentLimit(ViSession vi, ViConstString channelName, ViInt32 behavior, ViReal64 limit);
Private Declare PtrSafe Function niDCPower_ConfigureCurrentLimit Lib "niDCPower_64" ( _
    ByVal vi As Long, ByVal channelName As String, ByVal behavior As Long, ByVal limit As Double) As Long

'ViStatus niDCPower_ConfigureCurrentLimitRange(ViSession vi, ViConstString channelName, ViReal64 range);
Private Declare PtrSafe Function niDCPower_ConfigureCurrentLimitRange Lib "niDCPower_64" ( _
    ByVal vi As Long, ByVal channelName As String, ByVal range As Double) As Long

'ViStatus _VI_FUNC niDCPower_Commit(ViSession vi);
Private Declare PtrSafe Function niDCPower_Commit Lib "niDCPower_64" (ByVal vi As Long) As Long

'ViStatus _VI_FUNC niDCPower_Initiate(ViSession vi);
Private Declare PtrSafe Function niDCPower_Initiate Lib "niDCPower_64" (ByVal vi As Long) As Long

'ViStatus _VI_FUNC niDCPower_Abort(ViSession vi);
Private Declare PtrSafe Function niDCPower_Abort Lib "niDCPower_64" (ByVal vi As Long) As Long

'ViStatus _VI_FUNC niDCPower_WaitForEvent(ViSession vi, ViInt32 eventId, ViReal64 timeout);
Private Declare PtrSafe Function niDCPower_WaitForEvent Lib "niDCPower_64" ( _
    ByVal vi As Long, ByVal eventId As Long, ByVal timeout As Double) As Long
    
'ViStatus _VI_FUNC niDCPower_QueryInCompliance(ViSession vi, ViConstString channelName, ViBoolean *inCompliance);
Private Declare PtrSafe Function niDCPower_QueryInCompliance Lib "niDCPower_64" ( _
    ByVal vi As Long, ByVal channelName As String, ByRef inCompliance As Boolean) As Long

'ViStatus niDCPower_MeasureMultiple(ViSession vi, ViConstString channelName, ViReal64 voltageMeasurements[], ViReal64 currentMeasurements[]);
Private Declare PtrSafe Function niDCPower_MeasureMultiple Lib "niDCPower_64" ( _
    ByVal vi As Long, ByVal channelName As String, ByVal voltageMeasurements As LongPtr, ByVal currentMeasurements As LongPtr) As Long

' Internal session
Private m_Session As Long
Private m_ResourceName As String

' initialize internal variables, call Init first to create a valid session
Private Sub Class_Initialize()
    m_Session = 0
    m_ResourceName = ""
End Sub

' Automatically clear session when object gets destroyed
Private Sub Class_Terminate()
    CloseSession
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
 
    size = niDCPower_GetError(m_Session, errorCode, 0, 0)
    ReDim buffer(size) As Byte
 
    status = niDCPower_GetError(m_Session, errorCode, size, VarPtr(buffer(0)))
    errorMsg = StrConv(buffer(), vbUnicode)
    
    niTools_RaiseError errorCode, errorMsg, "NI-DCPower"
End Sub

Public Sub InitSession(resourceName As String, channels As String, Reset As Boolean, optionString As String)
    ' Make sure session is closed before opening
    CloseSession
    
    m_ResourceName = resourceName
    CheckError niDCPower_InitializeWithChannels(resourceName, channels, Reset, optionString, m_Session)
End Sub

Private Sub CloseSession()
    If m_Session = 0 Then Exit Sub
    
    CheckError niDCPower_close(m_Session)
    m_Session = 0
    m_ResourceName = ""
End Sub

Public Sub Reset()
    CheckError niDCPower_reset(m_Session)
End Sub

Public Sub GetAttributeLong(channelName As String, attributeID As niDCPower_AttributeIDs, ByRef value As Long)
    CheckError niDCPower_GetAttributeViInt32(m_Session, channelName, attributeID, value)
End Sub

Public Sub SetAttributeLong(channelName As String, attributeID As niDCPower_AttributeIDs, value As Long)
    CheckError niDCPower_SetAttributeViInt32(m_Session, channelName, attributeID, value)
End Sub

Public Sub GetAttributeDouble(channelName As String, attributeID As niDCPower_AttributeIDs, ByRef value As Double)
    CheckError niDCPower_GetAttributeViReal64(m_Session, channelName, attributeID, value)
End Sub

Public Sub SetAttributeDouble(channelName As String, attributeID As niDCPower_AttributeIDs, value As Double)
    CheckError niDCPower_SetAttributeViReal64(m_Session, channelName, attributeID, value)
End Sub

Public Sub GetAttributeBoolean(channelName As String, attributeID As niDCPower_AttributeIDs, ByRef value As Boolean)
    CheckError niDCPower_GetAttributeViBoolean(m_Session, channelName, attributeID, value)
End Sub

Public Sub SetAttributeBoolean(channelName As String, attributeID As niDCPower_AttributeIDs, value As Boolean)
    CheckError niDCPower_SetAttributeViBoolean(m_Session, channelName, attributeID, value)
End Sub

Public Sub GetAttributeString(channelName As String, attributeID As niDCPower_AttributeIDs, ByRef value As String)
    Dim size As Long
    Dim buffer() As Byte
    
    size = niDCPower_GetAttributeViString(m_Session, channelName, attributeID, 0, 0)
    ReDim buffer(size - 1) As Byte

    CheckError niDCPower_GetAttributeViString(m_Session, channelName, attributeID, size, VarPtr(buffer(0)))
    value = StrConv(LeftB(buffer(), size - 1), vbUnicode) ' Remove \0 character and convert to unicode
End Sub

Public Sub SetAttributeString(channelName As String, attributeID As niDCPower_AttributeIDs, value As String)
    CheckError niDCPower_SetAttributeViString(m_Session, channelName, attributeID, value)
End Sub

Public Sub ConfigureSourceMode(sourceMode As niDCPower_SourceMode)
    CheckError niDCPower_ConfigureSourceMode(m_Session, sourceMode)
End Sub

Public Sub ConfigureOutputFunction(channelName As String, outputFunction As niDCPower_OutputFunction)
    CheckError niDCPower_ConfigureOutputFunction(m_Session, channelName, outputFunction)
End Sub

Public Sub ConfigureVoltageLevel(channelName As String, level As Double)
    CheckError niDCPower_ConfigureVoltageLevel(m_Session, channelName, level)
End Sub

Public Sub ConfigureVoltageLevelRange(channelName As String, range As Double)
    CheckError niDCPower_ConfigureVoltageLevelRange(m_Session, channelName, range)
End Sub

Public Sub ConfigureCurrentLimit(channelName As String, limit As Double)
    CheckError niDCPower_ConfigureCurrentLimit(m_Session, channelName, NIDCPOWER_VAL_CURRENT_REGULATE, limit)
End Sub

Public Sub ConfigureCurrentLimitRange(channelName As String, range As Double)
    CheckError niDCPower_ConfigureCurrentLimitRange(m_Session, channelName, range)
End Sub

Public Sub Commit()
    CheckError niDCPower_Commit(m_Session)
End Sub

Public Sub Initiate()
    CheckError niDCPower_Initiate(m_Session)
End Sub

Public Sub Abort()
    CheckError niDCPower_Abort(m_Session)
End Sub

Public Sub WaitForEvent(eventId As niDCPower_Events, Optional timeout As Double = 10)
    CheckError niDCPower_WaitForEvent(m_Session, eventId, timeout)
End Sub

Public Sub QueryInCompliance(channelName As String, ByRef inCompliance As Boolean)
    CheckError niDCPower_QueryInCompliance(m_Session, channelName, inCompliance)
End Sub

Public Sub MeasureMultiple(channelName, ByRef voltageMeasurement() As Double, ByRef currentMeasurement() As Double)
     CheckError niDCPower_MeasureMultiple(m_Session, channelName, _
                    VarPtr(voltageMeasurement(LBound(voltageMeasurement))), _
                    VarPtr(currentMeasurement(LBound(currentMeasurement))))
End Sub
