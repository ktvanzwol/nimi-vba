VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "niDMM_Session"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Attribute IDs
' Note: in header files
' IVI_SPECIFIC_PUBLIC_ATTR_BASE = 1150000
' IVI_CLASS_PUBLIC_ATTR_BASE = 1250000
Public Enum niDMM_AttributeIDs
    NIDMM_ATTR_RESOLUTION_ABSOLUTE = (1250000 + 8)
    NIDMM_ATTR_POWERLINE_FREQ = (1250000 + 333)
End Enum

' Measurement Functions
Public Enum niDMM_MeasurementFunction
    NIDMM_VAL_DC_VOLTS = 1                'DC Voltage                     All
    NIDMM_VAL_AC_VOLTS = 2                'AC Voltage with AC Coupling    All
    NIDMM_VAL_DC_CURRENT = 3              'DC Current                     All
    NIDMM_VAL_AC_CURRENT = 4              'AC Current                     All
    NIDMM_VAL_2_WIRE_RES = 5              '2-Wire Resistance              All
    NIDMM_VAL_4_WIRE_RES = 101            '4-Wire Resistance              NI 4060, NI 4065, NI 4070/4071/4072, NI 4080/4081/4082
    NIDMM_VAL_FREQ = 104                  'Frequency                      NI 4070/4071/4072 and NI 4080/4081/4082
    NIDMM_VAL_PERIOD = 105                'Period                         NI 4070/4071/4072 and NI 4080/4081/4082
    NIDMM_VAL_TEMPERATURE = 108           'Temperature
    NIDMM_VAL_AC_VOLTS_DC_COUPLED = 1001  'AC Voltage with DC Coupling    NI 4070/4071/4072 and NI 4080/4081/4082
    NIDMM_VAL_DIODE = 1002                'Diode                          All
    NIDMM_VAL_WAVEFORM_VOLTAGE = 1003     'Waveform Voltage               NI 4070/4071/4072 and NI 4080/4081/4082
    NIDMM_VAL_WAVEFORM_CURRENT = 1004     'Waveform Current               NI 4070/4071/4072 and NI 4080/4081/4082
    NIDMM_VAL_CAPACITANCE = 1005          'Capacitance                    NI 4072 and NI 4082
    NIDMM_VAL_INDUCTANCE = 1006           'Inductance                     NI 4072 and NI 4082
End Enum


'ViStatus _VI_FUNC niDMM_init(ViRsrc resourceName, ViBoolean IDQuery, ViBoolean reset, ViSession *newVi);
Private Declare PtrSafe Function niDMM_init Lib "niDMM_64" ( _
    ByVal resourceName As String, ByVal IDQuery As Boolean, ByVal reset As Boolean, ByRef newVi As Long) As Long

'ViStatus _VI_FUNC niDMM_close(ViSession vi);
Private Declare PtrSafe Function niDMM_close Lib "niDMM_64" (ByVal vi As Long) As Long

'ViStatus _VI_FUNC niDMM_reset(ViSession vi);
Private Declare PtrSafe Function niDMM_reset Lib "niDMM_64" (ByVal vi As Long) As Long

'ViStatus _VI_FUNC niDMM_SelfCal(ViSession vi);
Private Declare PtrSafe Function niDMM_SelfCal Lib "niDMM_64" (ByVal vi As Long) As Long

'ViStatus _VI_FUNC niDMM_ConfigureMeasurementDigits(ViSession vi, ViInt32 measFunction, ViReal64 range, ViReal64 resolutionDigits);
Private Declare PtrSafe Function niDMM_ConfigureMeasurementDigits Lib "niDMM_64" ( _
    ByVal vi As Long, ByVal measFunction As Long, ByVal range As Double, ByVal resolutionDigits As Double) As Long

'ViStatus _VI_FUNC niDMM_ConfigureMeasurementAbsolute(ViSession vi, ViInt32 measFunction, ViReal64 range, ViReal64    resolutionAbsolute);
Private Declare PtrSafe Function niDMM_ConfigureMeasurementAbsolute Lib "niDMM_64" ( _
    ByVal vi As Long, ByVal measFunction As Long, ByVal range As Double, ByVal resolutionAbsolute As Double) As Long

'ViStatus _VI_FUNC niDMM_Read(ViSession vi, ViInt32 maxTime, ViReal64 *reading);
Private Declare PtrSafe Function niDMM_Read Lib "niDMM_64" ( _
    ByVal vi As Long, ByVal maxTime As Long, ByRef reading As Double) As Long

'ViStatus _VI_FUNC niDMM_GetAttributeViInt32(ViSession vi, ViConstString channelName, ViAttr attributeId, ViInt32 *value);
Private Declare PtrSafe Function niDMM_GetAttributeViInt32 Lib "niDMM_64" ( _
    ByVal vi As Long, ByVal channelName As String, ByVal attributeID As Long, ByRef value As Long) As Long

'ViStatus _VI_FUNC niDMM_SetAttributeViInt32(ViSession vi, ViConstString channelName, ViAttr attributeId, ViInt32 value);
Private Declare PtrSafe Function niDMM_SetAttributeViInt32 Lib "niDMM_64" ( _
    ByVal vi As Long, ByVal channelName As String, ByVal attributeID As Long, ByVal value As Long) As Long

'ViStatus _VI_FUNC niDMM_GetAttributeViReal64(ViSession vi, ViConstString channelName, ViAttr attributeId, ViReal64 *value);
Private Declare PtrSafe Function niDMM_GetAttributeViReal64 Lib "niDMM_64" ( _
    ByVal vi As Long, ByVal channelName As String, ByVal attributeID As Long, ByRef value As Double) As Long

'ViStatus _VI_FUNC niDMM_SetAttributeViReal64(ViSession vi, ViConstString channelName, ViAttr attributeId, ViReal64 value);
Private Declare PtrSafe Function niDMM_SetAttributeViReal64 Lib "niDMM_64" ( _
    ByVal vi As Long, ByVal channelName As String, ByVal attributeID As Long, ByVal value As Double) As Long

'ViStatus _VI_FUNC niDMM_GetAttributeViString(ViSession vi, ViConstString channelName, ViAttr attributeId, ViInt32 bufSize, ViChar value[]);
Private Declare PtrSafe Function niDMM_GetAttributeViString Lib "niDMM_64" ( _
    ByVal vi As Long, ByVal channelName As String, ByVal attributeID As Long, ByVal bufSize As Long, ByVal value As LongPtr) As Long

'ViStatus _VI_FUNC niDMM_SetAttributeViString(ViSession vi, ViConstString channelName, ViAttr attributeId, ViChar value[]);
Private Declare PtrSafe Function niDMM_SetAttributeViString Lib "niDMM_64" ( _
    ByVal vi As Long, ByVal channelName As String, ByVal attributeID As Long, ByVal value As String) As Long

'ViStatus _VI_FUNC niDMM_GetAttributeViBoolean(ViSession vi, ViConstString channelName, ViAttr attributeId, ViBoolean *value);
Private Declare PtrSafe Function niDMM_GetAttributeViBoolean Lib "niDMM_64" ( _
    ByVal vi As Long, ByVal channelName As String, ByVal attributeID As Long, ByRef value As Boolean) As Long

'ViStatus _VI_FUNC niDMM_SetAttributeViBoolean(ViSession vi, ViConstString channelName, ViAttr attributeId, ViBoolean value);
Private Declare PtrSafe Function niDMM_SetAttributeViBoolean Lib "niDMM_64" ( _
    ByVal vi As Long, ByVal channelName As String, ByVal attributeID As Long, ByVal value As Boolean) As Long

'ViStatus _VI_FUNC niDMM_GetError(ViSession vi, ViStatus *errorCode, ViInt32 bufferSize, ViChar description[]);
Private Declare PtrSafe Function niDMM_GetError Lib "niDMM_64" ( _
    ByVal vi As Long, ByRef errorCode As Long, ByVal bufferSize As Long, ByVal errMessage As LongPtr) As Long

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
 
    size = niDMM_GetError(m_Session, errorCode, 0, 0)
    ReDim buffer(size) As Byte
 
    status = niDMM_GetError(m_Session, errorCode, size, VarPtr(buffer(0)))
    errorMsg = StrConv(buffer(), vbUnicode)
    
    niTools_RaiseError errorCode, errorMsg, "NI-DMM"
End Sub

Public Sub InitSession(resourceName As String, IDQuery As Boolean, reset As Boolean)
    ' Make sure session is closed before opening
    CloseSession
    
    m_ResourceName = resourceName
    CheckError niDMM_init(resourceName, IDQuery, reset, m_Session)
End Sub

Private Sub CloseSession()
    If m_Session = 0 Then Exit Sub
    
    CheckError niDMM_close(m_Session)
    m_Session = 0
    m_ResourceName = ""
End Sub

Public Sub reset()
    CheckError niDMM_reset(m_Session)
End Sub

Public Sub SelfCal()
    CheckError niDMM_SelfCal(m_Session)
End Sub

Public Sub ConfigureMeasurementDigits(measFunction As niDMM_MeasurementFunction, range As Double, resolutionDigits As Double)
    CheckError niDMM_ConfigureMeasurementDigits(m_Session, measFunction, range, resolutionDigits)
End Sub

Public Sub ConfigureMeasurementAbsolute(measFunction As niDMM_MeasurementFunction, range As Double, resolutionAbsolute As Double)
    CheckError niDMM_ConfigureMeasurementAbsolute(m_Session, measFunction, range, resolutionAbsolute)
End Sub

Public Sub Read(ByRef reading As Double, Optional maxTime As Long = NIDMM_VAL_TIME_LIMIT_AUTO)
    CheckError niDMM_Read(m_Session, maxTime, reading)
End Sub

Public Sub GetAttributeLong(channelName As String, attributeID As niDMM_AttributeIDs, ByRef value As Long)
    CheckError niDMM_GetAttributeViInt32(m_Session, channelName, attributeID, value)
End Sub

Public Sub SetAttributeLong(channelName As String, attributeID As niDMM_AttributeIDs, value As Long)
    CheckError niDMM_SetAttributeViInt32(m_Session, channelName, attributeID, value)
End Sub

Public Sub GetAttributeDouble(channelName As String, attributeID As niDMM_AttributeIDs, ByRef value As Double)
    CheckError niDMM_GetAttributeViReal64(m_Session, channelName, attributeID, value)
End Sub

Public Sub SetAttributeDouble(channelName As String, attributeID As niDMM_AttributeIDs, value As Double)
    CheckError niDMM_SetAttributeViReal64(m_Session, channelName, attributeID, value)
End Sub

Public Sub GetAttributeBoolean(channelName As String, attributeID As niDMM_AttributeIDs, ByRef value As Boolean)
    CheckError niDMM_GetAttributeViBoolean(m_Session, channelName, attributeID, value)
End Sub

Public Sub SetAttributeBoolean(channelName As String, attributeID As niDMM_AttributeIDs, value As Boolean)
    CheckError niDMM_SetAttributeViBoolean(m_Session, channelName, attributeID, value)
End Sub

Public Sub GetAttributeString(channelName As String, attributeID As niDMM_AttributeIDs, ByRef value As String)
    Dim size As Long
    Dim buffer() As Byte
    
    size = niDMM_GetAttributeViString(m_Session, channelName, attributeID, 0, 0)
    ReDim buffer(size - 1) As Byte

    CheckError niDMM_GetAttributeViString(m_Session, channelName, attributeID, size, VarPtr(buffer(0)))
    value = StrConv(LeftB(buffer(), size - 1), vbUnicode) ' Remove \0 character and convert to unicode
End Sub

Public Sub SetAttributeString(channelName As String, attributeID As niDMM_AttributeIDs, value As String)
    CheckError niDMM_SetAttributeViString(m_Session, channelName, attributeID, value)
End Sub

' Mapping properties to Get/Set Attribute function for compatibity
' This can always be done to improve usability but its adds overhead to adding support for attributes.
Public Property Get Powerline_Freq() As Double
    GetAttributeDouble "", NIDMM_ATTR_POWERLINE_FREQ, Powerline_Freq
End Property

Public Property Let Powerline_Freq(ByVal value As Double)
    SetAttributeDouble "", NIDMM_ATTR_POWERLINE_FREQ, value
End Property

Public Property Get Resolution_Absolute() As Double
    GetAttributeDouble "", NIDMM_ATTR_RESOLUTION_ABSOLUTE, Resolution_Absolute
End Property

Public Property Let Resolution_Absolute(ByVal value As Double)
    SetAttributeDouble "", NIDMM_ATTR_RESOLUTION_ABSOLUTE, value
End Property



