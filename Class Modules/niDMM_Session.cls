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
Const NIDMM_ATTR_RESOLUTION_ABSOLUTE As Long = 1250008
Const NIDMM_ATTR_POWERLINE_FREQ As Long = 1250333

'ViStatus _VI_FUNC niDMM_init(
'    ViRsrc      resourceName,
'    ViBoolean   IDQuery,
'    ViBoolean   reset,
'    ViSession   *newVi);
Private Declare PtrSafe Function niDMM_init Lib "niDMM_64" ( _
    ByVal resourceName As String, _
    ByVal IDQuery As Boolean, _
    ByVal reset As Boolean, _
    ByRef newVi As Long _
) As Long

'ViStatus _VI_FUNC niDMM_close(
'    ViSession   vi);
Private Declare PtrSafe Function niDMM_close Lib "niDMM_64" ( _
    ByVal vi As Long _
) As Long

'ViStatus _VI_FUNC niDMM_reset(
'    ViSession   vi);
Private Declare PtrSafe Function niDMM_reset Lib "niDMM_64" ( _
    ByVal vi As Long _
) As Long

'ViStatus _VI_FUNC niDMM_SelfCal(
'    ViSession   vi);
Private Declare PtrSafe Function niDMM_SelfCal Lib "niDMM_64" ( _
    ByVal vi As Long _
) As Long


'ViStatus _VI_FUNC niDMM_ConfigureMeasurementDigits(
'   ViSession   vi,
'   ViInt32     measFunction,
'   ViReal64    range,
'   ViReal64    resolutionDigits);
Private Declare PtrSafe Function niDMM_ConfigureMeasurementDigits Lib "niDMM_64" ( _
    ByVal vi As Long, _
    ByVal measFunction As Long, _
    ByVal range As Double, _
    ByVal resolutionDigits As Double _
) As Long

'ViStatus _VI_FUNC niDMM_ConfigureMeasurementAbsolute(
'   ViSession   vi,
'   ViInt32     measFunction,
'   ViReal64    range,
'   ViReal64    resolutionAbsolute);
Private Declare PtrSafe Function niDMM_ConfigureMeasurementAbsolute Lib "niDMM_64" ( _
    ByVal vi As Long, _
    ByVal measFunction As Long, _
    ByVal range As Double, _
    ByVal resolutionAbsolute As Double _
) As Long

'ViStatus _VI_FUNC niDMM_Read(
'    ViSession   vi,
'    ViInt32     maxTime,
'    ViReal64    *reading);
Private Declare PtrSafe Function niDMM_Read Lib "niDMM_64" ( _
    ByVal vi As Long, _
    ByVal maxTime As Long, _
    ByRef reading As Double _
) As Long


'ViStatus _VI_FUNC niDMM_GetAttributeViInt32(
'    ViSession     vi,
'    ViConstString channelName,
'    ViAttr        attributeId,
'    ViInt32       *value);
Private Declare PtrSafe Function niDMM_GetAttributeViInt32 Lib "niDMM_64" ( _
    ByVal vi As Long, _
    ByVal channelName As String, _
    ByVal attributeID As Long, _
    ByRef value As Long _
) As Long

'ViStatus _VI_FUNC niDMM_SetAttributeViInt32(
'    ViSession     vi,
'    ViConstString channelName,
'    ViAttr        attributeId,
'    ViInt32       value);
Private Declare PtrSafe Function niDMM_SetAttributeViInt32 Lib "niDMM_64" ( _
    ByVal vi As Long, _
    ByVal channelName As String, _
    ByVal attributeID As Long, _
    ByVal value As Long _
) As Long

'ViStatus _VI_FUNC niDMM_GetAttributeViReal64(
'    ViSession     vi,
'    ViConstString channelName,
'    ViAttr        attributeId,
'    ViReal64      *value);
Private Declare PtrSafe Function niDMM_GetAttributeViReal64 Lib "niDMM_64" ( _
    ByVal vi As Long, _
    ByVal channelName As String, _
    ByVal attributeID As Long, _
    ByRef value As Double _
) As Long

'ViStatus _VI_FUNC niDMM_SetAttributeViReal64(
'    ViSession     vi,
'    ViConstString channelName,
'    ViAttr        attributeId,
'    ViReal64      value);
Private Declare PtrSafe Function niDMM_SetAttributeViReal64 Lib "niDMM_64" ( _
    ByVal vi As Long, _
    ByVal channelName As String, _
    ByVal attributeID As Long, _
    ByVal value As Double _
) As Long

'ViStatus _VI_FUNC niDMM_GetAttributeViString(
'    ViSession     vi,
'    ViConstString channelName,
'    ViAttr        attributeId,
'    ViInt32       bufSize,
'    ViChar        value[]);
Private Declare PtrSafe Function niDMM_GetAttributeViString Lib "niDMM_64" ( _
    ByVal vi As Long, _
    ByVal channelName As String, _
    ByVal attributeID As Long, _
    ByVal bufSize As Long, _
    ByVal value As LongPtr _
) As Long

'ViStatus _VI_FUNC niDMM_SetAttributeViString(
'    ViSession     vi,
'    ViConstString channelName,
'    ViAttr        attributeId,
'    ViChar        value[]);
Private Declare PtrSafe Function niDMM_SetAttributeViString Lib "niDMM_64" ( _
    ByVal vi As Long, _
    ByVal channelName As String, _
    ByVal attributeID As Long, _
    ByVal value As String _
) As Long

'ViStatus _VI_FUNC niDMM_GetAttributeViBoolean(
'    ViSession     vi,
'    ViConstString channelName,
'    ViAttr        attributeId,
'    ViBoolean     *value);
Private Declare PtrSafe Function niDMM_GetAttributeViBoolean Lib "niDMM_64" ( _
    ByVal vi As Long, _
    ByVal channelName As String, _
    ByVal attributeID As Long, _
    ByRef value As Boolean _
) As Long

'ViStatus _VI_FUNC niDMM_SetAttributeViBoolean(
'    ViSession     vi,
'    ViConstString channelName,
'    ViAttr        attributeId,
'    ViBoolean     value);
Private Declare PtrSafe Function niDMM_SetAttributeViBoolean Lib "niDMM_64" ( _
    ByVal vi As Long, _
    ByVal channelName As String, _
    ByVal attributeID As Long, _
    ByVal value As Boolean _
) As Long

'ViStatus _VI_FUNC niDMM_GetError(ViSession vi, ViStatus *errorCode, ViInt32 bufferSize, ViChar description[]);
Private Declare PtrSafe Function niDMM_GetError Lib "niDMM_64" ( _
    ByVal vi As Long, _
    ByRef errorCode As Long, _
    ByVal bufferSize As Long, _
    ByVal errMessage As LongPtr _
) As Long

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
    Dim buffer As String
    Dim errorMsg As String
 
    size = niDMM_GetError(m_Session, errorCode, 0, 0)
    buffer = String(size - 1, vbNullChar)
 
    status = niDMM_GetError(m_Session, errorCode, size, StrPtr(buffer))
    niTools_CStrPtrToStr size, StrPtr(buffer), errorMsg
    
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

Public Sub ConfigureMeasurementDigits(measFunction As Long, range As Double, resolutionDigits As Double)
    CheckError niDMM_ConfigureMeasurementDigits(m_Session, measFunction, range, resolutionDigits)
End Sub

Public Sub ConfigureMeasurementAbsolute(measFunction As Long, range As Double, resolutionAbsolute As Double)
    CheckError niDMM_ConfigureMeasurementAbsolute(m_Session, measFunction, range, resolutionAbsolute)
End Sub

Public Sub Read(ByRef reading As Double, Optional maxTime As Long = NIDMM_VAL_TIME_LIMIT_AUTO)
    CheckError niDMM_Read(m_Session, maxTime, reading)
End Sub

Public Property Get Powerline_Freq() As Double
    CheckError niDMM_GetAttributeViReal64(m_Session, "", NIDMM_ATTR_POWERLINE_FREQ, Powerline_Freq)
End Property

Public Property Let Powerline_Freq(ByVal value As Double)
    CheckError niDMM_SetAttributeViReal64(m_Session, "", NIDMM_ATTR_POWERLINE_FREQ, value)
End Property

Public Property Get Resolution_Absolute() As Double
    CheckError niDMM_GetAttributeViReal64(m_Session, "", NIDMM_ATTR_RESOLUTION_ABSOLUTE, Resolution_Absolute)
End Property

Public Property Let Resolution_Absolute(ByVal value As Double)
    CheckError niDMM_SetAttributeViReal64(m_Session, "", NIDMM_ATTR_RESOLUTION_ABSOLUTE, value)
End Property

