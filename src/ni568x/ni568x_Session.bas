VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ni568x_Session"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Attribute IDs
Const NI568X_ATTR_ As Long = 0 ' dummy

'ViStatus _VI_FUNC ni568x_init(ViRsrc resourceName, ViBoolean IDQuery, ViBoolean resetDevice, ViSession* vi);
Private Declare PtrSafe Function ni568x_init Lib "ni568x_64" ( _
    ByVal resourceName As String, _
    ByVal IDQuery As Boolean, _
    ByVal resetDevice As Boolean, _
    ByRef vi As Long _
) As Long

'ViStatus _VI_FUNC ni568x_close(ViSession vi);
Private Declare PtrSafe Function ni568x_close Lib "ni568x_64" ( _
    ByVal vi As Long _
) As Long

'ViStatus _VI_FUNC ni568x_reset(ViSession vi);
Private Declare PtrSafe Function ni568x_reset Lib "ni568x_64" ( _
    ByVal vi As Long _
) As Long

'ViStatus _VI_FUNC ni568x_ConfigureUnits(ViSession vi, ViInt32 units);
Private Declare PtrSafe Function ni568x_ConfigureUnits Lib "ni568x_64" ( _
    ByVal vi As Long, _
    ByVal units As Long _
) As Long

'ViStatus _VI_FUNC ni568x_Read(ViSession vi, ViInt32 maxTimeMillisecond, ViReal64* power);
Private Declare PtrSafe Function ni568x_Read Lib "ni568x_64" ( _
    ByVal vi As Long, _
    ByVal maxTimeMillisecond As Long, _
    ByRef power As Double _
) As Long

'ViStatus _VI_FUNC ni568x_GetAttributeViInt32(ViSession vi, ViConstString channelName, ViAttr attribute, ViInt32 *value);
Private Declare PtrSafe Function ni568x_GetAttributeViInt32 Lib "ni568x_64" ( _
    ByVal vi As Long, _
    ByVal channelName As String, _
    ByVal attributeID As Long, _
    ByRef value As Long _
) As Long

'ViStatus _VI_FUNC ni568x_SetAttributeViInt32(ViSession vi, ViConstString channelName, ViAttr attribute, ViInt32 value);
Private Declare PtrSafe Function ni568x_SetAttributeViInt32 Lib "ni568x_64" ( _
    ByVal vi As Long, _
    ByVal channelName As String, _
    ByVal attributeID As Long, _
    ByVal value As Long _
) As Long

'ViStatus _VI_FUNC ni568x_GetAttributeViReal64(ViSession vi, ViConstString channelName, ViAttr attribute, ViReal64 *value);
Private Declare PtrSafe Function ni568x_GetAttributeViReal64 Lib "ni568x_64" ( _
    ByVal vi As Long, _
    ByVal channelName As String, _
    ByVal attributeID As Long, _
    ByRef value As Double _
) As Long

'ViStatus _VI_FUNC ni568x_SetAttributeViReal64(ViSession vi, ViConstString channelName, ViAttr attribute, ViReal64 value);
Private Declare PtrSafe Function ni568x_SetAttributeViReal64 Lib "ni568x_64" ( _
    ByVal vi As Long, _
    ByVal channelName As String, _
    ByVal attributeID As Long, _
    ByVal value As Double _
) As Long

'ViStatus _VI_FUNC ni568x_GetAttributeViString(ViSession vi, ViConstString channelName, ViAttr attribute, ViInt32 bufSize, ViChar value[]);
Private Declare PtrSafe Function ni568x_GetAttributeViString Lib "ni568x_64" ( _
    ByVal vi As Long, _
    ByVal channelName As String, _
    ByVal attributeID As Long, _
    ByVal bufSize As Long, _
    ByVal value As LongPtr _
) As Long

'ViStatus _VI_FUNC ni568x_SetAttributeViString(ViSession vi, ViConstString channelName, ViAttr attribute, ViConstString value);
Private Declare PtrSafe Function ni568x_SetAttributeViString Lib "ni568x_64" ( _
    ByVal vi As Long, _
    ByVal channelName As String, _
    ByVal attributeID As Long, _
    ByVal value As String _
) As Long

'ViStatus _VI_FUNC ni568x_GetAttributeViBoolean(ViSession vi, ViConstString channelName, ViAttr attribute, ViBoolean *value);
Private Declare PtrSafe Function ni568x_GetAttributeViBoolean Lib "ni568x_64" ( _
    ByVal vi As Long, _
    ByVal channelName As String, _
    ByVal attributeID As Long, _
    ByRef value As Boolean _
) As Long

'ViStatus _VI_FUNC ni568x_SetAttributeViBoolean(ViSession vi, ViConstString channelName, ViAttr attribute, ViBoolean value);
Private Declare PtrSafe Function ni568x_SetAttributeViBoolean Lib "ni568x_64" ( _
    ByVal vi As Long, _
    ByVal channelName As String, _
    ByVal attributeID As Long, _
    ByVal value As Boolean _
) As Long

'ViStatus _VI_FUNC ni568x_GetError(ViSession vi, ViStatus *errorCode, ViInt32 bufferSize, ViChar description[]);
Private Declare PtrSafe Function ni568x_GetError Lib "ni568x_64" ( _
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
    Dim buffer() As Byte
    Dim errorMsg As String
        
    size = ni568x_GetError(m_Session, errorCode, 0, 0)
    ReDim buffer(size)

    status = ni568x_GetError(m_Session, errorCode, 1024, VarPtr(buffer(0)))
    errorMsg = StrConv(buffer(), vbUnicode)
        
    niTools_RaiseError errorCode, errorMsg, "NI-568x", m_ResourceName
End Sub

Public Sub InitSession(resourceName As String, IDQuery As Boolean, reset As Boolean)
    ' Make sure session is closed before opening
    CloseSession
    
    m_ResourceName = resourceName
    CheckError ni568x_init(resourceName, IDQuery, reset, m_Session)
End Sub

Private Sub CloseSession()
    If m_Session = 0 Then Exit Sub
    
    CheckError ni568x_close(m_Session)
    m_Session = 0
    m_ResourceName = ""
End Sub

Public Sub reset()
    CheckError ni568x_reset(m_Session)
End Sub

Public Sub Read(ByRef power As Double, Optional maxTimeMillisecond As Long = 5000)
    CheckError ni568x_Read(m_Session, maxTimeMillisecond, power)
End Sub

Public Sub ConfigureUnits(units As Long)
    CheckError ni568x_ConfigureUnits(m_Session, units)
End Sub
