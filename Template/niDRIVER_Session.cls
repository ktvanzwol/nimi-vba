VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "niDRIVER_Session"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Attribute IDs
Const NI<DRIVER>_ATTR_ As Long = 0 ' Attribute constant defined privately, class property should be used to expose the attribute

'ViStatus _VI_FUNC ni<DRIVER>_init(ViRsrc resourceName, ViBoolean IDQuery, ViBoolean resetDevice, ViSession* vi);
Private Declare PtrSafe Function ni<DRIVER>_init Lib "<DRIVER_DLL_NAME>" ( _
    ByVal resourceName As String, _
    ByVal IDQuery As Boolean, _
    ByVal resetDevice As Boolean, _
    ByRef vi As Long _
) As Long

'ViStatus _VI_FUNC ni<DRIVER>_close(ViSession vi);
Private Declare PtrSafe Function ni<DRIVER>_close Lib "<DRIVER_DLL_NAME>" ( _
    ByVal vi As Long _
) As Long

'ViStatus _VI_FUNC ni<DRIVER>_reset(ViSession vi);
Private Declare PtrSafe Function ni<DRIVER>_reset Lib "<DRIVER_DLL_NAME>" ( _
    ByVal vi As Long _
) As Long


'ViStatus _VI_FUNC ni<DRIVER>_GetAttributeViInt32(ViSession vi, ViConstString channelName, ViAttr attribute, ViInt32 *value);
Private Declare PtrSafe Function ni<DRIVER>_GetAttributeViInt32 Lib "<DRIVER_DLL_NAME>" ( _
    ByVal vi As Long, _
    ByVal channelName As String, _
    ByVal attributeID As Long, _
    ByRef value As Long _
) As Long

'ViStatus _VI_FUNC ni<DRIVER>_SetAttributeViInt32(ViSession vi, ViConstString channelName, ViAttr attribute, ViInt32 value);
Private Declare PtrSafe Function ni<DRIVER>_SetAttributeViInt32 Lib "<DRIVER_DLL_NAME>" ( _
    ByVal vi As Long, _
    ByVal channelName As String, _
    ByVal attributeID As Long, _
    ByVal value As Long _
) As Long

'ViStatus _VI_FUNC ni<DRIVER>_GetAttributeViReal64(ViSession vi, ViConstString channelName, ViAttr attribute, ViReal64 *value);
Private Declare PtrSafe Function ni<DRIVER>_GetAttributeViReal64 Lib "<DRIVER_DLL_NAME>" ( _
    ByVal vi As Long, _
    ByVal channelName As String, _
    ByVal attributeID As Long, _
    ByRef value As Double _
) As Long

'ViStatus _VI_FUNC ni<DRIVER>_SetAttributeViReal64(ViSession vi, ViConstString channelName, ViAttr attribute, ViReal64 value);
Private Declare PtrSafe Function ni<DRIVER>_SetAttributeViReal64 Lib "<DRIVER_DLL_NAME>" ( _
    ByVal vi As Long, _
    ByVal channelName As String, _
    ByVal attributeID As Long, _
    ByVal value As Double _
) As Long

'ViStatus _VI_FUNC ni<DRIVER>_GetAttributeViString(ViSession vi, ViConstString channelName, ViAttr attribute, ViInt32 bufSize, ViChar value[]);
Private Declare PtrSafe Function ni<DRIVER>_GetAttributeViString Lib "<DRIVER_DLL_NAME>" ( _
    ByVal vi As Long, _
    ByVal channelName As String, _
    ByVal attributeID As Long, _
    ByVal bufSize As Long, _
    ByVal value As LongPtr _
) As Long

'ViStatus _VI_FUNC ni<DRIVER>_SetAttributeViString(ViSession vi, ViConstString channelName, ViAttr attribute, ViConstString value);
Private Declare PtrSafe Function ni<DRIVER>_SetAttributeViString Lib "<DRIVER_DLL_NAME>" ( _
    ByVal vi As Long, _
    ByVal channelName As String, _
    ByVal attributeID As Long, _
    ByVal value As String _
) As Long

'ViStatus _VI_FUNC ni<DRIVER>_GetAttributeViBoolean(ViSession vi, ViConstString channelName, ViAttr attribute, ViBoolean *value);
Private Declare PtrSafe Function ni<DRIVER>_GetAttributeViBoolean Lib "<DRIVER_DLL_NAME>" ( _
    ByVal vi As Long, _
    ByVal channelName As String, _
    ByVal attributeID As Long, _
    ByRef value As Boolean _
) As Long

'ViStatus _VI_FUNC ni<DRIVER>_SetAttributeViBoolean(ViSession vi, ViConstString channelName, ViAttr attribute, ViBoolean value);
Private Declare PtrSafe Function ni<DRIVER>_SetAttributeViBoolean Lib "<DRIVER_DLL_NAME>" ( _
    ByVal vi As Long, _
    ByVal channelName As String, _
    ByVal attributeID As Long, _
    ByVal value As Boolean _
) As Long

'ViStatus _VI_FUNC ni<DRIVER>_GetError(ViSession vi, ViStatus *errorCode, ViInt32 bufferSize, ViChar description[]);
Private Declare PtrSafe Function niRFSA_GetError Lib "<DRIVER_DLL_NAME>" ( _
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
    
    size = ni<DRIVER>_GetError(m_Session, errorCode, 0, 0)
    buffer = String(size - 1, vbNullChar)
    
    status = ni<DRIVER>_GetError(m_Session, errorCode, size, StrPtr(buffer))
    niTools_CStrPtrToStr size, StrPtr(buffer), errorMsg
    
    niTools_RaiseError errorCode, errorMsg, "NI-<DRIVER>"
End Sub

Public Sub InitSession(resourceName As String, IDQuery As Boolean, reset As Boolean)
    ' Make sure session is closed before opening
    CloseSession
    
    m_ResourceName = resourceName
    CheckError ni<DRIVER>_init(resourceName, IDQuery, reset, m_Session)
End Sub

Private Sub CloseSession()
    If m_Session = 0 Then Exit Sub
    
    CheckError ni<DRIVER>_close(m_Session)
    m_Session = 0
    m_ResourceName = ""
End Sub

Public Sub reset()
    CheckError ni<DRIVER>_reset(m_Session)
End Sub

Public Property Get <PROPERTY_NAME>() As Double
    CheckError ni<DRIVER>_GetAttributeViReal64(m_Session, "", NI<DRIVER>_ATTR_, <PROPERTY_NAME>)
End Property

Public Property Let <PROPERTY_NAME>(ByVal value As Double)
    CheckError ni<DRIVER>_SetAttributeViReal64(m_Session, "", NI<DRIVER>_ATTR_, value)
End Property