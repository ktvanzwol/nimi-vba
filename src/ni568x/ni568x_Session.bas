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

'ViStatus _VI_FUNC ni568x_init(ViRsrc resourceName, ViBoolean IDQuery, ViBoolean resetDevice, ViSession* vi);
Private Declare PtrSafe Function ni568x_init Lib "ni568x_64" ( _
    ByVal resourceName As String, ByVal IDQuery As Boolean, ByVal resetDevice As Boolean, ByRef vi As Long) As Long

'ViStatus _VI_FUNC ni568x_close(ViSession vi);
Private Declare PtrSafe Function ni568x_close Lib "ni568x_64" ( _
    ByVal vi As Long) As Long

'ViStatus _VI_FUNC ni568x_reset(ViSession vi);
Private Declare PtrSafe Function ni568x_reset Lib "ni568x_64" ( _
    ByVal vi As Long) As Long

'ViStatus _VI_FUNC ni568x_ConfigureUnits(ViSession vi, ViInt32 units);
Private Declare PtrSafe Function ni568x_ConfigureUnits Lib "ni568x_64" ( _
    ByVal vi As Long, ByVal units As Long) As Long

'ViStatus _VI_FUNC ni568x_Read(ViSession vi, ViInt32 maxTimeMillisecond, ViReal64* power);
Private Declare PtrSafe Function ni568x_Read Lib "ni568x_64" ( _
    ByVal vi As Long, ByVal maxTimeMillisecond As Long, ByRef power As Double) As Long

'ViStatus _VI_FUNC ni568x_GetAttributeViInt32(ViSession vi, ViConstString channelName, ViAttr attribute, ViInt32 *value);
Private Declare PtrSafe Function ni568x_GetAttributeViInt32 Lib "ni568x_64" ( _
    ByVal vi As Long, ByVal channelName As String, ByVal attributeID As Long, ByRef value As Long) As Long

'ViStatus _VI_FUNC ni568x_SetAttributeViInt32(ViSession vi, ViConstString channelName, ViAttr attribute, ViInt32 value);
Private Declare PtrSafe Function ni568x_SetAttributeViInt32 Lib "ni568x_64" ( _
    ByVal vi As Long, ByVal channelName As String, ByVal attributeID As Long, ByVal value As Long) As Long

'ViStatus _VI_FUNC ni568x_GetAttributeViReal64(ViSession vi, ViConstString channelName, ViAttr attribute, ViReal64 *value);
Private Declare PtrSafe Function ni568x_GetAttributeViReal64 Lib "ni568x_64" ( _
    ByVal vi As Long, ByVal channelName As String, ByVal attributeID As Long, ByRef value As Double) As Long

'ViStatus _VI_FUNC ni568x_SetAttributeViReal64(ViSession vi, ViConstString channelName, ViAttr attribute, ViReal64 value);
Private Declare PtrSafe Function ni568x_SetAttributeViReal64 Lib "ni568x_64" ( _
    ByVal vi As Long, ByVal channelName As String, ByVal attributeID As Long, ByVal value As Double) As Long

'ViStatus _VI_FUNC ni568x_GetAttributeViString(ViSession vi, ViConstString channelName, ViAttr attribute, ViInt32 bufSize, ViChar value[]);
Private Declare PtrSafe Function ni568x_GetAttributeViString Lib "ni568x_64" ( _
    ByVal vi As Long, ByVal channelName As String, ByVal attributeID As Long, ByVal bufSize As Long, ByVal value As LongPtr) As Long

'ViStatus _VI_FUNC ni568x_SetAttributeViString(ViSession vi, ViConstString channelName, ViAttr attribute, ViConstString value);
Private Declare PtrSafe Function ni568x_SetAttributeViString Lib "ni568x_64" ( _
    ByVal vi As Long, ByVal channelName As String, ByVal attributeID As Long, ByVal value As String) As Long

'ViStatus _VI_FUNC ni568x_GetAttributeViBoolean(ViSession vi, ViConstString channelName, ViAttr attribute, ViBoolean *value);
Private Declare PtrSafe Function ni568x_GetAttributeViBoolean Lib "ni568x_64" ( _
    ByVal vi As Long, ByVal channelName As String, ByVal attributeID As Long, ByRef value As Boolean) As Long

'ViStatus _VI_FUNC ni568x_SetAttributeViBoolean(ViSession vi, ViConstString channelName, ViAttr attribute, ViBoolean value);
Private Declare PtrSafe Function ni568x_SetAttributeViBoolean Lib "ni568x_64" ( _
    ByVal vi As Long, ByVal channelName As String, ByVal attributeID As Long, ByVal value As Boolean) As Long

'ViStatus _VI_FUNC ni568x_GetError(ViSession vi, ViStatus *errorCode, ViInt32 bufferSize, ViChar description[]);
Private Declare PtrSafe Function ni568x_GetError Lib "ni568x_64" ( _
    ByVal vi As Long, ByRef errorCode As Long, ByVal bufferSize As Long, ByVal errMessage As LongPtr) As Long
    
'ViStatus _VI_FUNC ni568x_Zero(ViSession vi, ViConstString channelName);
Private Declare PtrSafe Function ni568x_Zero Lib "ni568x_64" ( _
    ByVal vi As Long, ByVal channelName As String) As Long

'ViStatus _VI_FUNC ni568x_ZeroAllChannels(ViSession vi);
Private Declare PtrSafe Function ni568x_ZeroAllChannels Lib "ni568x_64" ( _
    ByVal vi As Long) As Long

'ViStatus _VI_FUNC ni568x_IsZeroComplete( ViSession vi, ViInt32* zeroStatus );
Private Declare PtrSafe Function ni568x_IsZeroComplete Lib "ni568x_64" ( _
    ByVal vi As Long, ByRef zeroStatus As Long) As Long

' Internal session
Private m_Session As Long
Private m_ResourceName As String
Private m_Offset As Double

' initialize internal variables, call Init first to create a valid session
Private Sub Class_Initialize()
    m_Session = 0
    m_ResourceName = ""
    m_Offset = 0#
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
    ReDim buffer(size - 1) As Byte

    status = ni568x_GetError(m_Session, errorCode, 1024, VarPtr(buffer(0)))
    errorMsg = StrConv(LeftB(buffer(), size - 1), vbUnicode) 'Remove \0 character and convert to Unicode
        
    niTools_RaiseError errorCode, errorMsg, "NI-568x", m_ResourceName
End Sub

Public Sub InitSession(resourceName As String, IDQuery As Boolean, Reset As Boolean)
    ' Make sure session is closed before opening
    CloseSession
    
    m_ResourceName = resourceName
    CheckError ni568x_init(resourceName, IDQuery, Reset, m_Session)
End Sub

Private Sub CloseSession()
    If m_Session = 0 Then Exit Sub
    
    CheckError ni568x_close(m_Session)
    m_Session = 0
    m_ResourceName = ""
End Sub

Public Sub Reset()
    CheckError ni568x_reset(m_Session)
End Sub

Public Sub Read(ByRef power As Double, Optional maxTimeMillisecond As Long = 5000)
    CheckError ni568x_Read(m_Session, maxTimeMillisecond, power)
End Sub

Public Sub ConfigureUnits(units As ni568x_Units)
    CheckError ni568x_ConfigureUnits(m_Session, units)
End Sub

Public Sub GetAttributeViInt32(channelName As String, attributeID As ni568x_AttributeIDs, ByRef value As Long)
    CheckError ni568x_GetAttributeViInt32(m_Session, channelName, attributeID, value)
End Sub

Public Sub SetAttributeViInt32(channelName As String, attributeID As ni568x_AttributeIDs, value As Long)
    CheckError ni568x_SetAttributeViInt32(m_Session, channelName, attributeID, value)
End Sub

Public Sub GetAttributeViReal64(channelName As String, attributeID As ni568x_AttributeIDs, ByRef value As Double)
    CheckError ni568x_GetAttributeViReal64(m_Session, channelName, attributeID, value)
End Sub

Public Sub SetAttributeViReal64(channelName As String, attributeID As ni568x_AttributeIDs, value As Double)
    CheckError ni568x_SetAttributeViReal64(m_Session, channelName, attributeID, value)
End Sub

Public Sub GetAttributeViBoolean(channelName As String, attributeID As ni568x_AttributeIDs, ByRef value As Boolean)
    CheckError ni568x_GetAttributeViBoolean(m_Session, channelName, attributeID, value)
End Sub

Public Sub SetAttributeViBoolean(channelName As String, attributeID As ni568x_AttributeIDs, value As Boolean)
    CheckError ni568x_SetAttributeViBoolean(m_Session, channelName, attributeID, value)
End Sub

Public Sub GetAttributeViString(channelName As String, attributeID As ni568x_AttributeIDs, ByRef value As String)
    Dim size As Long
    Dim buffer() As Byte
    
    size = ni568x_GetAttributeViString(m_Session, channelName, attributeID, 0, 0)
    ReDim buffer(size - 1) As Byte

    CheckError ni568x_GetAttributeViString(m_Session, channelName, attributeID, size, VarPtr(buffer(0)))
    value = StrConv(LeftB(buffer(), size - 1), vbUnicode) ' Remove \0 character and convert to unicode
End Sub

Public Sub SetAttributeViString(channelName As String, attributeID As ni568x_AttributeIDs, value As String)
    CheckError ni568x_SetAttributeViString(m_Session, channelName, attributeID, value)
End Sub

Public Sub Zero()
    CheckError ni568x_ZeroAllChannels(m_Session)
End Sub

Public Sub IsZeroCompleted(ByRef zeroStatus As ni568x_ZeroStatus)
    CheckError ni568x_IsZeroComplete(m_Session, zeroStatus)
End Sub

Public Sub DisableOffset()
    GetAttributeViReal64 "", NI568X_ATTR_OFFSET, m_Offset
    SetAttributeViReal64 "", NI568X_ATTR_OFFSET, 0#
End Sub

Public Sub EnableOffset()
    SetAttributeViReal64 "", NI568X_ATTR_OFFSET, m_Offset
End Sub

