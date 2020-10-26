VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "niRFSG_Session"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Attribute IDs
' Note: in header files IVI_SPECIFIC_PUBLIC_ATTR_BASE = 1150000
Public Enum niRFSG_AttributeIDs
    NIRFSG_ATTR_EXTERNAL_GAIN = (1150000 + &H55)
End Enum

' - Values for attribute NIRFSG_ATTR_GENERATION_MODE -
Public Enum niRFSG_GenerationMode
    NIRFSG_VAL_CW = 1000
    NIRFSG_VAL_ARB_WAVEFORM = 1001
    NIRFSG_VAL_SCRIPT = 1002
End Enum

'ViStatus _VI_FUNC niRFSG_init(ViRsrc resourceName, ViBoolean IDQuery, ViBoolean resetDevice, ViSession* vi);
Private Declare PtrSafe Function niRFSG_init Lib "niRFSG_64" ( _
    ByVal resourceName As String, ByVal IDQuery As Boolean, ByVal resetDevice As Boolean, ByRef vi As Long) As Long

'ViStatus _VI_FUNC niRFSG_InitWithOptions(ViRsrc resourceName, ViBoolean IDQuery, ViBoolean reset, ViConstString optionString, ViSession* newVi);
Private Declare PtrSafe Function niRFSG_InitWithOptions Lib "niRFSG_64" ( _
    ByVal resourceName As String, ByVal IDQuery As Boolean, ByVal resetDevice As Boolean, ByVal optionString As String, ByRef vi As Long) As Long

'ViStatus _VI_FUNC niRFSG_close(ViSession vi);
Private Declare PtrSafe Function niRFSG_close Lib "niRFSG_64" ( _
    ByVal vi As Long) As Long

'ViStatus _VI_FUNC niRFSG_reset(ViSession vi);
Private Declare PtrSafe Function niRFSG_reset Lib "niRFSG_64" ( _
    ByVal vi As Long) As Long

'ViStatus _VI_FUNC niRFSG_SelfCal(ViSession vi);
Private Declare PtrSafe Function niRFSG_SelfCal Lib "niRFSG_64" ( _
    ByVal vi As Long) As Long

'ViStatus _VI_FUNC niRFSG_GetAttributeViInt32(ViSession vi, ViConstString channelName, ViAttr attribute, ViInt32 *value);
Private Declare PtrSafe Function niRFSG_GetAttributeViInt32 Lib "niRFSG_64" ( _
    ByVal vi As Long, ByVal channelName As String, ByVal attributeID As Long, ByRef value As Long) As Long

'ViStatus _VI_FUNC niRFSG_SetAttributeViInt32(ViSession vi, ViConstString channelName, ViAttr attribute, ViInt32 value);
Private Declare PtrSafe Function niRFSG_SetAttributeViInt32 Lib "niRFSG_64" ( _
    ByVal vi As Long, ByVal channelName As String, ByVal attributeID As Long, ByVal value As Long) As Long

'ViStatus _VI_FUNC niRFSG_GetAttributeViInt64(ViSession vi, ViConstString channelName, ViAttr attribute, ViInt64 *value);
Private Declare PtrSafe Function niRFSG_GetAttributeViInt64 Lib "niRFSG_64" ( _
    ByVal vi As Long, ByVal channelName As String, ByVal attributeID As Long, ByRef value As LongLong) As Long

'ViStatus _VI_FUNC niRFSG_SetAttributeViInt64(ViSession vi, ViConstString channelName, ViAttr attribute, ViInt64 value);
Private Declare PtrSafe Function niRFSG_SetAttributeViInt64 Lib "niRFSG_64" ( _
    ByVal vi As Long, ByVal channelName As String, ByVal attributeID As Long, ByVal value As LongLong) As Long

'ViStatus _VI_FUNC niRFSG_GetAttributeViReal64(ViSession vi, ViConstString channelName, ViAttr attribute, ViReal64 *value);
Private Declare PtrSafe Function niRFSG_GetAttributeViReal64 Lib "niRFSG_64" ( _
    ByVal vi As Long, ByVal channelName As String, ByVal attributeID As Long, ByRef value As Double) As Long

'ViStatus _VI_FUNC niRFSG_SetAttributeViReal64(ViSession vi, ViConstString channelName, ViAttr attribute, ViReal64 value);
Private Declare PtrSafe Function niRFSG_SetAttributeViReal64 Lib "niRFSG_64" ( _
    ByVal vi As Long, ByVal channelName As String, ByVal attributeID As Long, ByVal value As Double) As Long

'ViStatus _VI_FUNC niRFSG_GetAttributeViString(ViSession vi, ViConstString channelName, ViAttr attribute, ViInt32 bufSize, ViChar value[]);
Private Declare PtrSafe Function niRFSG_GetAttributeViString Lib "niRFSG_64" ( _
    ByVal vi As Long, ByVal channelName As String, ByVal attributeID As Long, ByVal bufSize As Long, ByVal value As LongPtr) As Long

'ViStatus _VI_FUNC niRFSG_SetAttributeViString(ViSession vi, ViConstString channelName, ViAttr attribute, ViConstString value);
Private Declare PtrSafe Function niRFSG_SetAttributeViString Lib "niRFSG_64" ( _
    ByVal vi As Long, ByVal channelName As String, ByVal attributeID As Long, ByVal value As String) As Long

'ViStatus _VI_FUNC niRFSG_GetAttributeViBoolean(ViSession vi, ViConstString channelName, ViAttr attribute, ViBoolean *value);
Private Declare PtrSafe Function niRFSG_GetAttributeViBoolean Lib "niRFSG_64" ( _
    ByVal vi As Long, ByVal channelName As String, ByVal attributeID As Long, ByRef value As Boolean) As Long

'ViStatus _VI_FUNC niRFSG_SetAttributeViBoolean(ViSession vi, ViConstString channelName, ViAttr attribute, ViBoolean value);
Private Declare PtrSafe Function niRFSG_SetAttributeViBoolean Lib "niRFSG_64" ( _
    ByVal vi As Long, ByVal channelName As String, ByVal attributeID As Long, ByVal value As Boolean) As Long

'ViStatus _VI_FUNC niRFSG_GetError(ViSession vi, ViStatus *errorCode, ViInt32 bufferSize, ViChar description[]);
Private Declare PtrSafe Function niRFSG_GetError Lib "niRFSG_64" ( _
    ByVal vi As Long, ByRef errorCode As Long, ByVal bufferSize As Long, ByVal errMessage As LongPtr) As Long

' ViStatus _VI_FUNC niRFSG_ConfigureRefClock(ViSession vi, ViConstString refClockSource, ViReal64 refClockRate);
Private Declare PtrSafe Function niRFSG_ConfigureRefClock Lib "niRFSG_64" ( _
    ByVal vi As Long, ByVal refClockSource As String, ByVal refClockRate As Double) As Long

'ViStatus _VI_FUNC niRFSG_ConfigureRF(ViSession vi, ViReal64 frequency, ViReal64 powerLevel);
Private Declare PtrSafe Function niRFSG_ConfigureRF Lib "niRFSG_64" ( _
    ByVal vi As Long, ByVal frequency As Double, ByVal powerLevel As Double) As Long

'ViStatus _VI_FUNC niRFSG_ConfigureGenerationMode(ViSession vi, ViInt32 generationMode);
Private Declare PtrSafe Function niRFSG_ConfigureGenerationMode Lib "niRFSG_64" ( _
    ByVal vi As Long, ByVal generationMode As Long) As Long

'ViStatus _VI_FUNC niRFSG_Commit(ViSession vi);
Private Declare PtrSafe Function niRFSG_Commit Lib "niRFSG_64" ( _
    ByVal vi As Long) As Long

'ViStatus _VI_FUNC niRFSG_Initiate(ViSession vi);
Private Declare PtrSafe Function niRFSG_Initiate Lib "niRFSG_64" ( _
    ByVal vi As Long) As Long

'ViStatus _VI_FUNC niRFSG_Abort(ViSession vi);
Private Declare PtrSafe Function niRFSG_Abort Lib "niRFSG_64" ( _
    ByVal vi As Long) As Long

'ViStatus _VI_FUNC niRFSG_CheckGenerationStatus(ViSession vi, ViBoolean* isDone);
Private Declare PtrSafe Function niRFSG_CheckGenerationStatus Lib "niRFSG_64" ( _
    ByVal vi As Long, ByRef isDone As Boolean) As Long

'ViStatus _VI_FUNC niRFSG_ConfigureOutputEnabled(ViSession vi, ViBoolean outputEnabled);
Private Declare PtrSafe Function niRFSG_ConfigureOutputEnabled Lib "niRFSG_64" ( _
    ByVal vi As Long, ByVal outputEnabled As Boolean) As Long

' Internal session
Private m_Session As Long
Private m_ResourceName As String
Private m_Playback As niRFSG_Playback

' Access to RFSG Playback Library object
Public Property Get Playback() As niRFSG_Playback
    Set Playback = m_Playback
End Property

' initialize internal variables, call Init first to create a valid session
Private Sub Class_Initialize()
    m_Session = 0
    m_ResourceName = ""
    Set m_Playback = Nothing
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
    
    size = niRFSG_GetError(m_Session, errorCode, 0, 0)
    ReDim buffer(size - 1) As Byte
 
    status = niRFSG_GetError(m_Session, errorCode, size, VarPtr(buffer(0)))
    errorMsg = StrConv(LeftB(buffer(), size - 1), vbUnicode) 'Remove \0 character and convert to Unicode
    
    niTools_RaiseError errorCode, errorMsg, "NI-RFSA"
End Sub

Public Sub InitSession(resourceName As String, IDQuery As Boolean, reset As Boolean, optionString As String)
    ' Make sure session is closed before opening
    CloseSession
    
    m_ResourceName = resourceName
    CheckError niRFSG_InitWithOptions(resourceName, IDQuery, reset, optionString, m_Session)
    
    Set m_Playback = New niRFSG_Playback
    m_Playback.InitSession m_Session
    
End Sub

Private Sub CloseSession()
    If m_Session = 0 Then Exit Sub
    
    CheckError niRFSG_close(m_Session)
    m_Session = 0
    m_ResourceName = ""
    Set m_Playback = Nothing
End Sub

Public Sub reset()
    CheckError niRFSG_reset(m_Session)
End Sub

Public Sub SelfCal()
    CheckError niRFSG_SelfCal(m_Session)
End Sub

Public Sub SetAttributeLong(channelName As String, attributeID As niRFSG_AttributeIDs, value As Long)
    CheckError niRFSG_SetAttributeViInt32(m_Session, channelName, attributeID, value)
End Sub

Public Sub GetAttributeLongLong(channelName As String, attributeID As niRFSG_AttributeIDs, ByRef value As LongLong)
    CheckError niRFSG_GetAttributeViInt64(m_Session, channelName, attributeID, value)
End Sub

Public Sub SetAttributeLongLong(channelName As String, attributeID As niRFSG_AttributeIDs, value As LongLong)
    CheckError niRFSG_SetAttributeViInt64(m_Session, channelName, attributeID, value)
End Sub

Public Sub GetAttributeDouble(channelName As String, attributeID As niRFSG_AttributeIDs, ByRef value As Double)
    CheckError niRFSG_GetAttributeViReal64(m_Session, channelName, attributeID, value)
End Sub

Public Sub SetAttributeDouble(channelName As String, attributeID As niRFSG_AttributeIDs, value As Double)
    CheckError niRFSG_SetAttributeViReal64(m_Session, channelName, attributeID, value)
End Sub

Public Sub GetAttributeString(channelName As String, attributeID As niRFSG_AttributeIDs, ByRef value As String)
    Dim size As Long
    Dim buffer() As Byte
    
    size = niRFSG_GetAttributeViString(m_Session, channelName, attributeID, 0, 0)
    ReDim buffer(size - 1) As Byte

    CheckError niRFSG_GetAttributeViString(m_Session, channelName, attributeID, size, VarPtr(buffer(0)))
    value = StrConv(LeftB(buffer(), size - 1), vbUnicode) ' Remove \0 character and convert to unicode
End Sub

Public Sub SetAttributeString(channelName As String, attributeID As niRFSG_AttributeIDs, value As String)
    CheckError niRFSG_SetAttributeViString(m_Session, channelName, attributeID, value)
End Sub

Public Sub GetAttributeBoolean(channelName As String, attributeID As niRFSG_AttributeIDs, ByRef value As Boolean)
    CheckError niRFSG_GetAttributeViBoolean(m_Session, channelName, attributeID, value)
End Sub

Public Sub SetAttributeBoolean(channelName As String, attributeID As niRFSG_AttributeIDs, value As Boolean)
    CheckError niRFSG_SetAttributeViBoolean(m_Session, channelName, attributeID, value)
End Sub

Public Sub ConfigureRefClock(refClockSource As String, refClockRate As Double)
    CheckError niRFSG_ConfigureRefClock(m_Session, refClockSource, refClockRate)
End Sub

Public Sub ConfigureRF(frequency As Double, powerLevel As Double)
    CheckError niRFSG_ConfigureRF(m_Session, frequency, powerLevel)
End Sub

Public Sub ConfigureGenerationMode(generationMode As niRFSG_GenerationMode)
    CheckError niRFSG_ConfigureGenerationMode(m_Session, generationMode)
End Sub

Public Sub Commit()
    CheckError niRFSG_Commit(m_Session)
End Sub

Public Sub Initiate()
    CheckError niRFSG_Initiate(m_Session)
End Sub

Public Sub Abort()
    CheckError niRFSG_Abort(m_Session)
End Sub

Public Sub CheckGenerationStatus(ByRef isDone As Boolean)
    CheckError niRFSG_CheckGenerationStatus(m_Session, isDone)
End Sub

Public Sub ConfigureOutputEnabled(outputEnabled As Boolean)
    CheckError niRFSG_ConfigureOutputEnabled(m_Session, outputEnabled)
End Sub
