VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "niRFSG_Playback"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'ViStatus __stdcall niRFSGPlayback_GetError(int* errorCode, int errorDescriptionBufferSize, ViChar errorDescription[]);
Private Declare PtrSafe Function niRFSGPlayback_GetError Lib "niRFSGPlayback" ( _
    ByRef errorCode As Long, ByVal bufferSize As Long, ByVal errMessage As LongPtr) As Long
    
'ViStatus __stdcall niRFSGPlayback_ReadAndDownloadWaveformFromFile(ViSession rfsgSession, ViConstString filePath, ViConstString rfsgWaveformName);
Private Declare PtrSafe Function niRFSGPlayback_ReadAndDownloadWaveformFromFile Lib "niRFSGPlayback" ( _
    ByVal rfsgSession As Long, ByVal filePath As String, ByVal rfsgWaveformName As String) As Long

'ViStatus __stdcall niRFSGPlayback_SetScriptToGenerateSingleRFSG(ViSession rfsgSession, ViConstString  scriptText);
Private Declare PtrSafe Function niRFSGPlayback_SetScriptToGenerateSingleRFSG Lib "niRFSGPlayback" ( _
    ByVal rfsgSession As Long, ByVal scriptText As String) As Long

'ViStatus __stdcall niRFSGPlayback_ClearAllWaveforms(ViSession rfsgSession);
Private Declare PtrSafe Function niRFSGPlayback_ClearAllWaveforms Lib "niRFSGPlayback" ( _
    ByVal rfsgSession As Long) As Long


Private m_Session As Long

' initialize internal variables, call Init first to create a valid session
Private Sub Class_Initialize()
    m_Session = 0
End Sub

' Automatically clear session when object gets destroyed
Private Sub Class_Terminate()
    m_Session = 0
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
    
    size = niRFSGPlayback_GetError(errorCode, 0, 0)
    ReDim buffer(size - 1) As Byte
 
    status = niRFSGPlayback_GetError(errorCode, size, VarPtr(buffer(0)))
    errorMsg = StrConv(LeftB(buffer(), size - 1), vbUnicode) 'Remove \0 character and convert to Unicode
    
    niTools_RaiseError errorCode, errorMsg, "NI-RFSA"
End Sub

Public Sub InitPlayback(session As Long)
    m_Session = session
End Sub

Public Sub ReadAndDownloadWaveformFromFile(filePath As String, rfsgWaveformName As String)
    CheckError niRFSGPlayback_ReadAndDownloadWaveformFromFile(m_Session, filePath, rfsgWaveformName)
End Sub

Public Sub SetScriptToGenerateSingleRFSG(scriptText As String)
    CheckError niRFSGPlayback_SetScriptToGenerateSingleRFSG(m_Session, scriptText)
End Sub

Public Sub ClearAllWaveforms()
    CheckError niRFSGPlayback_ClearAllWaveforms(m_Session)
End Sub


