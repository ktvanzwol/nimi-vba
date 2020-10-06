VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "niModInst_Session"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'ViStatus niModInst_OpenInstalledDevicesSession (ViConstString driver, ViSession* handle, ViInt32* deviceCount);
Private Declare PtrSafe Function niModInst_OpenInstalledDevicesSession Lib "niModInst_64" ( _
    ByVal driver As String, _
    ByRef handle As Long, _
    ByRef deviceCount As Long _
) As Long

'ViStatus DllExport _VI_FUNC niModInst_GetInstalledDeviceAttributeViInt32(ViSession handle, ViInt32 index, ViInt32 attributeId, ViInt32* attributeValue);
Private Declare PtrSafe Function niModInst_GetInstalledDeviceAttributeViInt32 Lib "niModInst_64" ( _
    ByVal handle As Long, _
    ByVal index As Long, _
    ByVal attributeID As Long, _
    ByRef attributeValue As Long _
) As Long

'ViStatus niModInst_GetInstalledDeviceAttributeViString (ViSession handle, ViInt32 index, ViInt32 attributeID, ViInt32 attributeValueBufferSize, ViChar attributeValue[]);
Private Declare PtrSafe Function niModInst_GetInstalledDeviceAttributeViString Lib "niModInst_64" ( _
    ByVal handle As Long, _
    ByVal index As Long, _
    ByVal attributeID As Long, _
    ByVal attributeValueBufferSize As Long, _
    ByVal attributeValue As LongPtr _
) As Long

'ViStatus niModInst_CloseInstalledDevicesSession (ViSession handle);
Private Declare PtrSafe Function niModInst_CloseInstalledDevicesSession Lib "niModInst_64" ( _
    ByVal handle As Long _
) As Long

' ViStatus niModInst_GetExtendedErrorInfo (ViInt32 errorInfoBufferSize, ViChar errorInfo[]);
Private Declare PtrSafe Function niModInst_GetExtendedErrorInfo Lib "niModInst_64" ( _
    ByVal errorInfoBufferSize As Long, _
    ByVal errorInfo As LongPtr _
) As Long

Private m_Session As Long
Private m_Count As Long

' initialize internal variables, call Init first to create a valid session
Private Sub Class_Initialize()
    m_Session = 0
    m_Count = 0
End Sub

' Automatically clear session when object gets destroyed
Private Sub Class_Terminate()
    CloseSesion
End Sub

' Error Checker
Private Sub CheckError(status As Long)
    If status < 0 Then
        ErrorHandler status
    End If
End Sub

Private Sub ErrorHandler(errorCode As Long)
    Dim status As Long
    Dim size As Long
    Dim buffer As String
    Dim errorMsg As String
       
    size = niModInst_GetExtendedErrorInfo(0, 0) + 1
    buffer = String(size - 1, " ")
    
    status = niModInst_GetExtendedErrorInfo(size, StrPtr(buffer))
    niTools_CStrPtrToStr size, StrPtr(buffer), errorMsg
    
    niTools_RaiseError errorCode, errorMsg, "niTools"
End Sub

Public Sub InitSession(driver As String)
    CheckError niModInst_OpenInstalledDevicesSession(driver, m_Session, m_Count)
End Sub

Private Sub CloseSesion()
    If m_Session = 0 Then Exit Sub
    
    CheckError niModInst_CloseInstalledDevicesSession(m_Session)
    m_Session = 0
    m_Count = 0
End Sub

Public Property Get Count() As Long
    Count = m_Count
End Property

Public Sub GetInstalledDeviceAttributeString(ByVal index As Long, ByVal attributeID As Long, ByRef attributeValue As String)
    Dim size As Long
    Dim buffer As String
    
    size = niModInst_GetInstalledDeviceAttributeViString(m_Session, index, attributeID, 0, 0)
    buffer = String(size - 1, vbNullChar)
    
    CheckError niModInst_GetInstalledDeviceAttributeViString(m_Session, index, attributeID, size, StrPtr(buffer))
    niTools_CStrPtrToStr size, StrPtr(buffer), attributeValue
End Sub

Public Sub GetInstalledDeviceAttributeLong(ByVal index As Long, ByVal attributeID As Long, ByRef attributeValue As Long)
    CheckError niModInst_GetInstalledDeviceAttributeViInt32(m_Session, index, attributeID, attributeValue)
End Sub




