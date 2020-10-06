VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "niRFSA_Session"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Attribute IDs
Const NIRFSA_ATTR_ As Long = 0 ' Attribute constant defined privately, class property should be used to expose the attribute

'ViStatus _VI_FUNC niRFSA_init(ViRsrc resourceName, ViBoolean IDQuery, ViBoolean resetDevice, ViSession* vi);
Private Declare PtrSafe Function niRFSA_init Lib "C:\Program Files\IVI Foundation\IVI\Bin\niRFSA_64.dll" ( _
    ByVal resourceName As String, _
    ByVal IDQuery As Boolean, _
    ByVal resetDevice As Boolean, _
    ByRef vi As Long _
) As Long

'ViStatus _VI_FUNC niRFSA_close(ViSession vi);
Private Declare PtrSafe Function niRFSA_close Lib "C:\Program Files\IVI Foundation\IVI\Bin\niRFSA_64.dll" ( _
    ByVal vi As Long _
) As Long

'ViStatus _VI_FUNC niRFSA_reset(ViSession vi);
Private Declare PtrSafe Function niRFSA_reset Lib "C:\Program Files\IVI Foundation\IVI\Bin\niRFSA_64.dll" ( _
    ByVal vi As Long _
) As Long

'ViStatus _VI_FUNC niRFSA_GetAttributeViInt32(ViSession vi, ViConstString channelName, ViAttr attribute, ViInt32 *value);
Private Declare PtrSafe Function niRFSA_GetAttributeViInt32 Lib "C:\Program Files\IVI Foundation\IVI\Bin\niRFSA_64.dll" ( _
    ByVal vi As Long, _
    ByVal channelName As String, _
    ByVal attributeID As Long, _
    ByRef value As Long _
) As Long

'ViStatus _VI_FUNC niRFSA_SetAttributeViInt32(ViSession vi, ViConstString channelName, ViAttr attribute, ViInt32 value);
Private Declare PtrSafe Function niRFSA_SetAttributeViInt32 Lib "C:\Program Files\IVI Foundation\IVI\Bin\niRFSA_64.dll" ( _
    ByVal vi As Long, _
    ByVal channelName As String, _
    ByVal attributeID As Long, _
    ByVal value As Long _
) As Long

'ViStatus _VI_FUNC niRFSA_GetAttributeViInt64(ViSession vi, ViConstString channelName, ViAttr attribute, ViInt64 *value);
Private Declare PtrSafe Function niRFSA_GetAttributeViInt64 Lib "C:\Program Files\IVI Foundation\IVI\Bin\niRFSA_64.dll" ( _
    ByVal vi As Long, _
    ByVal channelName As String, _
    ByVal attributeID As Long, _
    ByRef value As LongLong _
) As Long

'ViStatus _VI_FUNC niRFSA_SetAttributeViInt64(ViSession vi, ViConstString channelName, ViAttr attribute, ViInt64 value);
Private Declare PtrSafe Function niRFSA_SetAttributeViInt64 Lib "C:\Program Files\IVI Foundation\IVI\Bin\niRFSA_64.dll" ( _
    ByVal vi As Long, _
    ByVal channelName As String, _
    ByVal attributeID As Long, _
    ByVal value As LongLong _
) As Long

'ViStatus _VI_FUNC niRFSA_GetAttributeViReal64(ViSession vi, ViConstString channelName, ViAttr attribute, ViReal64 *value);
Private Declare PtrSafe Function niRFSA_GetAttributeViReal64 Lib "C:\Program Files\IVI Foundation\IVI\Bin\niRFSA_64.dll" ( _
    ByVal vi As Long, _
    ByVal channelName As String, _
    ByVal attributeID As Long, _
    ByRef value As Double _
) As Long

'ViStatus _VI_FUNC niRFSA_SetAttributeViReal64(ViSession vi, ViConstString channelName, ViAttr attribute, ViReal64 value);
Private Declare PtrSafe Function niRFSA_SetAttributeViReal64 Lib "C:\Program Files\IVI Foundation\IVI\Bin\niRFSA_64.dll" ( _
    ByVal vi As Long, _
    ByVal channelName As String, _
    ByVal attributeID As Long, _
    ByVal value As Double _
) As Long

'ViStatus _VI_FUNC niRFSA_GetAttributeViString(ViSession vi, ViConstString channelName, ViAttr attribute, ViInt32 bufSize, ViChar value[]);
Private Declare PtrSafe Function niRFSA_GetAttributeViString Lib "C:\Program Files\IVI Foundation\IVI\Bin\niRFSA_64.dll" ( _
    ByVal vi As Long, _
    ByVal channelName As String, _
    ByVal attributeID As Long, _
    ByVal bufSize As Long, _
    ByVal value As LongPtr _
) As Long

'ViStatus _VI_FUNC niRFSA_SetAttributeViString(ViSession vi, ViConstString channelName, ViAttr attribute, ViConstString value);
Private Declare PtrSafe Function niRFSA_SetAttributeViString Lib "C:\Program Files\IVI Foundation\IVI\Bin\niRFSA_64.dll" ( _
    ByVal vi As Long, _
    ByVal channelName As String, _
    ByVal attributeID As Long, _
    ByVal value As String _
) As Long

'ViStatus _VI_FUNC niRFSA_GetAttributeViBoolean(ViSession vi, ViConstString channelName, ViAttr attribute, ViBoolean *value);
Private Declare PtrSafe Function niRFSA_GetAttributeViBoolean Lib "C:\Program Files\IVI Foundation\IVI\Bin\niRFSA_64.dll" ( _
    ByVal vi As Long, _
    ByVal channelName As String, _
    ByVal attributeID As Long, _
    ByRef value As Boolean _
) As Long

'ViStatus _VI_FUNC niRFSA_SetAttributeViBoolean(ViSession vi, ViConstString channelName, ViAttr attribute, ViBoolean value);
Private Declare PtrSafe Function niRFSA_SetAttributeViBoolean Lib "C:\Program Files\IVI Foundation\IVI\Bin\niRFSA_64.dll" ( _
    ByVal vi As Long, _
    ByVal channelName As String, _
    ByVal attributeID As Long, _
    ByVal value As Boolean _
) As Long

'ViStatus _VI_FUNC niRFSA_GetError(ViSession vi, ViStatus *errorCode, ViInt32 bufferSize, ViChar description[]);
Private Declare PtrSafe Function niRFSA_GetError Lib "C:\Program Files\IVI Foundation\IVI\Bin\niRFSA_64.dll" ( _
    ByVal vi As Long, _
    ByRef errorCode As Long, _
    ByVal bufferSize As Long, _
    ByVal errMessage As LongPtr _
) As Long

'ViStatus _VI_FUNC niRFSA_ConfigureRefClock(ViSession vi, ViConstString clockSource, ViReal64 refClockRate);
Private Declare PtrSafe Function niRFSA_ConfigureRefClock Lib "C:\Program Files\IVI Foundation\IVI\Bin\niRFSA_64.dll" ( _
    ByVal vi As Long, _
    ByVal clockSource As String, _
    ByVal refClockRate As Double _
) As Long

'ViStatus _VI_FUNC niRFSA_ConfigureReferenceLevel(ViSession vi, ViConstString channelList, ViReal64 referenceLevel);
Private Declare PtrSafe Function niRFSA_ConfigureReferenceLevel Lib "C:\Program Files\IVI Foundation\IVI\Bin\niRFSA_64.dll" ( _
    ByVal vi As Long, _
    ByVal channelList As String, _
    ByVal referenceLevel As Double _
) As Long

'ViStatus _VI_FUNC niRFSA_ConfigureAcquisitionType(ViSession vi, ViInt32 acquisitionType);
Private Declare PtrSafe Function niRFSA_ConfigureAcquisitionType Lib "C:\Program Files\IVI Foundation\IVI\Bin\niRFSA_64.dll" ( _
    ByVal vi As Long, _
    ByVal acquisitionType As Long _
) As Long

'ViStatus _VI_FUNC niRFSA_ConfigureIQCarrierFrequency(ViSession vi, ViConstString channelList, ViReal64 carrierFrequency);
Private Declare PtrSafe Function niRFSA_ConfigureIQCarrierFrequency Lib "C:\Program Files\IVI Foundation\IVI\Bin\niRFSA_64.dll" ( _
    ByVal vi As Long, _
    ByVal channelList As String, _
    ByVal carrierFrequency As Double _
) As Long

'ViStatus _VI_FUNC niRFSA_ConfigureNumberOfSamples(ViSession vi, ViConstString channelList, ViBoolean numberOfSamplesIsFinite, ViInt64 samplesPerRecord);
Private Declare PtrSafe Function niRFSA_ConfigureNumberOfSamples Lib "C:\Program Files\IVI Foundation\IVI\Bin\niRFSA_64.dll" ( _
    ByVal vi As Long, _
    ByVal channelList As String, _
    ByVal numberOfSamplesIsFinite As Boolean, _
    ByVal samplesPerRecord As LongLong _
) As Long

'ViStatus _VI_FUNC niRFSA_ConfigureIQRate(ViSession vi, ViConstString channelList, ViReal64 iqRate);
Private Declare PtrSafe Function niRFSA_ConfigureIQRate Lib "C:\Program Files\IVI Foundation\IVI\Bin\niRFSA_64.dll" ( _
    ByVal vi As Long, _
    ByVal channelList As String, _
    ByVal iqRate As LongLong _
) As Long

'ViStatus _VI_FUNC niRFSA_ReadIQSingleRecordComplexF64(ViSession vi, ViConstString channelList, ViReal64 timeout, NIComplexNumber* data, ViInt64 dataArraySize, niRFSA_wfmInfo* wfmInfo);
Private Declare PtrSafe Function niRFSA_ReadIQSingleRecordComplexF64 Lib "C:\Program Files\IVI Foundation\IVI\Bin\niRFSA_64.dll" ( _
    ByVal vi As Long, _
    ByVal channelList As String, _
    ByVal timeout As Double, _
    ByRef data As NIComplexNumber, _
    ByVal dataArraySize As Long, _
    ByRef wfmInfo As niRFSA_wfmInfo _
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
    
    size = niRFSA_GetError(m_Session, errorCode, 0, 0)
    buffer = String(size - 1, vbNullChar)
 
    status = niRFSA_GetError(m_Session, errorCode, size, StrPtr(buffer))
    niTools_CStrPtrToStr size, StrPtr(buffer), errorMsg
    
    niTools_RaiseError errorCode, errorMsg, "NI-RFSA"
End Sub

Public Sub InitSession(resourceName As String, IDQuery As Boolean, reset As Boolean)
    ' Make sure session is closed before opening
    CloseSession
    
    m_ResourceName = resourceName
    CheckError niRFSA_init(resourceName, IDQuery, reset, m_Session)
End Sub

Private Sub CloseSession()
    If m_Session = 0 Then Exit Sub
    
    CheckError niRFSA_close(m_Session)
    m_Session = 0
    m_ResourceName = ""
End Sub

Public Sub reset()
    CheckError niRFSA_reset(m_Session)
End Sub

Public Sub ConfigureRefClock(clockSource As String, refClockRate As Double)
    CheckError niRFSA_ConfigureRefClock(m_Session, clockSource, refClockRate)
End Sub

Public Sub ConfigureReferenceLevel(channelList As String, referenceLevel As Double)
    CheckError niRFSA_ConfigureReferenceLevel(m_Session, channelList, referenceLevel)
End Sub

Public Sub ConfigureAcquisitionType(acquisitionType As Long)
    CheckError niRFSA_ConfigureAcquisitionType(m_Session, acquisitionType)
End Sub

Public Sub ConfigureIQCarrierFrequency(channelList As String, carrierFrequency As Double)
    CheckError niRFSA_ConfigureIQCarrierFrequency(m_Session, channelList, carrierFrequency)
End Sub

Public Sub ConfigureNumberOfSamples(channelList As String, numberOfSamplesIsFinite As Boolean, samplesPerRecord As LongLong)
    CheckError niRFSA_ConfigureNumberOfSamples(m_Session, channelList, numberOfSamplesIsFinite, samplesPerRecord)
End Sub

Public Sub ConfigureIQRate(channelList As String, iqRate As LongLong)
    CheckError niRFSA_ConfigureIQRate(m_Session, channelList, iqRate)
End Sub

Public Sub ReadIQSingleRecordComplexF64(channelList As String, timeout As Double, ByRef data() As NIComplexNumber, ByRef wfmInfo As niRFSA_wfmInfo)
    CheckError niRFSA_ReadIQSingleRecordComplexF64(m_Session, channelList, timeout, data(LBound(data)), UBound(data), wfmInfo)
End Sub

