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
Public Enum niRFSA_AttributeIDs
    NIRFSA_ATTR_ACQUISITION_TYPE = (1150000 + 1)
    NIRFSA_ATTR_IQ_CARRIER_FREQUENCY = (1150000 + 59)
    NIRFSA_ATTR_REFERENCE_LEVEL = (1150000 + 4)
    NIRFSA_ATTR_REFERENCE_LEVEL_HEADROOM = (1150000 + 309)
    NIRFSA_ATTR_ATTENUATION = (1150000 + 5)
    NIRFSA_ATTR_RF_ATTENUATION_STEP_SIZE = (1150000 + 155)
    NIRFSA_ATTR_DEVICE_CONFIGURATION_TEMPERATURE = (1150000 + 159)
    NIRFSA_ATTR_AMPLITUDE_SETTLING = (1150000 + 163)
    NIRFSA_ATTR_DIGITAL_GAIN = (1150000 + 301)
    NIRFSA_ATTR_SELECTED_PORTS = (1150000 + 297)
    NIRFSA_ATTR_AVAILABLE_PORTS = (1150000 + 306)
End Enum

'- NIRFSA_ATTR_ACQUISITION_TYPE Values -
Public Enum niRFSA_AcquisitionType
    NIRFSA_VAL_IQ = 100
    NIRFSA_VAL_SPECTRUM = 101
End Enum


' - Values for SELF CAL steps -
Public Enum niRFSA_SelfCalSteps
    NIRFSA_VAL_SELF_CAL_OMIT_NONE = &H0
    NIRFSA_VAL_SELF_CAL_PRESELECTOR_ALIGNMENT = &H1
    NIRFSA_VAL_SELF_CAL_GAIN_REFERENCE = &H2
    NIRFSA_VAL_SELF_CAL_IF_FLATNESS = &H4
    NIRFSA_VAL_SELF_CAL_DIGITIZER_SELF_CAL = &H8
    NIRFSA_VAL_SELF_CAL_LO_SELF_CAL = &H10
    NIRFSA_VAL_SELF_CAL_AMPLITUDE_ACCURACY = &H20
    NIRFSA_VAL_SELF_CAL_RESIDUAL_LO_POWER = &H40
    NIRFSA_VAL_SELF_CAL_IMAGE_SUPPRESSION = &H80
    NIRFSA_VAL_SELF_CAL_SYNTHESIZER_ALIGNMENT = &H100
    NIRFSA_VAL_SELF_CAL_DC_OFFSET = &H200
End Enum

'ViStatus _VI_FUNC niRFSA_init(ViRsrc resourceName, ViBoolean IDQuery, ViBoolean resetDevice, ViSession* vi);
Private Declare PtrSafe Function niRFSA_init Lib "niRFSA_64" ( _
    ByVal resourceName As String, _
    ByVal IDQuery As Boolean, _
    ByVal resetDevice As Boolean, _
    ByRef vi As Long _
) As Long

'ViStatus _VI_FUNC niRFSA_InitWithOptions(ViRsrc resourceName, ViBoolean IDQuery, ViBoolean reset, ViConstString optionString, ViSession* newVi);
Private Declare PtrSafe Function niRFSA_InitWithOptions Lib "niRFSA_64" ( _
    ByVal resourceName As String, _
    ByVal IDQuery As Boolean, _
    ByVal resetDevice As Boolean, _
    ByVal optionString As String, _
    ByRef vi As Long _
) As Long

'ViStatus _VI_FUNC niRFSA_close(ViSession vi);
Private Declare PtrSafe Function niRFSA_close Lib "niRFSA_64" ( _
    ByVal vi As Long _
) As Long

'ViStatus _VI_FUNC niRFSA_reset(ViSession vi);
Private Declare PtrSafe Function niRFSA_reset Lib "niRFSA_64" ( _
    ByVal vi As Long _
) As Long

'ViStatus _VI_FUNC niRFSA_SelfCalibrate(ViSession vi, ViInt64 stepsToOmit);
Private Declare PtrSafe Function niRFSA_SelfCalibrate Lib "niRFSA_64" ( _
    ByVal vi As Long, _
    ByVal stepsToOmit As LongLong _
) As Long

'ViStatus _VI_FUNC niRFSA_GetAttributeViInt32(ViSession vi, ViConstString channelName, ViAttr attribute, ViInt32 *value);
Private Declare PtrSafe Function niRFSA_GetAttributeViInt32 Lib "niRFSA_64" ( _
    ByVal vi As Long, _
    ByVal channelName As String, _
    ByVal attributeID As Long, _
    ByRef value As Long _
) As Long

'ViStatus _VI_FUNC niRFSA_SetAttributeViInt32(ViSession vi, ViConstString channelName, ViAttr attribute, ViInt32 value);
Private Declare PtrSafe Function niRFSA_SetAttributeViInt32 Lib "niRFSA_64" ( _
    ByVal vi As Long, _
    ByVal channelName As String, _
    ByVal attributeID As Long, _
    ByVal value As Long _
) As Long

'ViStatus _VI_FUNC niRFSA_GetAttributeViInt64(ViSession vi, ViConstString channelName, ViAttr attribute, ViInt64 *value);
Private Declare PtrSafe Function niRFSA_GetAttributeViInt64 Lib "niRFSA_64" ( _
    ByVal vi As Long, _
    ByVal channelName As String, _
    ByVal attributeID As Long, _
    ByRef value As LongLong _
) As Long

'ViStatus _VI_FUNC niRFSA_SetAttributeViInt64(ViSession vi, ViConstString channelName, ViAttr attribute, ViInt64 value);
Private Declare PtrSafe Function niRFSA_SetAttributeViInt64 Lib "niRFSA_64" ( _
    ByVal vi As Long, _
    ByVal channelName As String, _
    ByVal attributeID As Long, _
    ByVal value As LongLong _
) As Long

'ViStatus _VI_FUNC niRFSA_GetAttributeViReal64(ViSession vi, ViConstString channelName, ViAttr attribute, ViReal64 *value);
Private Declare PtrSafe Function niRFSA_GetAttributeViReal64 Lib "niRFSA_64" ( _
    ByVal vi As Long, _
    ByVal channelName As String, _
    ByVal attributeID As Long, _
    ByRef value As Double _
) As Long

'ViStatus _VI_FUNC niRFSA_SetAttributeViReal64(ViSession vi, ViConstString channelName, ViAttr attribute, ViReal64 value);
Private Declare PtrSafe Function niRFSA_SetAttributeViReal64 Lib "niRFSA_64" ( _
    ByVal vi As Long, _
    ByVal channelName As String, _
    ByVal attributeID As Long, _
    ByVal value As Double _
) As Long

'ViStatus _VI_FUNC niRFSA_GetAttributeViString(ViSession vi, ViConstString channelName, ViAttr attribute, ViInt32 bufSize, ViChar value[]);
Private Declare PtrSafe Function niRFSA_GetAttributeViString Lib "niRFSA_64" ( _
    ByVal vi As Long, _
    ByVal channelName As String, _
    ByVal attributeID As Long, _
    ByVal bufSize As Long, _
    ByVal value As LongPtr _
) As Long

'ViStatus _VI_FUNC niRFSA_SetAttributeViString(ViSession vi, ViConstString channelName, ViAttr attribute, ViConstString value);
Private Declare PtrSafe Function niRFSA_SetAttributeViString Lib "niRFSA_64" ( _
    ByVal vi As Long, _
    ByVal channelName As String, _
    ByVal attributeID As Long, _
    ByVal value As String _
) As Long

'ViStatus _VI_FUNC niRFSA_GetAttributeViBoolean(ViSession vi, ViConstString channelName, ViAttr attribute, ViBoolean *value);
Private Declare PtrSafe Function niRFSA_GetAttributeViBoolean Lib "niRFSA_64" ( _
    ByVal vi As Long, _
    ByVal channelName As String, _
    ByVal attributeID As Long, _
    ByRef value As Boolean _
) As Long

'ViStatus _VI_FUNC niRFSA_SetAttributeViBoolean(ViSession vi, ViConstString channelName, ViAttr attribute, ViBoolean value);
Private Declare PtrSafe Function niRFSA_SetAttributeViBoolean Lib "niRFSA_64" ( _
    ByVal vi As Long, _
    ByVal channelName As String, _
    ByVal attributeID As Long, _
    ByVal value As Boolean _
) As Long

'ViStatus _VI_FUNC niRFSA_GetError(ViSession vi, ViStatus *errorCode, ViInt32 bufferSize, ViChar description[]);
Private Declare PtrSafe Function niRFSA_GetError Lib "niRFSA_64" ( _
    ByVal vi As Long, _
    ByRef errorCode As Long, _
    ByVal bufferSize As Long, _
    ByVal errMessage As LongPtr _
) As Long

'ViStatus _VI_FUNC niRFSA_ConfigureRefClock(ViSession vi, ViConstString clockSource, ViReal64 refClockRate);
Private Declare PtrSafe Function niRFSA_ConfigureRefClock Lib "niRFSA_64" ( _
    ByVal vi As Long, _
    ByVal clockSource As String, _
    ByVal refClockRate As Double _
) As Long

'ViStatus _VI_FUNC niRFSA_ConfigureReferenceLevel(ViSession vi, ViConstString channelList, ViReal64 referenceLevel);
Private Declare PtrSafe Function niRFSA_ConfigureReferenceLevel Lib "niRFSA_64" ( _
    ByVal vi As Long, _
    ByVal channelList As String, _
    ByVal referenceLevel As Double _
) As Long

'ViStatus _VI_FUNC niRFSA_ConfigureAcquisitionType(ViSession vi, ViInt32 acquisitionType);
Private Declare PtrSafe Function niRFSA_ConfigureAcquisitionType Lib "niRFSA_64" ( _
    ByVal vi As Long, _
    ByVal AcquisitionType As Long _
) As Long

'ViStatus _VI_FUNC niRFSA_ConfigureIQCarrierFrequency(ViSession vi, ViConstString channelList, ViReal64 carrierFrequency);
Private Declare PtrSafe Function niRFSA_ConfigureIQCarrierFrequency Lib "niRFSA_64" ( _
    ByVal vi As Long, _
    ByVal channelList As String, _
    ByVal CarrierFrequency As Double _
) As Long

'ViStatus _VI_FUNC niRFSA_ConfigureNumberOfSamples(ViSession vi, ViConstString channelList, ViBoolean numberOfSamplesIsFinite, ViInt64 samplesPerRecord);
Private Declare PtrSafe Function niRFSA_ConfigureNumberOfSamples Lib "niRFSA_64" ( _
    ByVal vi As Long, _
    ByVal channelList As String, _
    ByVal numberOfSamplesIsFinite As Boolean, _
    ByVal samplesPerRecord As LongLong _
) As Long

'ViStatus _VI_FUNC niRFSA_ConfigureIQRate(ViSession vi, ViConstString channelList, ViReal64 iqRate);
Private Declare PtrSafe Function niRFSA_ConfigureIQRate Lib "niRFSA_64" ( _
    ByVal vi As Long, _
    ByVal channelList As String, _
    ByVal iqRate As Double _
) As Long

'ViStatus _VI_FUNC niRFSA_ReadIQSingleRecordComplexF64(ViSession vi, ViConstString channelList, ViReal64 timeout, NIComplexNumber* data, ViInt64 dataArraySize, niRFSA_wfmInfo* wfmInfo);
Private Declare PtrSafe Function niRFSA_ReadIQSingleRecordComplexF64 Lib "niRFSA_64" ( _
    ByVal vi As Long, _
    ByVal channelList As String, _
    ByVal timeout As Double, _
    ByVal data As LongPtr, _
    ByVal dataArraySize As Long, _
    ByVal wfmInfo As LongPtr _
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
    
    size = niRFSA_GetError(m_Session, errorCode, 0, 0)
    ReDim buffer(size - 1) As Byte
 
    status = niRFSA_GetError(m_Session, errorCode, size, VarPtr(buffer(0)))
    errorMsg = StrConv(LeftB(buffer(), size - 1), vbUnicode) 'Remove \0 character and convert to Unicode
    
    niTools_RaiseError errorCode, errorMsg, "NI-RFSA"
End Sub

Public Sub InitSession(resourceName As String, IDQuery As Boolean, reset As Boolean, optionString As String)
    ' Make sure session is closed before opening
    CloseSession
    
    m_ResourceName = resourceName
    CheckError niRFSA_InitWithOptions(resourceName, IDQuery, reset, optionString, m_Session)
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

Public Sub SelfCalibrate(Optional stepsToOmit As LongLong = NIRFSA_VAL_SELF_CAL_OMIT_NONE)
    CheckError niRFSA_SelfCalibrate(m_Session, stepsToOmit)
End Sub

Public Sub ConfigureRefClock(clockSource As String, refClockRate As Double)
    CheckError niRFSA_ConfigureRefClock(m_Session, clockSource, refClockRate)
End Sub

Public Sub ConfigureReferenceLevel(channelList As String, referenceLevel As Double)
    CheckError niRFSA_ConfigureReferenceLevel(m_Session, channelList, referenceLevel)
End Sub

Public Sub ConfigureAcquisitionType(AcquisitionType As niRFSA_AcquisitionType)
    CheckError niRFSA_ConfigureAcquisitionType(m_Session, AcquisitionType)
End Sub

Public Sub ConfigureIQCarrierFrequency(channelList As String, CarrierFrequency As Double)
    CheckError niRFSA_ConfigureIQCarrierFrequency(m_Session, channelList, CarrierFrequency)
End Sub

Public Sub ConfigureNumberOfSamples(channelList As String, numberOfSamplesIsFinite As Boolean, samplesPerRecord As LongLong)
    CheckError niRFSA_ConfigureNumberOfSamples(m_Session, channelList, numberOfSamplesIsFinite, samplesPerRecord)
End Sub

Public Sub ConfigureIQRate(channelList As String, iqRate As Double)
    CheckError niRFSA_ConfigureIQRate(m_Session, channelList, iqRate)
End Sub

Public Sub ReadIQSingleRecordComplexF64(channelList As String, timeout As Double, ByRef data() As NIComplexNumber, ByRef wfmInfo As niRFSA_wfmInfo)
    Dim length As Long
    Dim lb As Long
    
    lb = LBound(data)
    length = (UBound(data) - lb + 1)
    CheckError niRFSA_ReadIQSingleRecordComplexF64(m_Session, channelList, timeout, VarPtr(data(lb)), length, VarPtr(wfmInfo))
End Sub

Public Sub GetAttributeLong(channelName As String, attributeID As niRFSA_AttributeIDs, ByRef value As Long)
    CheckError niRFSA_GetAttributeViInt32(m_Session, channelName, attributeID, value)
End Sub

Public Sub SetAttributeLong(channelName As String, attributeID As niRFSA_AttributeIDs, value As Long)
    CheckError niRFSA_SetAttributeViInt32(m_Session, channelName, attributeID, value)
End Sub

Public Sub GetAttributeLongLong(channelName As String, attributeID As niRFSA_AttributeIDs, ByRef value As LongLong)
    CheckError niRFSA_GetAttributeViInt64(m_Session, channelName, attributeID, value)
End Sub

Public Sub SetAttributeLongLong(channelName As String, attributeID As niRFSA_AttributeIDs, value As LongLong)
    CheckError niRFSA_SetAttributeViInt64(m_Session, channelName, attributeID, value)
End Sub

Public Sub GetAttributeDouble(channelName As String, attributeID As niRFSA_AttributeIDs, ByRef value As Double)
    CheckError niRFSA_GetAttributeViReal64(m_Session, channelName, attributeID, value)
End Sub

Public Sub SetAttributeDouble(channelName As String, attributeID As niRFSA_AttributeIDs, value As Double)
    CheckError niRFSA_SetAttributeViReal64(m_Session, channelName, attributeID, value)
End Sub

Public Sub GetAttributeString(channelName As String, attributeID As niRFSA_AttributeIDs, ByRef value As String)
    Dim size As Long
    Dim buffer() As Byte
    
    size = niRFSA_GetAttributeViString(m_Session, channelName, attributeID, 0, 0)
    ReDim buffer(size - 1) As Byte

    CheckError niRFSA_GetAttributeViString(m_Session, channelName, attributeID, size, VarPtr(buffer(0)))
    value = StrConv(LeftB(buffer(), size - 1), vbUnicode) ' Remove \0 character and convert to unicode
End Sub

Public Sub SetAttributeString(channelName As String, attributeID As niRFSA_AttributeIDs, value As String)
    CheckError niRFSA_SetAttributeViString(m_Session, channelName, attributeID, value)
End Sub

Public Sub GetAttributeBoolean(channelName As String, attributeID As niRFSA_AttributeIDs, ByRef value As Boolean)
    CheckError niRFSA_GetAttributeViBoolean(m_Session, channelName, attributeID, value)
End Sub

Public Sub SetAttributeBoolean(channelName As String, attributeID As niRFSA_AttributeIDs, value As Boolean)
    CheckError niRFSA_SetAttributeViBoolean(m_Session, channelName, attributeID, value)
End Sub
