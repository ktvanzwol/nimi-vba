Attribute VB_Name = "niRFSA_Examples"
Option Explicit

Sub niRFSA_MeasureIQPower()
    Dim cRFSA As niRFSA_Session
    Dim sResourceName As String
    Dim dCenterFrequency As Double
    Dim dReferenceLevel As Double
    Dim dIQRate As Double
    Dim llNumSamples As LongLong
    Dim lNumSamples As Long
    Dim data() As NIComplexNumber
    Dim wfmInfo As niRFSA_wfmInfo
    Dim i As Long
    Dim magnitudeSquared As Double
    Dim accumulator As Double
    
    On Error GoTo Error
    
    sResourceName = "VST_5841_C1_S13"
    dCenterFrequency = 1000000000#
    dReferenceLevel = 0#
    dIQRate = 1000000#
    llNumSamples = 1000000
    
    Set cRFSA = niRFSA_CreateSession(sResourceName)
    With cRFSA
        .ConfigureAcquisitionType NIRFSA_VAL_IQ
        .ConfigureRefClock "OnboardClock", 10000000#
        .ConfigureIQCarrierFrequency "", dCenterFrequency
        .ConfigureReferenceLevel "", dReferenceLevel
        .ConfigureIQRate "", dIQRate
        .ConfigureNumberOfSamples "", True, llNumSamples
    End With
    
    ReDim data(CLng(llNumSamples) - 1) As NIComplexNumber
    cRFSA.ReadIQSingleRecordComplexF64 "", 10, data, wfmInfo
    
    ' Do something useful with the data
    ' We will present average power: 10log(((I^2 + Q ^2) / 2R) * 1000), where R = 50 Ohms
    
    lNumSamples = CLng(wfmInfo.actualSamples) 'Convert to Long required for indexing arrays
    If lNumSamples > 0 Then
        For i = 0 To lNumSamples - 1
            magnitudeSquared = data(i).real * data(i).real + data(i).imaginary * data(i).imaginary
            
            ' we need to handle this because log(0) return a range error.
            If magnitudeSquared = 0# Then
                magnitudeSquared = 0.00000001
            End If
            
            accumulator = accumulator + (10# * (Math.Log((magnitudeSquared / (2# * 50#)) * 1000#) / Math.Log(10#)))
        Next
        
        Debug.Print "Average Power = "; accumulator / lNumSamples; "dBm"
    End If
      
Error:
    If Err Then niTools_ErrorMsgBox Err
End Sub

Sub SelfCalibrate()
    Dim cRFSA As niRFSA_Session
    Dim sResourceName As String
    
    On Error GoTo Error
    
    sResourceName = "VST_5841_C1_S13"
    
    Set cRFSA = niRFSA_CreateSession(sResourceName)
    cRFSA.SelfCalibrate NIRFSA_VAL_SELF_CAL_OMIT_NONE
          
Error:
    If Err Then niTools_ErrorMsgBox Err
End Sub

Sub OptionStringAndAttributeTests()
    Dim cRFSA As niRFSA_Session
    Dim sResourceName As String
    Dim ports As String
    Dim selectedPort As String
    Dim referenceLevel As Double
    Dim acqType As niRFSA_AcquisitionType
    
    On Error GoTo Error
    
    sResourceName = "VST_5831_C1_S13"
    
    Set cRFSA = niRFSA_CreateSession(sResourceName, optionString:="Simulate=1,DriverSetup=Model:5831")
    With cRFSA
        .GetAttributeString "", NIRFSA_ATTR_AVAILABLE_PORTS, ports
        .SetAttributeString "", NIRFSA_ATTR_SELECTED_PORTS, "if1"
        .GetAttributeString "", NIRFSA_ATTR_SELECTED_PORTS, selectedPort
        .SetAttributeDouble "", NIRFSA_ATTR_REFERENCE_LEVEL, 1.234
        .GetAttributeDouble "", NIRFSA_ATTR_REFERENCE_LEVEL, referenceLevel
        .SetAttributeLong "", NIRFSA_ATTR_ACQUISITION_TYPE, NIRFSA_VAL_SPECTRUM
        .GetAttributeLong "", NIRFSA_ATTR_ACQUISITION_TYPE, acqType
    End With
    
    Debug.Print "Available Ports  = "; ports
    Debug.Print "Selected Port    = "; selectedPort
    Debug.Print "Reference Level  ="; referenceLevel
    Debug.Print "Acquisition Type = "; IIf(acqType = NIRFSA_VAL_IQ, "IQ", "Spectrum"); acqType
Error:
    If Err Then niTools_ErrorMsgBox Err
End Sub

