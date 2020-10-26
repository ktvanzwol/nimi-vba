Attribute VB_Name = "niRFSG_Examples"
Option Explicit

Sub Example_RFSG_SingleToneGeneration()
    Dim cRFSG As niRFSG_Session
    Dim sResourceName As String
    Dim dFrequency As Double
    Dim dPowerLevel As Double
    Dim sFinishTime As Single
    Dim isDone As Boolean
    
    On Error GoTo Error
    
    sResourceName = "VST_5841_C1_S13"
    dFrequency = 1000000000#
    dPowerLevel = -5#
    
    Set cRFSG = niRFSG_CreateSession(sResourceName) ' , optionString:="Simulate=1,DriverSetup=Model:5841")
    With cRFSG
        .ConfigureRefClock "OnboardClock", 10000000#
        .ConfigureRF dFrequency, dPowerLevel
        .ConfigureGenerationMode NIRFSG_VAL_CW
        .ConfigureOutputEnabled True
    End With
        
    ' Start Generation
    cRFSG.Initiate
    
    Debug.Print "Generating CW for 10 Seconds."
    sFinishTime = Timer + 10# ' Generate CW for ~10 seconds
    Do
        DoEvents
        
        ' Check status of generation, an error will be raised when a generation error occurs
        ' isDone is used to signal completion of a finite generation via script or waveform
        cRFSG.CheckGenerationStatus isDone
        
    Loop While (isDone = False) Or (Timer <= sFinishTime)

    Debug.Print "CW Generation done."

Error:
    ' Make sure the output is disabled even when a error occured
    cRFSG.ConfigureOutputEnabled False
    
    If Err Then niTools_ErrorMsgBox Err
End Sub

Sub Example_RFSG_GenerateWaveform()
    Dim cRFSG As niRFSG_Session
    Dim sResourceName As String
    Dim dCarrierFrequency As Double
    Dim dPowerLevel As Double
    Dim dExternalAttinuation As Double
    Dim sFilePath As String
    Dim sWfmName As String
    Dim sScript As String
    Dim sTimeOut As Single
    Dim isDone As Boolean
    
    On Error GoTo Error
    
    sResourceName = "VST_5841_C1_S13"
    dCarrierFrequency = 1000000000#
    dPowerLevel = -5#
    dExternalAttinuation = 1.5
    sFilePath = Environ("PUBLIC") & _
        "\Documents\National Instruments\NI-RFSG Playback Library\Examples\C\Support\LTE_FDD_PUSCH_10MHz_QPSK.tdms"
    sWfmName = "myWaveform"
    sScript = "script myScript " & _
                 "  repeat 5 " & _
                 "    generate myWaveform " & _
                 "  end repeat " & _
                 "end script"
    
    Set cRFSG = niRFSG_CreateSession(sResourceName) ', optionString:="Simulate=1,DriverSetup=Model:5841")
    With cRFSG
        .ConfigureRefClock "OnboardClock", 10000000#
        .ConfigureRF dCarrierFrequency, dPowerLevel
        .SetAttributeDouble "", NIRFSG_ATTR_EXTERNAL_GAIN, -1 * dExternalAttinuation
    End With
    
    cRFSG.Playback.ReadAndDownloadWaveformFromFile sFilePath, sWfmName
    cRFSG.Playback.SetScriptToGenerateSingleRFSG sScript
    
    ' Start Generation
    cRFSG.Initiate
    
    Debug.Print "Generating..."
    sTimeOut = Timer + 30# ' Generate with a 1min timeout
    Do
        DoEvents
        
        ' Check status of generation, an error will be raised when a generation error occurs
        ' isDone is used to signal completion of a finite generation via script or waveform
        cRFSG.CheckGenerationStatus isDone
        
    Loop While (isDone = False) Or (Timer <= sTimeOut)

    Debug.Print IIf(isDone, "Generation completed.", "Generation did not complete before Timeout.")

Error:
    cRFSG.Abort
    cRFSG.ConfigureOutputEnabled False
    cRFSG.Commit
    cRFSG.Playback.ClearAllWaveforms
    
    If Err Then niTools_ErrorMsgBox Err
End Sub

