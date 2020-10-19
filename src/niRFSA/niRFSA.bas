Attribute VB_Name = "niRFSA"
Option Explicit

'- NIRFSA_ATTR_ACQUISITION_TYPE Values -
Public Const NIRFSA_VAL_IQ As Long = 100
Public Const NIRFSA_VAL_SPECTRUM As Long = 101

' - Values for SELF CAL steps -
Public Const NIRFSA_VAL_SELF_CAL_OMIT_NONE As LongLong = &H0
Public Const NIRFSA_VAL_SELF_CAL_PRESELECTOR_ALIGNMENT As LongLong = &H1
Public Const NIRFSA_VAL_SELF_CAL_GAIN_REFERENCE As LongLong = &H2
Public Const NIRFSA_VAL_SELF_CAL_IF_FLATNESS As LongLong = &H4
Public Const NIRFSA_VAL_SELF_CAL_DIGITIZER_SELF_CAL As LongLong = &H8
Public Const NIRFSA_VAL_SELF_CAL_LO_SELF_CAL As LongLong = &H10
Public Const NIRFSA_VAL_SELF_CAL_AMPLITUDE_ACCURACY As LongLong = &H20
Public Const NIRFSA_VAL_SELF_CAL_RESIDUAL_LO_POWER As LongLong = &H40
Public Const NIRFSA_VAL_SELF_CAL_IMAGE_SUPPRESSION As LongLong = &H80
Public Const NIRFSA_VAL_SELF_CAL_SYNTHESIZER_ALIGNMENT As LongLong = &H100
Public Const NIRFSA_VAL_SELF_CAL_DC_OFFSET As LongLong = &H200


Type niRFSA_wfmInfo
   absoluteInitialX As Double
   relativeInitialX As Double
   xIncrement As Double
   actualSamples As LongLong
   offset As Double
   gain As Double
   reserved1 As Double
   reserved2 As Double
End Type

' niRFSA Factory Method
Public Function niRFSA_CreateSession( _
        resourceName As String, _
        Optional IDQuery As Boolean = True, _
        Optional reset As Boolean = True, _
        Optional optionString As String = "" _
    ) As niRFSA_Session
    
    Dim Session As niRFSA_Session
    
    Set Session = New niRFSA_Session
    Session.InitSession resourceName, IDQuery, reset, optionString
    
    Set niRFSA_CreateSession = Session
End Function
