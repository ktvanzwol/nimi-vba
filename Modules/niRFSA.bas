Attribute VB_Name = "niRFSA"
Option Explicit

'- NIRFSA_ATTR_ACQUISITION_TYPE Values -
Public Const NIRFSA_VAL_IQ As Long = 100
Public Const NIRFSA_VAL_SPECTRUM As Long = 101

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
