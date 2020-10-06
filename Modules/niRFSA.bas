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
Public Function niRFSA_CreateSession(resourceName As String, Optional IDQuery As Boolean = True, Optional reset As Boolean = True) As niRFSA_Session
    Dim session As niRFSA_Session
    
    Set session = New niRFSA_Session
    session.InitSession resourceName, IDQuery, reset
    
    Set niRFSA_CreateSession = session
End Function


