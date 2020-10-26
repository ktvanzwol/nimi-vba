Attribute VB_Name = "niRFSA"
Option Explicit

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
    
    Dim session As niRFSA_Session
    
    Set session = New niRFSA_Session
    session.InitSession resourceName, IDQuery, reset, optionString
    
    Set niRFSA_CreateSession = session
End Function
