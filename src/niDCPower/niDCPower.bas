Attribute VB_Name = "niDCPower"
Option Explicit

' Values
' Public Const NIDCPOWER_VAL_ As Double = -1#

' niDCPower Factory Method
Public Function niDCPower_CreateSession(resourceName As String, channels As String, Optional Reset As Boolean = True, Optional optionString As String = "") As niDCPower_Session
    Dim session As niDCPower_Session
    
    Set session = New niDCPower_Session
    session.InitSession resourceName, channels, Reset, optionString
    
    Set niDCPower_CreateSession = session
End Function

