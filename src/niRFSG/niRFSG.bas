Attribute VB_Name = "niRFSG"
Option Explicit

' niRFSG Factory Method
Public Function niRFSG_CreateSession( _
        resourceName As String, _
        Optional IDQuery As Boolean = True, _
        Optional Reset As Boolean = True, _
        Optional optionString As String = "" _
    ) As niRFSG_Session
    
    Dim session As niRFSG_Session
    
    Set session = New niRFSG_Session
    session.InitSession resourceName, IDQuery, Reset, optionString
    
    Set niRFSG_CreateSession = session
End Function
