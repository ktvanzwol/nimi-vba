Attribute VB_Name = "niDRIVER"
Option Explicit

' Public Driver Constants
Public Const NIDRIVER_VAL_X As Long = 1         'Public Driver constants defined here

' ni<DRIVER> Factory Method
Public Function ni<DRIVER>_CreateSession(resourceName As String, Optional IDQuery As Boolean = True, Optional reset As Boolean = True) As ni<DRIVER>_Session
    Dim session As ni<DRIVER>_Session
    
    Set session = New ni<DRIVER>_Session
    session.InitSession resourceName, IDQuery, reset
    
    Set ni<DRIVER>_CreateSession = session
End Function


