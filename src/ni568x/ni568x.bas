Attribute VB_Name = "ni568x"
Option Explicit

' Measurement Units
Public Const NI568X_VAL_DBM As Long = 1         'Sets the units to dBm.
Public Const NI568X_VAL_WATTS As Long = 4       'Sets the units to watts.
Public Const NI568X_VAL_MWATTS As Long = 1001   'Sets the units to milliwatts.
Public Const NI568X_VAL_UWATTS As Long = 1002   'Sets the units to microwatts.

' Time Limit Constants
Public Const NI568X_VAL_MAX_TIME_IMMEDIATE As Long = 0 'Immediate timeout.
Public Const NI568X_VAL_MAX_TIME_INFINITE As Long = -1 'Infinite timeout.

' ni568x Factory Method
Public Function ni568x_CreateSession(resourceName As String, Optional IDQuery As Boolean = True, Optional Reset As Boolean = True) As ni568x_Session
    Dim session As ni568x_Session
    
    Set session = New ni568x_Session
    session.InitSession resourceName, IDQuery, Reset
    
    Set ni568x_CreateSession = session
End Function


