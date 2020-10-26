Attribute VB_Name = "niDMM"
Option Explicit

' Auto Range Values
Public Const NIDMM_VAL_AUTO_RANGE_ON As Double = -1#
Public Const NIDMM_VAL_AUTO_RANGE_OFF As Double = -2#
Public Const NIDMM_VAL_AUTO_RANGE_ONCE As Double = -3#

' Powerline Frequencies
Public Const NIDMM_VAL_50_HERTZ As Double = 50#
Public Const NIDMM_VAL_60_HERTZ As Double = 60#

' Auto Time Limit
Public Const NIDMM_VAL_TIME_LIMIT_AUTO As Double = -1

' niDMM Factory Method
Public Function niDMM_CreateSession(resourceName As String, Optional IDQuery As Boolean = True, Optional reset As Boolean = True) As niDMM_Session
    Dim session As niDMM_Session
    
    Set session = New niDMM_Session
    session.InitSession resourceName, IDQuery, reset
    
    Set niDMM_CreateSession = session
End Function

