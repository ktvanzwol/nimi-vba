Attribute VB_Name = "niDMM"
Option Explicit

' Measurement Functions
Public Const NIDMM_VAL_DC_VOLTS As Long = 1                'DC Voltage                     All
Public Const NIDMM_VAL_AC_VOLTS As Long = 2                'AC Voltage with AC Coupling    All
Public Const NIDMM_VAL_DC_CURRENT As Long = 3              'DC Current                     All
Public Const NIDMM_VAL_AC_CURRENT As Long = 4              'AC Current                     All
Public Const NIDMM_VAL_2_WIRE_RES As Long = 5              '2-Wire Resistance              All
Public Const NIDMM_VAL_4_WIRE_RES As Long = 101            '4-Wire Resistance              NI 4060, NI 4065, NI 4070/4071/4072, NI 4080/4081/4082
Public Const NIDMM_VAL_FREQ As Long = 104                  'Frequency                      NI 4070/4071/4072 and NI 4080/4081/4082
Public Const NIDMM_VAL_PERIOD As Long = 105                'Period                         NI 4070/4071/4072 and NI 4080/4081/4082
Public Const NIDMM_VAL_TEMPERATURE As Long = 108           'Temperature
Public Const NIDMM_VAL_AC_VOLTS_DC_COUPLED As Long = 1001  'AC Voltage with DC Coupling    NI 4070/4071/4072 and NI 4080/4081/4082
Public Const NIDMM_VAL_DIODE As Long = 1002                'Diode                          All
Public Const NIDMM_VAL_WAVEFORM_VOLTAGE As Long = 1003     'Waveform Voltage               NI 4070/4071/4072 and NI 4080/4081/4082
Public Const NIDMM_VAL_WAVEFORM_CURRENT As Long = 1004     'Waveform Current               NI 4070/4071/4072 and NI 4080/4081/4082
Public Const NIDMM_VAL_CAPACITANCE As Long = 1005          'Capacitance                    NI 4072 and NI 4082
Public Const NIDMM_VAL_INDUCTANCE As Long = 1006           'Inductance                     NI 4072 and NI 4082

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

