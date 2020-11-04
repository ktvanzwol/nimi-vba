Attribute VB_Name = "niTools"
Option Explicit

' IVI Definition
Public Const IVI_ATTR_BASE = 1000000
Public Const IVI_ENGINE_PRIVATE_ATTR_BASE = (IVI_ATTR_BASE + 0)         '/* base for private attributes of the IVI engine */
Public Const IVI_ENGINE_PUBLIC_ATTR_BASE = (IVI_ATTR_BASE + 50000)      '/* base for public attributes of the IVI engine */
Public Const IVI_SPECIFIC_PUBLIC_ATTR_BASE = (IVI_ATTR_BASE + 150000)   '/* base for public attributes of specific drivers */
Public Const IVI_SPECIFIC_PRIVATE_ATTR_BASE = (IVI_ATTR_BASE + 200000)  '/* base for private attributes of specific drivers */
                                                                        '/* This value was changed from IVI_ATTR_BASE + 100000 in the version of this file released in August 2013 (ICP 4.6). */
                                                                        '/* A private attribute, by its very definition, should not be passed to another module; it should stay private to the compiled module. */
Public Const IVI_CLASS_PUBLIC_ATTR_BASE = (IVI_ATTR_BASE + 250000)      '/* base for public attributes of class drivers */
Public Const IVI_CLASS_PRIVATE_ATTR_BASE = (IVI_ATTR_BASE + 300000)     '/* base for private attributes of class drivers */
                                                                        '/* This value was changed from IVI_ATTR_BASE + 200000 in the version of this file released in August 2013 (ICP 4.6). */
                                                                        '/* A private attribute, by its very definition, should not be passed to another module; it should stay private to the compiled module. */

Public Enum Ivi_AttributeIDs
    IVI_ATTR_NONE = &HFFFFFFFF
    IVI_ATTR_RANGE_CHECK = (IVI_ENGINE_PUBLIC_ATTR_BASE + 2)
    IVI_ATTR_QUERY_INSTRUMENT_STATUS = (IVI_ENGINE_PUBLIC_ATTR_BASE + 3)
    IVI_ATTR_CACHE = (IVI_ENGINE_PUBLIC_ATTR_BASE + 4)
    IVI_ATTR_SIMULATE = (IVI_ENGINE_PUBLIC_ATTR_BASE + 5)
    IVI_ATTR_RECORD_COERCIONS = (IVI_ENGINE_PUBLIC_ATTR_BASE + 6)
    IVI_ATTR_DRIVER_SETUP = (IVI_ENGINE_PUBLIC_ATTR_BASE + 7)
    IVI_ATTR_INTERCHANGE_CHECK = (IVI_ENGINE_PUBLIC_ATTR_BASE + 21)
    IVI_ATTR_SPY = (IVI_ENGINE_PUBLIC_ATTR_BASE + 22)
    IVI_ATTR_USE_SPECIFIC_SIMULATION = (IVI_ENGINE_PUBLIC_ATTR_BASE + 23)
    IVI_ATTR_PRIMARY_ERROR = (IVI_ENGINE_PUBLIC_ATTR_BASE + 101)
    IVI_ATTR_SECONDARY_ERROR = (IVI_ENGINE_PUBLIC_ATTR_BASE + 102)
    IVI_ATTR_ERROR_ELABORATION = (IVI_ENGINE_PUBLIC_ATTR_BASE + 103)
    IVI_ATTR_CHANNEL_COUNT = (IVI_ENGINE_PUBLIC_ATTR_BASE + 203)
    IVI_ATTR_CLASS_DRIVER_PREFIX = (IVI_ENGINE_PUBLIC_ATTR_BASE + 301)
    IVI_ATTR_SPECIFIC_DRIVER_PREFIX = (IVI_ENGINE_PUBLIC_ATTR_BASE + 302)
    IVI_ATTR_SPECIFIC_DRIVER_LOCATOR = (IVI_ENGINE_PUBLIC_ATTR_BASE + 303)
    IVI_ATTR_IO_RESOURCE_DESCRIPTOR = (IVI_ENGINE_PUBLIC_ATTR_BASE + 304)
    IVI_ATTR_LOGICAL_NAME = (IVI_ENGINE_PUBLIC_ATTR_BASE + 305)
    IVI_ATTR_VISA_RM_SESSION = (IVI_ENGINE_PUBLIC_ATTR_BASE + 321)
    IVI_ATTR_SYSTEM_IO_SESSION = (IVI_ENGINE_PUBLIC_ATTR_BASE + 322)
    IVI_ATTR_IO_SESSION_TYPE = (IVI_ENGINE_PUBLIC_ATTR_BASE + 324)
    IVI_ATTR_SYSTEM_IO_TIMEOUT = (IVI_ENGINE_PUBLIC_ATTR_BASE + 325)
    IVI_ATTR_SUPPORTED_INSTRUMENT_MODELS = (IVI_ENGINE_PUBLIC_ATTR_BASE + 327)
    IVI_ATTR_GROUP_CAPABILITIES = (IVI_ENGINE_PUBLIC_ATTR_BASE + 401)
    IVI_ATTR_FUNCTION_CAPABILITIES = (IVI_ENGINE_PUBLIC_ATTR_BASE + 402)
    IVI_ATTR_ENGINE_MAJOR_VERSION = (IVI_ENGINE_PUBLIC_ATTR_BASE + 501)
    IVI_ATTR_ENGINE_MINOR_VERSION = (IVI_ENGINE_PUBLIC_ATTR_BASE + 502)
    IVI_ATTR_SPECIFIC_DRIVER_MAJOR_VERSION = (IVI_ENGINE_PUBLIC_ATTR_BASE + 503)
    IVI_ATTR_SPECIFIC_DRIVER_MINOR_VERSION = (IVI_ENGINE_PUBLIC_ATTR_BASE + 504)
    IVI_ATTR_CLASS_DRIVER_MAJOR_VERSION = (IVI_ENGINE_PUBLIC_ATTR_BASE + 505)
    IVI_ATTR_CLASS_DRIVER_MINOR_VERSION = (IVI_ENGINE_PUBLIC_ATTR_BASE + 506)
    IVI_ATTR_INSTRUMENT_FIRMWARE_REVISION = (IVI_ENGINE_PUBLIC_ATTR_BASE + 510)
    IVI_ATTR_INSTRUMENT_MANUFACTURER = (IVI_ENGINE_PUBLIC_ATTR_BASE + 511)
    IVI_ATTR_INSTRUMENT_MODEL = (IVI_ENGINE_PUBLIC_ATTR_BASE + 512)
    IVI_ATTR_SPECIFIC_DRIVER_VENDOR = (IVI_ENGINE_PUBLIC_ATTR_BASE + 513)
    IVI_ATTR_SPECIFIC_DRIVER_DESCRIPTION = (IVI_ENGINE_PUBLIC_ATTR_BASE + 514)
    IVI_ATTR_SPECIFIC_DRIVER_CLASS_SPEC_MAJOR_VERSION = (IVI_ENGINE_PUBLIC_ATTR_BASE + 515)
    IVI_ATTR_SPECIFIC_DRIVER_CLASS_SPEC_MINOR_VERSION = (IVI_ENGINE_PUBLIC_ATTR_BASE + 516)
    IVI_ATTR_CLASS_DRIVER_VENDOR = (IVI_ENGINE_PUBLIC_ATTR_BASE + 517)
    IVI_ATTR_CLASS_DRIVER_DESCRIPTION = (IVI_ENGINE_PUBLIC_ATTR_BASE + 518)
    IVI_ATTR_CLASS_DRIVER_CLASS_SPEC_MAJOR_VERSION = (IVI_ENGINE_PUBLIC_ATTR_BASE + 519)
    IVI_ATTR_CLASS_DRIVER_CLASS_SPEC_MINOR_VERSION = (IVI_ENGINE_PUBLIC_ATTR_BASE + 520)
    IVI_ATTR_SPECIFIC_DRIVER_REVISION = (IVI_ENGINE_PUBLIC_ATTR_BASE + 551)
    IVI_ATTR_CLASS_DRIVER_REVISION = (IVI_ENGINE_PUBLIC_ATTR_BASE + 552)
    IVI_ATTR_ENGINE_REVISION = (IVI_ENGINE_PUBLIC_ATTR_BASE + 553)
    IVI_ATTR_OPC_CALLBACK = (IVI_ENGINE_PUBLIC_ATTR_BASE + 602)
    IVI_ATTR_CHECK_STATUS_CALLBACK = (IVI_ENGINE_PUBLIC_ATTR_BASE + 603)
    IVI_ATTR_USER_INTERCHANGE_CHECK_CALLBACK = (IVI_ENGINE_PUBLIC_ATTR_BASE + 801)
End Enum

Type NIComplexDouble
    real As Double
    imaginary As Double
End Type

Type NIComplexSingle
    real As Single
    imaginary As Single
End Type

' VBA Error code used for NI Driver errors
Public Const niErrorNumber As Long = vbObjectError + 1024

' Utility function to Raise a NI Driver Error, works in combination with niTools_ErrorMsgBox and On Error Goto <label>
Public Sub niTools_RaiseError(errorCode As Long, errorMsg As String, driver As String, Optional resourceName As String = "")
    Dim msg As String
    
    msg = "Error " & Format(errorCode) & " occurred." & vbNewLine & vbNewLine & errorMsg
    If resourceName <> "" Then
        msg = msg & vbNewLine & vbNewLine & "Resource Name: '" & resourceName & "'"
    End If
    
    Err.Raise niErrorNumber, driver & " error occured.", msg
End Sub

Public Sub niTools_ErrorMsgBox(e As ErrObject)
    If e.Number = niErrorNumber Then
        ' Show NI formated error info
        MsgBox e.Description, vbOKOnly + vbCritical + vbDefaultButton1 + vbApplicationModal, e.Source
    Else
        ' Try to mimic vba error
        MsgBox "Run-time error '" & Format(e.Number) & "': " & vbNewLine & vbNewLine & e.Description, _
            vbMsgBoxHelpButton + vbExclamation + vbDefaultButton1 + vbApplicationModal, _
            "Microsoft Visual Basic for Applications", e.HelpFile, e.HelpContext
    End If
End Sub

