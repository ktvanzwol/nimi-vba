Attribute VB_Name = "niTools"
Option Explicit

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

