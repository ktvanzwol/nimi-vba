Attribute VB_Name = "niTools"
Option Explicit

Type NIComplexNumber
    real As Double
    imaginary As Double
End Type

' Based on example found at https://github.com/ReneNyffenegger/VBA-calls-DLL/blob/master/return-char-array/vba.bas
Private Const CP_UTF8 As Long = 65001
Private Declare PtrSafe Function MultiByteToWideChar Lib "kernel32" ( _
    ByVal CodePage As Long, _
    ByVal dwFlags As Long, _
    ByVal lpMultiByteStr As LongPtr, _
    ByVal cbMultiByte As Long, _
    ByVal lpWideCharStr As LongPtr, _
    ByVal cchWideChar As Long _
) As Long

Private Declare PtrSafe Function GetLastError Lib "kernel32" () As Long

' VBA Error code used for NI Driver errors
Public Const niErrorNumber As Long = vbObjectError + 1024

' Utility fuction to convert C Strings returned from API calls to the VBA String Type.
Public Sub niTools_CStrPtrToStr(size As Long, cStrPtr As LongPtr, ByRef vbStr As String)
    Dim mbVal As Long
    Dim Error As Long
    
    ' Convert C String Pointer to VB String Type
    vbStr = String(size - 1, "*")
    mbVal = MultiByteToWideChar(CP_UTF8, 0, cStrPtr, -1, StrPtr(vbStr), size)

    If mbVal = 0 Then
        niTools_RaiseError -1, "'MultiByteToWideChar' in 'niTools_CStringPtrToString' failed.", "niTools"
    End If
End Sub

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
