VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RFmx_Session"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'int32 __stdcall RFmxInstr_Initialize (char resourceName[], char optionString[], niRFmxInstrHandle *handleOut, int32 *isNewSession);
Private Declare PtrSafe Function RFmxInstr_Initialize Lib "niRFmxInstr" ( _
    ByVal resourceName As String, ByVal optionString As String, ByRef handleOut As LongPtr, ByRef isNewSession As Long) As Long

'int32 __stdcall RFmxInstr_InitializeFromNIRFSASession(uInt32 NIRFSASession, niRFmxInstrHandle *handleOut);
Private Declare PtrSafe Function RFmxInstr_InitializeFromNIRFSASession Lib "niRFmxInstr" ( _
    ByVal niRFSASession As Long, ByRef handleOut As LongPtr) As Long
    
'int32 __stdcall RFmxInstr_GetNIRFSASession(niRFmxInstrHandle instrumentHandle, uInt32 *niRfsaSession);
Private Declare PtrSafe Function RFmxInstr_GetNIRFSASession Lib "niRFmxInstr" ( _
    ByVal handleOut As LongPtr, ByRef niRFSASession As Long) As Long

'int32 __stdcall RFmxInstr_Close (niRFmxInstrHandle instrumentHandle, int32 forceDestroy);
Private Declare PtrSafe Function RFmxInstr_Close Lib "niRFmxInstr" ( _
    ByVal instrumentHandle As LongPtr, ByVal forceDestroy As Long) As Long

'int32 __stdcall RFmxInstr_GetError (niRFmxInstrHandle instrumentHandle, int32* errorCode, int32 errorDescriptionBufferSize, char errorDescription[]);
Private Declare PtrSafe Function RFmxInstr_GetError Lib "niRFmxInstr" ( _
    ByVal instrumentHandle As LongPtr, ByRef errorCode As Long, ByVal errorDescriptionBufferSize As Long, ByVal errorDescription As LongPtr) As Long

'int32 __stdcall RFmxInstr_SetAttributeString(niRFmxInstrHandle instrumentHandle, char selectorString[], int32 attributeID, char attrVal[]);
'int32 __stdcall RFmxInstr_GetAttributeString(niRFmxInstrHandle instrumentHandle, char selectorString[], int32 attributeID, int32 arraySize, char attrVal[]);
Private Declare PtrSafe Function RFmxInstr_SetAttributeString Lib "niRFmxInstr" ( _
    ByVal instrumentHandle As LongPtr, ByVal selectorString As String, ByVal attributeID As Long, ByVal attrVal As String) As Long
Private Declare PtrSafe Function RFmxInstr_GetAttributeString Lib "niRFmxInstr" ( _
    ByVal instrumentHandle As LongPtr, ByVal selectorString As String, ByVal attributeID As Long, ByVal arraySize As Long, ByVal attrVal As LongPtr) As Long

'int32 __stdcall RFmxInstr_SetAttributeI8(niRFmxInstrHandle instrumentHandle, char selectorString[], int32 attributeID, int8 attrVal);
'int32 __stdcall RFmxInstr_GetAttributeI8(niRFmxInstrHandle instrumentHandle, char selectorString[], int32 attributeID, int8 *attrVal);
Private Declare PtrSafe Function RFmxInstr_SetAttributeI8 Lib "niRFmxInstr" ( _
    ByVal instrumentHandle As LongPtr, ByVal selectorString As String, ByVal attributeID As Long, ByVal attrVal As Byte) As Long
Private Declare PtrSafe Function RFmxInstr_GetAttributeI8 Lib "niRFmxInstr" ( _
    ByVal instrumentHandle As LongPtr, ByVal selectorString As String, ByVal attributeID As Long, ByRef attrVal As Byte) As Long

'int32 __stdcall RFmxInstr_SetAttributeU8(niRFmxInstrHandle instrumentHandle, char selectorString[], int32 attributeID, uInt8 attrVal);
'int32 __stdcall RFmxInstr_GetAttributeU8(niRFmxInstrHandle instrumentHandle, char selectorString[], int32 attributeID, uInt8 *attrVal);
Private Declare PtrSafe Function RFmxInstr_SetAttributeU8 Lib "niRFmxInstr" ( _
    ByVal instrumentHandle As LongPtr, ByVal selectorString As String, ByVal attributeID As Long, ByVal attrVal As Byte) As Long
Private Declare PtrSafe Function RFmxInstr_GetAttributeU8 Lib "niRFmxInstr" ( _
    ByVal instrumentHandle As LongPtr, ByVal selectorString As String, ByVal attributeID As Long, ByRef attrVal As Byte) As Long

'int32 __stdcall RFmxInstr_SetAttributeI16(niRFmxInstrHandle instrumentHandle, char selectorString[], int32 attributeID, int16 attrVal);
'int32 __stdcall RFmxInstr_GetAttributeI16(niRFmxInstrHandle instrumentHandle, char selectorString[], int32 attributeID, int16 *attrVal);
Private Declare PtrSafe Function RFmxInstr_SetAttributeI16 Lib "niRFmxInstr" ( _
    ByVal instrumentHandle As LongPtr, ByVal selectorString As String, ByVal attributeID As Long, ByVal attrVal As Integer) As Long
Private Declare PtrSafe Function RFmxInstr_GetAttributeI16 Lib "niRFmxInstr" ( _
    ByVal instrumentHandle As LongPtr, ByVal selectorString As String, ByVal attributeID As Long, ByRef attrVal As Integer) As Long

'int32 __stdcall RFmxInstr_SetAttributeU16(niRFmxInstrHandle instrumentHandle, char selectorString[], int32 attributeID, uInt16 attrVal);
'int32 __stdcall RFmxInstr_GetAttributeU16(niRFmxInstrHandle instrumentHandle, char selectorString[], int32 attributeID, uInt16 *attrVal);
Private Declare PtrSafe Function RFmxInstr_SetAttributeU16 Lib "niRFmxInstr" ( _
    ByVal instrumentHandle As LongPtr, ByVal selectorString As String, ByVal attributeID As Long, ByVal attrVal As Integer) As Long
Private Declare PtrSafe Function RFmxInstr_GetAttributeU16 Lib "niRFmxInstr" ( _
    ByVal instrumentHandle As LongPtr, ByVal selectorString As String, ByVal attributeID As Long, ByRef attrVal As Integer) As Long

'int32 __stdcall RFmxInstr_SetAttributeI32(niRFmxInstrHandle instrumentHandle, char selectorString[], int32 attributeID, int32 attrVal);
'int32 __stdcall RFmxInstr_GetAttributeI32(niRFmxInstrHandle instrumentHandle, char selectorString[], int32 attributeID, int32 *attrVal);
Private Declare PtrSafe Function RFmxInstr_SetAttributeI32 Lib "niRFmxInstr" ( _
    ByVal instrumentHandle As LongPtr, ByVal selectorString As String, ByVal attributeID As Long, ByVal attrVal As Long) As Long
Private Declare PtrSafe Function RFmxInstr_GetAttributeI32 Lib "niRFmxInstr" ( _
    ByVal instrumentHandle As LongPtr, ByVal selectorString As String, ByVal attributeID As Long, ByRef attrVal As Long) As Long

'int32 __stdcall RFmxInstr_SetAttributeU32(niRFmxInstrHandle instrumentHandle, char selectorString[], int32 attributeID, uInt32 attrVal);
'int32 __stdcall RFmxInstr_GetAttributeU32(niRFmxInstrHandle instrumentHandle, char selectorString[], int32 attributeID, uInt32 *attrVal);
Private Declare PtrSafe Function RFmxInstr_SetAttributeU32 Lib "niRFmxInstr" ( _
    ByVal instrumentHandle As LongPtr, ByVal selectorString As String, ByVal attributeID As Long, ByVal attrVal As Long) As Long
Private Declare PtrSafe Function RFmxInstr_GetAttributeU32 Lib "niRFmxInstr" ( _
    ByVal instrumentHandle As LongPtr, ByVal selectorString As String, ByVal attributeID As Long, ByRef attrVal As Long) As Long

'int32 __stdcall RFmxInstr_SetAttributeI64(niRFmxInstrHandle instrumentHandle, char selectorString[], int32 attributeID, int64 attrVal);
'int32 __stdcall RFmxInstr_GetAttributeI64(niRFmxInstrHandle instrumentHandle, char selectorString[], int32 attributeID, int64 *attrVal);
Private Declare PtrSafe Function RFmxInstr_SetAttributeI64 Lib "niRFmxInstr" ( _
    ByVal instrumentHandle As LongPtr, ByVal selectorString As String, ByVal attributeID As Long, ByVal attrVal As LongLong) As Long
Private Declare PtrSafe Function RFmxInstr_GetAttributeI64 Lib "niRFmxInstr" ( _
    ByVal instrumentHandle As LongPtr, ByVal selectorString As String, ByVal attributeID As Long, ByRef attrVal As LongLong) As Long

'int32 __stdcall RFmxInstr_SetAttributeF64(niRFmxInstrHandle instrumentHandle, char selectorString[], int32 attributeID, float64 attrVal);
'int32 __stdcall RFmxInstr_GetAttributeF64(niRFmxInstrHandle instrumentHandle, char selectorString[], int32 attributeID, float64 *attrVal);
Private Declare PtrSafe Function RFmxInstr_SetAttributeF64 Lib "niRFmxInstr" ( _
    ByVal instrumentHandle As LongPtr, ByVal selectorString As String, ByVal attributeID As Long, ByVal attrVal As Double) As Long
Private Declare PtrSafe Function RFmxInstr_GetAttributeF64 Lib "niRFmxInstr" ( _
    ByVal instrumentHandle As LongPtr, ByVal selectorString As String, ByVal attributeID As Long, ByRef attrVal As Double) As Long

'int32 __stdcall RFmxInstr_SetAttributeF32(niRFmxInstrHandle instrumentHandle, char selectorString[], int32 attributeID, float32 attrVal);
'int32 __stdcall RFmxInstr_GetAttributeF32(niRFmxInstrHandle instrumentHandle, char selectorString[], int32 attributeID, float32 *attrVal);
Private Declare PtrSafe Function RFmxInstr_SetAttributeF32 Lib "niRFmxInstr" ( _
    ByVal instrumentHandle As LongPtr, ByVal selectorString As String, ByVal attributeID As Long, ByVal attrVal As Single) As Long
Private Declare PtrSafe Function RFmxInstr_GetAttributeF32 Lib "niRFmxInstr" ( _
    ByVal instrumentHandle As LongPtr, ByVal selectorString As String, ByVal attributeID As Long, ByRef attrVal As Single) As Long

'int32 __stdcall RFmxInstr_SetAttributeI8Array(niRFmxInstrHandle instrumentHandle, char selectorString[], int32 attributeID, int8 attrVal[], int32 arraySize);
'int32 __stdcall RFmxInstr_GetAttributeI8Array(niRFmxInstrHandle instrumentHandle, char selectorString[], int32 attributeID, int8 attrVal[], int32 arraySize, int32 *actualArraySize);
Private Declare PtrSafe Function RFmxInstr_SetAttributeI8Array Lib "niRFmxInstr" ( _
    ByVal instrumentHandle As LongPtr, ByVal selectorString As String, ByVal attributeID As Long, ByVal attrVal As LongPtr, ByVal arraySize As Long) As Long
Private Declare PtrSafe Function RFmxInstr_GetAttributeI8Array Lib "niRFmxInstr" ( _
    ByVal instrumentHandle As LongPtr, ByVal selectorString As String, ByVal attributeID As Long, ByVal attrVal As LongPtr, ByVal arraySize As Long, ByRef actualArraySize As Long) As Long

'int32 __stdcall RFmxInstr_SetAttributeI32Array(niRFmxInstrHandle instrumentHandle, char selectorString[], int32 attributeID, int32 attrVal[], int32 arraySize);
'int32 __stdcall RFmxInstr_GetAttributeI32Array(niRFmxInstrHandle instrumentHandle, char selectorString[], int32 attributeID, int32 attrVal[], int32 arraySize, int32 *actualArraySize);
Private Declare PtrSafe Function RFmxInstr_SetAttributeI32Array Lib "niRFmxInstr" ( _
    ByVal instrumentHandle As LongPtr, ByVal selectorString As String, ByVal attributeID As Long, ByVal attrVal As LongPtr, ByVal arraySize As Long) As Long
Private Declare PtrSafe Function RFmxInstr_GetAttributeI32Array Lib "niRFmxInstr" ( _
    ByVal instrumentHandle As LongPtr, ByVal selectorString As String, ByVal attributeID As Long, ByVal attrVal As LongPtr, ByVal arraySize As Long, ByRef actualArraySize As Long) As Long

'int32 __stdcall RFmxInstr_SetAttributeI64Array(niRFmxInstrHandle instrumentHandle, char selectorString[], int32 attributeID, int64 attrVal[], int32 arraySize);
'int32 __stdcall RFmxInstr_GetAttributeI64Array(niRFmxInstrHandle instrumentHandle, char selectorString[], int32 attributeID, int64 attrVal[], int32 arraySize, int32 *actualArraySize);
Private Declare PtrSafe Function RFmxInstr_SetAttributeI64Array Lib "niRFmxInstr" ( _
    ByVal instrumentHandle As LongPtr, ByVal selectorString As String, ByVal attributeID As Long, ByVal attrVal As LongPtr, ByVal arraySize As Long) As Long
Private Declare PtrSafe Function RFmxInstr_GetAttributeI64Array Lib "niRFmxInstr" ( _
    ByVal instrumentHandle As LongPtr, ByVal selectorString As String, ByVal attributeID As Long, ByVal attrVal As LongPtr, ByVal arraySize As Long, ByRef actualArraySize As Long) As Long

'int32 __stdcall RFmxInstr_SetAttributeU64Array(niRFmxInstrHandle instrumentHandle, char selectorString[], int32 attributeID, uInt64 attrVal[], int32 arraySize);
'int32 __stdcall RFmxInstr_GetAttributeU64Array(niRFmxInstrHandle instrumentHandle, char selectorString[], int32 attributeID, uInt64 attrVal[], int32 arraySize, int32 *actualArraySize);
Private Declare PtrSafe Function RFmxInstr_SetAttributeU64Array Lib "niRFmxInstr" ( _
    ByVal instrumentHandle As LongPtr, ByVal selectorString As String, ByVal attributeID As Long, ByVal attrVal As LongPtr, ByVal arraySize As Long) As Long
Private Declare PtrSafe Function RFmxInstr_GetAttributeU64Array Lib "niRFmxInstr" ( _
    ByVal instrumentHandle As LongPtr, ByVal selectorString As String, ByVal attributeID As Long, ByVal attrVal As LongPtr, ByVal arraySize As Long, ByRef actualArraySize As Long) As Long

'int32 __stdcall RFmxInstr_SetAttributeU8Array(niRFmxInstrHandle instrumentHandle, char selectorString[], int32 attributeID, uInt8 attrVal[], int32 arraySize);
'int32 __stdcall RFmxInstr_GetAttributeU8Array(niRFmxInstrHandle instrumentHandle, char selectorString[], int32 attributeID, uInt8 attrVal[], int32 arraySize, int32 *actualArraySize);
Private Declare PtrSafe Function RFmxInstr_SetAttributeU8Array Lib "niRFmxInstr" ( _
    ByVal instrumentHandle As LongPtr, ByVal selectorString As String, ByVal attributeID As Long, ByVal attrVal As LongPtr, ByVal arraySize As Long) As Long
Private Declare PtrSafe Function RFmxInstr_GetAttributeU8Array Lib "niRFmxInstr" ( _
    ByVal instrumentHandle As LongPtr, ByVal selectorString As String, ByVal attributeID As Long, ByVal attrVal As LongPtr, ByVal arraySize As Long, ByRef actualArraySize As Long) As Long

'int32 __stdcall RFmxInstr_SetAttributeU32Array(niRFmxInstrHandle instrumentHandle, char selectorString[], int32 attributeID, uInt32 attrVal[], int32 arraySize);
'int32 __stdcall RFmxInstr_GetAttributeU32Array(niRFmxInstrHandle instrumentHandle, char selectorString[], int32 attributeID, uInt32 attrVal[], int32 arraySize, int32 *actualArraySize);
Private Declare PtrSafe Function RFmxInstr_SetAttributeU32Array Lib "niRFmxInstr" ( _
    ByVal instrumentHandle As LongPtr, ByVal selectorString As String, ByVal attributeID As Long, ByVal attrVal As LongPtr, ByVal arraySize As Long) As Long
Private Declare PtrSafe Function RFmxInstr_GetAttributeU32Array Lib "niRFmxInstr" ( _
    ByVal instrumentHandle As LongPtr, ByVal selectorString As String, ByVal attributeID As Long, ByVal attrVal As LongPtr, ByVal arraySize As Long, ByRef actualArraySize As Long) As Long

'int32 __stdcall RFmxInstr_SetAttributeF32Array(niRFmxInstrHandle instrumentHandle, char selectorString[], int32 attributeID, float32 attrVal[], int32 arraySize);
'int32 __stdcall RFmxInstr_GetAttributeF32Array(niRFmxInstrHandle instrumentHandle, char selectorString[], int32 attributeID, float32 attrVal[], int32 arraySize, int32 *actualArraySize);
Private Declare PtrSafe Function RFmxInstr_SetAttributeF32Array Lib "niRFmxInstr" ( _
    ByVal instrumentHandle As LongPtr, ByVal selectorString As String, ByVal attributeID As Long, ByVal attrVal As LongPtr, ByVal arraySize As Long) As Long
Private Declare PtrSafe Function RFmxInstr_GetAttributeF32Array Lib "niRFmxInstr" ( _
    ByVal instrumentHandle As LongPtr, ByVal selectorString As String, ByVal attributeID As Long, ByVal attrVal As LongPtr, ByVal arraySize As Long, ByRef actualArraySize As Long) As Long

'int32 __stdcall RFmxInstr_SetAttributeF64Array(niRFmxInstrHandle instrumentHandle, char selectorString[], int32 attributeID, float64 attrVal[], int32 arraySize);
'int32 __stdcall RFmxInstr_GetAttributeF64Array(niRFmxInstrHandle instrumentHandle, char selectorString[], int32 attributeID, float64 attrVal[], int32 arraySize, int32 *actualArraySize);
Private Declare PtrSafe Function RFmxInstr_SetAttributeF64Array Lib "niRFmxInstr" ( _
    ByVal instrumentHandle As LongPtr, ByVal selectorString As String, ByVal attributeID As Long, ByVal attrVal As LongPtr, ByVal arraySize As Long) As Long
Private Declare PtrSafe Function RFmxInstr_GetAttributeF64Array Lib "niRFmxInstr" ( _
    ByVal instrumentHandle As LongPtr, ByVal selectorString As String, ByVal attributeID As Long, ByVal attrVal As LongPtr, ByVal arraySize As Long, ByRef actualArraySize As Long) As Long

'int32 __stdcall RFmxInstr_SetAttributeNIComplexSingleArray(niRFmxInstrHandle instrumentHandle, char selectorString[], int32 attributeID, NIComplexSingle attrVal[], int32 arraySize);
'int32 __stdcall RFmxInstr_GetAttributeNIComplexSingleArray(niRFmxInstrHandle instrumentHandle, char selectorString[], int32 attributeID, NIComplexSingle attrVal[], int32 arraySize, int32 *actualArraySize);
Private Declare PtrSafe Function RFmxInstr_SetAttributeNIComplexSingleArray Lib "niRFmxInstr" ( _
    ByVal instrumentHandle As LongPtr, ByVal selectorString As String, ByVal attributeID As Long, ByVal attrVal As LongPtr, ByVal arraySize As Long) As Long
Private Declare PtrSafe Function RFmxInstr_GetAttributeNIComplexSingleArray Lib "niRFmxInstr" ( _
    ByVal instrumentHandle As LongPtr, ByVal selectorString As String, ByVal attributeID As Long, ByVal attrVal As LongPtr, ByVal arraySize As Long, ByRef actualArraySize As Long) As Long

'int32 __stdcall RFmxInstr_SetAttributeNIComplexDoubleArray(niRFmxInstrHandle instrumentHandle, char selectorString[], int32 attributeID, NIComplexDouble attrVal[], int32 arraySize);
'int32 __stdcall RFmxInstr_GetAttributeNIComplexDoubleArray(niRFmxInstrHandle instrumentHandle, char selectorString[], int32 attributeID, NIComplexDouble attrVal[], int32 arraySize, int32 *actualArraySize);
Private Declare PtrSafe Function RFmxInstr_SetAttributeNIComplexDoubleArray Lib "niRFmxInstr" ( _
    ByVal instrumentHandle As LongPtr, ByVal selectorString As String, ByVal attributeID As Long, ByVal attrVal As LongPtr, ByVal arraySize As Long) As Long
Private Declare PtrSafe Function RFmxInstr_GetAttributeNIComplexDoubleArray Lib "niRFmxInstr" ( _
    ByVal instrumentHandle As LongPtr, ByVal selectorString As String, ByVal attributeID As Long, ByVal attrVal As LongPtr, ByVal arraySize As Long, ByRef actualArraySize As Long) As Long

'int32 __stdcall RFmxInstr_CfgFrequencyReference (niRFmxInstrHandle instrumentHandle, char selectorString[], char frequencyReferenceSource[], float64 frequencyReferenceFrequency);
Private Declare PtrSafe Function RFmxInstr_CfgFrequencyReference Lib "niRFmxInstr" ( _
    ByVal instrumentHandle As LongPtr, ByVal selectorString As String, ByVal frequencyReferenceSource As String, ByVal frequencyReferenceFrequency As Double) As Long

' Internal session
Private m_Handle As LongPtr
Private m_ResourceName As String
Private m_RFSAOwnedSession As Boolean

' Personality objects
Private m_SpecAn As RFmx_SpecAn

Public Property Get SpecAn() As RFmx_SpecAn
    Set SpecAn = m_SpecAn
End Property

Private Sub InitRFmxPersonalities()
    Set m_SpecAn = New RFmx_SpecAn
    m_SpecAn.InitSpecAn m_Handle
End Sub

Private Sub ClearRFmxPersonalities()
    Set m_SpecAn = Nothing
End Sub

' initialize internal variables, call Init first to create a valid session
Private Sub Class_Initialize()
    m_Handle = 0
    m_ResourceName = ""
    m_RFSAOwnedSession = False
    ClearRFmxPersonalities
End Sub

' Automatically clear session when object gets destroyed
Private Sub Class_Terminate()
    CloseSession
End Sub

' Error Checker
Private Sub CheckError(status As Long)
    If status < 0 Then
        ErrorHandler status
    End If
End Sub

' Error Handler
Private Sub ErrorHandler(errorCode As Long)
    Dim status As Long
    Dim size As Long
    Dim buffer() As Byte
    Dim errorMsg As String
    
    size = RFmxInstr_GetError(m_Handle, errorCode, 0, 0)
    ReDim buffer(size - 1) As Byte
 
    status = RFmxInstr_GetError(m_Handle, errorCode, size, VarPtr(buffer(0)))
    errorMsg = StrConv(LeftB(buffer(), size - 1), vbUnicode) 'Remove \0 character and convert to Unicode
    
    niTools_RaiseError errorCode, errorMsg, "NI-RFmxInstr"
End Sub

Public Sub InitSession(resourceName As String, optionString As String, ByRef isNewSession As RFmx_Binary)
    ' Make sure session is closed before opening
    CloseSession
    
    m_ResourceName = resourceName
    m_RFSAOwnedSession = False
    CheckError RFmxInstr_Initialize(resourceName, optionString, m_Handle, isNewSession)
    
    InitRFmxPersonalities
End Sub

Public Sub InitSessionFromNIRFSASession(rfsa As niRFSA_Session)
    ' Make sure session is closed before opening
    CloseSession
    
    m_ResourceName = rfsa.InternalResourceName
    m_RFSAOwnedSession = True
    CheckError RFmxInstr_InitializeFromNIRFSASession(rfsa.InternalSession, m_Handle)
    
    InitRFmxPersonalities
End Sub

Private Sub CloseSession()
    If m_Handle = 0 Then Exit Sub
    
    If m_RFSAOwnedSession = False Then
        CheckError RFmxInstr_Close(m_Handle, RFMX_VAL_FALSE)
    End If
    
    m_Handle = 0
    m_ResourceName = ""
    m_RFSAOwnedSession = False
    ClearRFmxPersonalities
End Sub

Public Property Get GetNIRFSASession() As niRFSA_Session
    Dim rfsaSession As niRFSA_Session
    Dim session As Long
    
    CheckError RFmxInstr_GetNIRFSASession(m_Handle, session)
    
    Set rfsaSession = New niRFSA_Session
    rfsaSession.InitSessionForRFmxGetNIRFSASession m_ResourceName, session
    
    Set GetNIRFSASession = rfsaSession
End Property

Public Sub SetAttributeString(selectorString As String, attributeID As RFmx_AttributeIDs, attrVal As String)
    CheckError RFmxInstr_SetAttributeString(m_Handle, selectorString, attributeID, attrVal)
End Sub

Public Sub GetAttributeString(selectorString As String, attributeID As RFmx_AttributeIDs, ByRef attrVal As String)
    Dim size As Long
    Dim buffer() As Byte
    
    size = RFmxInstr_GetAttributeString(m_Handle, selectorString, attributeID, 0, 0)
    ReDim buffer(size - 1) As Byte

    CheckError RFmxInstr_GetAttributeString(m_Handle, selectorString, attributeID, size, VarPtr(buffer(0)))
    attrVal = StrConv(LeftB(buffer(), size - 1), vbUnicode) ' Remove \0 character and convert to unicode
End Sub

Public Sub SetAttributeI8(selectorString As String, attributeID As RFmx_AttributeIDs, attrVal As Byte)
    CheckError RFmxInstr_SetAttributeI8(m_Handle, selectorString, attributeID, attrVal)
End Sub

Public Sub GetAttributeI8(selectorString As String, attributeID As RFmx_AttributeIDs, ByRef attrVal As Byte)
    CheckError RFmxInstr_GetAttributeI8(m_Handle, selectorString, attributeID, attrVal)
End Sub

Public Sub SetAttributeU8(selectorString As String, attributeID As RFmx_AttributeIDs, attrVal As Byte)
    CheckError RFmxInstr_SetAttributeU8(m_Handle, selectorString, attributeID, attrVal)
End Sub

Public Sub GetAttributeU8(selectorString As String, attributeID As RFmx_AttributeIDs, ByRef attrVal As Byte)
    CheckError RFmxInstr_GetAttributeU8(m_Handle, selectorString, attributeID, attrVal)
End Sub

Public Sub SetAttributeI16(selectorString As String, attributeID As RFmx_AttributeIDs, attrVal As Integer)
    CheckError RFmxInstr_SetAttributeI16(m_Handle, selectorString, attributeID, attrVal)
End Sub

Public Sub GetAttributeI16(selectorString As String, attributeID As RFmx_AttributeIDs, ByRef attrVal As Integer)
    CheckError RFmxInstr_GetAttributeI16(m_Handle, selectorString, attributeID, attrVal)
End Sub

Public Sub SetAttributeU16(selectorString As String, attributeID As RFmx_AttributeIDs, attrVal As Integer)
    CheckError RFmxInstr_SetAttributeU16(m_Handle, selectorString, attributeID, attrVal)
End Sub

Public Sub GetAttributeU16(selectorString As String, attributeID As RFmx_AttributeIDs, ByRef attrVal As Integer)
    CheckError RFmxInstr_GetAttributeU16(m_Handle, selectorString, attributeID, attrVal)
End Sub

Public Sub SetAttributeI32(selectorString As String, attributeID As RFmx_AttributeIDs, attrVal As Long)
    CheckError RFmxInstr_SetAttributeI32(m_Handle, selectorString, attributeID, attrVal)
End Sub

Public Sub GetAttributeI32(selectorString As String, attributeID As RFmx_AttributeIDs, ByRef attrVal As Long)
    CheckError RFmxInstr_GetAttributeI32(m_Handle, selectorString, attributeID, attrVal)
End Sub

Public Sub SetAttributeU32(selectorString As String, attributeID As RFmx_AttributeIDs, attrVal As Long)
    CheckError RFmxInstr_SetAttributeU32(m_Handle, selectorString, attributeID, attrVal)
End Sub

Public Sub GetAttributeU32(selectorString As String, attributeID As RFmx_AttributeIDs, ByRef attrVal As Long)
    CheckError RFmxInstr_GetAttributeU32(m_Handle, selectorString, attributeID, attrVal)
End Sub

Public Sub SetAttributeI64(selectorString As String, attributeID As RFmx_AttributeIDs, attrVal As LongLong)
    CheckError RFmxInstr_SetAttributeI64(m_Handle, selectorString, attributeID, attrVal)
End Sub

Public Sub GetAttributeI64(selectorString As String, attributeID As RFmx_AttributeIDs, ByRef attrVal As LongLong)
    CheckError RFmxInstr_GetAttributeI64(m_Handle, selectorString, attributeID, attrVal)
End Sub

Public Sub SetAttributeF64(selectorString As String, attributeID As RFmx_AttributeIDs, attrVal As Double)
    CheckError RFmxInstr_SetAttributeF64(m_Handle, selectorString, attributeID, attrVal)
End Sub

Public Sub GetAttributeF64(selectorString As String, attributeID As RFmx_AttributeIDs, ByRef attrVal As Double)
    CheckError RFmxInstr_GetAttributeF64(m_Handle, selectorString, attributeID, attrVal)
End Sub

Public Sub SetAttributeF32(selectorString As String, attributeID As RFmx_AttributeIDs, attrVal As Single)
    CheckError RFmxInstr_SetAttributeF32(m_Handle, selectorString, attributeID, attrVal)
End Sub

Public Sub GetAttributeF32(selectorString As String, attributeID As RFmx_AttributeIDs, ByRef attrVal As Single)
    CheckError RFmxInstr_GetAttributeF32(m_Handle, selectorString, attributeID, attrVal)
End Sub

Public Sub SetAttributeI8Array(selectorString As String, attributeID As RFmx_AttributeIDs, attrVal() As Byte)
    CheckError RFmxInstr_SetAttributeI8Array(m_Handle, selectorString, attributeID, _
                                                VarPtr(attrVal(LBound(attrVal))), _
                                                UBound(attrVal) - LBound(attrVal) + 1)
End Sub

Public Sub GetAttributeI8Array(selectorString As String, attributeID As RFmx_AttributeIDs, ByRef attrVal() As Byte)
    Dim size As Long
    
    CheckError RFmxInstr_GetAttributeI8Array(m_Handle, selectorString, attributeID, 0, 0, size)
    ReDim attrVal(size - 1) As Byte

    CheckError RFmxInstr_GetAttributeI8Array(m_Handle, selectorString, attributeID, VarPtr(attrVal(0)), size, size)
End Sub

Public Sub SetAttributeU8Array(selectorString As String, attributeID As RFmx_AttributeIDs, attrVal() As Byte)
    CheckError RFmxInstr_SetAttributeU8Array(m_Handle, selectorString, attributeID, _
                                                VarPtr(attrVal(LBound(attrVal))), _
                                                UBound(attrVal) - LBound(attrVal) + 1)
End Sub

Public Sub GetAttributeU8Array(selectorString As String, attributeID As RFmx_AttributeIDs, ByRef attrVal() As Byte)
    Dim size As Long
    
    CheckError RFmxInstr_GetAttributeU8Array(m_Handle, selectorString, attributeID, 0, 0, size)
    ReDim attrVal(size - 1) As Byte

    CheckError RFmxInstr_GetAttributeU8Array(m_Handle, selectorString, attributeID, VarPtr(attrVal(0)), size, size)
End Sub

Public Sub SetAttributeI32Array(selectorString As String, attributeID As RFmx_AttributeIDs, attrVal() As Long)
    CheckError RFmxInstr_SetAttributeI32Array(m_Handle, selectorString, attributeID, _
                                                VarPtr(attrVal(LBound(attrVal))), _
                                                UBound(attrVal) - LBound(attrVal) + 1)
End Sub

Public Sub GetAttributeI32Array(selectorString As String, attributeID As RFmx_AttributeIDs, ByRef attrVal() As Long)
    Dim size As Long
    
    CheckError RFmxInstr_GetAttributeI32Array(m_Handle, selectorString, attributeID, 0, 0, size)
    ReDim attrVal(size - 1) As Long

    CheckError RFmxInstr_GetAttributeI32Array(m_Handle, selectorString, attributeID, VarPtr(attrVal(0)), size, size)
End Sub

Public Sub SetAttributeU32Array(selectorString As String, attributeID As RFmx_AttributeIDs, attrVal() As Long)
    CheckError RFmxInstr_SetAttributeU32Array(m_Handle, selectorString, attributeID, _
                                                VarPtr(attrVal(LBound(attrVal))), _
                                                UBound(attrVal) - LBound(attrVal) + 1)
End Sub

Public Sub GetAttributeU32Array(selectorString As String, attributeID As RFmx_AttributeIDs, ByRef attrVal() As Long)
    Dim size As Long
    
    CheckError RFmxInstr_GetAttributeU32Array(m_Handle, selectorString, attributeID, 0, 0, size)
    ReDim attrVal(size - 1) As Long

    CheckError RFmxInstr_GetAttributeU32Array(m_Handle, selectorString, attributeID, VarPtr(attrVal(0)), size, size)
End Sub

Public Sub SetAttributeI64Array(selectorString As String, attributeID As RFmx_AttributeIDs, attrVal() As LongLong)
    CheckError RFmxInstr_SetAttributeI64Array(m_Handle, selectorString, attributeID, _
                                                VarPtr(attrVal(LBound(attrVal))), _
                                                UBound(attrVal) - LBound(attrVal) + 1)
End Sub

Public Sub GetAttributeI64Array(selectorString As String, attributeID As RFmx_AttributeIDs, ByRef attrVal() As LongLong)
    Dim size As Long
    
    CheckError RFmxInstr_GetAttributeI64Array(m_Handle, selectorString, attributeID, 0, 0, size)
    ReDim attrVal(size - 1) As LongLong

    CheckError RFmxInstr_GetAttributeI64Array(m_Handle, selectorString, attributeID, VarPtr(attrVal(0)), size, size)
End Sub

Public Sub SetAttributeU64Array(selectorString As String, attributeID As RFmx_AttributeIDs, attrVal() As LongLong)
    CheckError RFmxInstr_SetAttributeU64Array(m_Handle, selectorString, attributeID, _
                                                VarPtr(attrVal(LBound(attrVal))), _
                                                UBound(attrVal) - LBound(attrVal) + 1)
End Sub

Public Sub GetAttributeU64Array(selectorString As String, attributeID As RFmx_AttributeIDs, ByRef attrVal() As LongLong)
    Dim size As Long
    
    CheckError RFmxInstr_GetAttributeU64Array(m_Handle, selectorString, attributeID, 0, 0, size)
    ReDim attrVal(size - 1) As LongLong

    CheckError RFmxInstr_GetAttributeU64Array(m_Handle, selectorString, attributeID, VarPtr(attrVal(0)), size, size)
End Sub

Public Sub SetAttributeF64Array(selectorString As String, attributeID As RFmx_AttributeIDs, attrVal() As Double)
    CheckError RFmxInstr_SetAttributeU64Array(m_Handle, selectorString, attributeID, _
                                                VarPtr(attrVal(LBound(attrVal))), _
                                                UBound(attrVal) - LBound(attrVal) + 1)
End Sub

Public Sub GetAttributeF64Array(selectorString As String, attributeID As RFmx_AttributeIDs, ByRef attrVal() As Double)
    Dim size As Long
    
    CheckError RFmxInstr_GetAttributeF64Array(m_Handle, selectorString, attributeID, 0, 0, size)
    ReDim attrVal(size - 1) As Double

    CheckError RFmxInstr_GetAttributeF64Array(m_Handle, selectorString, attributeID, VarPtr(attrVal(0)), size, size)
End Sub

Public Sub SetAttributeF32Array(selectorString As String, attributeID As RFmx_AttributeIDs, attrVal() As Single)
    CheckError RFmxInstr_SetAttributeU64Array(m_Handle, selectorString, attributeID, _
                                                VarPtr(attrVal(LBound(attrVal))), _
                                                UBound(attrVal) - LBound(attrVal) + 1)
End Sub

Public Sub GetAttributeF32Array(selectorString As String, attributeID As RFmx_AttributeIDs, ByRef attrVal() As Single)
    Dim size As Long
    
    CheckError RFmxInstr_GetAttributeF32Array(m_Handle, selectorString, attributeID, 0, 0, size)
    ReDim attrVal(size - 1) As Single

    CheckError RFmxInstr_GetAttributeF32Array(m_Handle, selectorString, attributeID, VarPtr(attrVal(0)), size, size)
End Sub

Public Sub SetAttributeNIComplexSingleArray(selectorString As String, attributeID As RFmx_AttributeIDs, attrVal() As NIComplexSingle)
    CheckError RFmxInstr_SetAttributeNIComplexSingleArray(m_Handle, selectorString, attributeID, _
                                                            VarPtr(attrVal(LBound(attrVal))), _
                                                            UBound(attrVal) - LBound(attrVal) + 1)
End Sub

Public Sub GetAttributeNIComplexSingleArray(selectorString As String, attributeID As RFmx_AttributeIDs, ByRef attrVal() As NIComplexSingle)
    Dim size As Long
    
    CheckError RFmxInstr_GetAttributeNIComplexSingleArray(m_Handle, selectorString, attributeID, 0, 0, size)
    ReDim attrVal(size - 1) As NIComplexSingle

    CheckError RFmxInstr_GetAttributeNIComplexSingleArray(m_Handle, selectorString, attributeID, VarPtr(attrVal(0)), size, size)
End Sub

Public Sub SetAttributeNIComplexDoubleArray(selectorString As String, attributeID As RFmx_AttributeIDs, attrVal() As NIComplexDouble)
    CheckError RFmxInstr_SetAttributeNIComplexDoubleArray(m_Handle, selectorString, attributeID, _
                                                            VarPtr(attrVal(LBound(attrVal))), _
                                                            UBound(attrVal) - LBound(attrVal) + 1)
End Sub

Public Sub GetAttributeNIComplexDoubleArray(selectorString As String, attributeID As RFmx_AttributeIDs, ByRef attrVal() As NIComplexDouble)
    Dim size As Long
    
    CheckError RFmxInstr_GetAttributeNIComplexDoubleArray(m_Handle, selectorString, attributeID, 0, 0, size)
    ReDim attrVal(size - 1) As NIComplexDouble

    CheckError RFmxInstr_GetAttributeNIComplexDoubleArray(m_Handle, selectorString, attributeID, VarPtr(attrVal(0)), size, size)
End Sub

Public Sub CfgFrequencyReference(selectorString As String, frequencyReferenceSource As String, frequencyReferenceFrequency As Double)
    CheckError RFmxInstr_CfgFrequencyReference(m_Handle, selectorString, frequencyReferenceSource, frequencyReferenceFrequency)
End Sub
