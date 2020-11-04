# nimi-vba

VBA7 Win64 wrappers proof of concept for NI Modular Instrument C Drivers.

The wrappers use the VBA Declare statement to directly reference the NI Modular Instrument DLLs implementing the C API's. See the [Declare statement Microsoft Documentation](https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/declare-statement) for details on how to use the declare statement.

Currently this repository is a proof of concept and by no means complete. It currently include implementations for the following  Modular Instrument Drivers and Libraries:

- NI-ModInst
- NI-DMM
- NI-DCPower
- NI-568x
- NI-RFSA
- NI-RFSG
- NI RFSG Playback Library
- NI RFmx Instrument
- NI RFmx SpecAn

For each component the framework is implemented and the functions call for required for one or two examples are included. The components can easily be extended to map additional C functions on an as needed bases.

## Using nimi-vba

Download the contens of this repositoroy using your prefered way.

### Manual Import

To install and test the nimi-vba manually import you can import the modules from the sub folders under the **src** folder in this repository. When importing a component always import all modules in a sub folder and you will always need to import the modules in the **niTools** folder.

### Automatic Import

Inorder to automatically import you first need to enable *Access to the VBA Project object model* in the Trust Center. To do this go into the **Excel Options**, Select **Trust Center** and click the **Trust Center Settings...** button. Select **Macro Settings** and place a check mark before **Trust access to the VBA object model**

Once this is done open the **nimi-vba ExcelTool.xlsm** file then on the nimi-vba sheet set the **Target Workbook** to *\<Create new Workbook\>* and click the Import/Update button on the sheet. This will create a new sheet and import all the nimi-vba modules into it.

Alternatively you can open anexsition workbook first and then select this as the Target Workbook to Import/Update nimi-vba in an existing application.

### Exporting

If you fixed issues and or extended nimi-vba you can export modules manually or use the same **nimi-vba ExcelTool.xlsm** file to automatically export modules back into the **src** folder structure. To automatically export the modules open **nimi-vba ExcelTool.xlsm** and the excel application with the updates. Next Select the excel application workbook with the updates as the **Target Workbook** and click the Export button.

## Mapping C types to VBA Types

The following table shows how to map the common C datatypes used in the NI Modular Instrument drivers C APIs to VBA supported types with the VBA Declare statement.

| IVI / VISA data type | C data type | Visual Basic Type | Conversion needs |
| --- | --- | --- | --- |
| ``ViUInt64`` | ``unsigned __int64`` | ``LongLong`` | No unsigned ``LongLong`` type in VBA |
| ``ViInt64`` | ``signed __int64`` | ``LongLong`` | |
| ``ViUInt32`` | ``unsigned long`` | ``Long`` | No unsigned ``Long`` type in VBA |
| ``ViInt32`` | ``signed long`` | ``Long`` | |
| ``ViUInt16`` | ``unsigned short`` | ``Integer`` | No unsigned ``Integer`` type in VBA |
| ``ViInt16`` | ``signed short`` | ``Integer`` | |
| ``ViUInt8`` | ``unsigned char`` | ``Byte`` | |
| ``ViInt8`` | ``signed char`` | ``Byte`` | No signed ``Byte`` type in VBA |
| ``ViByte`` | ``unsigned char`` | ``Byte`` | |
| ``ViChar`` | ``char`` | ``Byte`` | No signed ``Byte`` type in VBA |
| ``ViReal32`` | ``float`` | ``Single`` | |
| ``ViReal64`` | ``double`` | ``Double`` | |
| ``ViBoolean`` | ``unsigned short`` | ``Boolean`` | |
| ``ViString`` | ``char *`` | ``ByVal LongPtr`` or ``ByVal String`` | See [Using Pointers and Strings](#Using-Pointers-and-Strings) |
| ``ViConstString`` | ``const char *`` | ``ByVal String`` | See [Using the ByVal str As String declaration](#Using-the-ByVal-str-As-String-declaration) |
| ``ViRsrc`` | ``char *`` | ``ByVal LongPtr`` or ``ByVal String`` | See [Using Pointers and Strings](#Using-Pointers-and-Strings) |
| ``ViStatus`` | ``signed long`` | ``Long`` | |
| ``ViSession`` | ``unsigned long`` | ``Long`` |  No unsigned ``Long`` in VBA |
| ``niRFmxInstrHandle`` | ``void *`` | ``ByVal LongPtr`` | See [Using Pointers and Strings](#Using-Pointers-and-Strings) |

### By Value vs By Reference

In general when passing C arguments By Value you need to specify ``ByVal`` for the argument in the VBA Declare statement. By default the VBA Declare statement will assume passing the the C argument by reference (e.g. using a pointer). However it is good practise to specify ``ByRef`` in the VBA declare statement in this case.

Example of a C Function followed by the corresponding VBA Declare statement.

```C
ViStatus _VI_FUNC niDMM_GetAttributeViReal64(ViSession vi, ViConstString channelName, ViAttr attributeId, ViReal64 *value);
```

```VBA
Private Declare PtrSafe Function niDMM_GetAttributeViReal64 Lib "niDMM_64" ( _
    ByVal vi As Long, ByVal channelName As String, ByVal attributeID As Long, ByRef value As Double) As Long
```

*Note how ByRef is used on the ``value`` argument.*

There are a few special cases that requires handling the pointer values manually to correctly pass data. This are discussed in the next section

### Array Pointers

Internally VBA stores Array types as a SAFEARRAY structure, as a result passing a array type ``ByRef`` will not work. C will expect the pointer to the start of the first array element. This can also be done on within VBA but requires you to pass the pointer by value and dereference is in VBA. To do this you need to define the array pointer argument as ``ByVal value As LongPtr`` in the VBA Declare statement. This will pass the value of the pointer between C and VBA.

Example of a C Function that contains a array pointer (``attrVal[]``) and the corresponding VBA Declare statement.

```C
int32 __stdcall RFmxInstr_GetAttributeF64Array(niRFmxInstrHandle instrumentHandle, char selectorString[], int32 attributeID, float64 attrVal[], int32 arraySize, int32 *actualArraySize);
```

```VBA
Private Declare PtrSafe Function RFmxInstr_GetAttributeF64Array Lib "niRFmxInstr" ( _
    ByVal instrumentHandle As LongPtr, ByVal selectorString As String, ByVal attributeID As Long, ByVal attrVal As LongPtr, ByVal arraySize As Long, ByRef actualArraySize As Long) As Long
```

Inorder to pass a VBA array to C you now need to get the pointer to the start of the first element of the array. This can be done using the ``VarPtr()`` function. To get the right pointer value use the ``VarPtr()`` function on the first ellement of the array variable like shown below.

```VBA
VarPtr( myArrayVarable( 0 ) )

' Or the more generic
VarPtr( myArrayVarable( LBound( myArrayVarable ) ) )
```

### Strings

VBA Strings BSTR's which are UNICODE encodes strings preceded by the string length and in C these are basic ASCII strings with a null terminating characters. This means strings need to be converted when passed between VBA and C.

When dealing static string beeing passed from VBA into a C DLL this can be done directly by VBA bu defining the argument as ``ByVal String`` In this case you can pass a VBA String directly to the C DLL and VBA will be automatically converted the string from BSTR to an ASCII null terminated string.

When you pass strings from the C DLL back to VBA you will need to do this conversion your self. 

C Strings are stored as Byte Arrays ``char *`` so we can use the same aproach as for arrays to receive a C style string from external code.
But VBA uses Ninicode and wide characters for its ``String`` type so we also need to deal with the convertion between Byte Array and String types.

#### Receiving Strings

When receiving strings you typically need to pass the pointer to a pre allocated string of the right lenght to the C function. When the function returned you then need to use the ``StrConv`` function with the byte array to convert to Unicode.

Example of receiving a string using the ``niDMM_GetError`` function:

```VBA
'ViStatus _VI_FUNC niDMM_GetError(ViSession vi, ViStatus *errorCode, ViInt32 bufferSize, ViChar description[]);
Private Declare PtrSafe Function niDMM_GetError Lib "niDMM_64" ( _
    ByVal vi As Long, _
    ByRef errorCode As Long, _
    ByVal bufferSize As Long, _
    ByVal errMessage As LongPtr _
) As Long

Sub GetErrorMessage(m_Session As Long, errorCode As Long, ByRef errorMsg As String)
    Dim status As Long
    Dim size As Long
    Dim buffer() As Byte
    Dim errorMsg As String

    size = niDMM_GetError(m_Session, errorCode, 0, 0)
    ReDim buffer(size) As Byte

    status = niDMM_GetError(m_Session, errorCode, size, VarPtr(buffer(0)))
    errorMsg = StrConv(buffer(), vbUnicode)
End Sub
```

First note that the errorMsg ``char *`` argument is declared ``ByVal`` and as a ``LongPtr``. ``niDMM_GetError`` is first called with size set to 0 and a ``NULL`` pointer value, this will return the number of bytes needed for the buffer. Then ``ReDim`` is used to allocate the needed buffer and on the second call the pointer to this buffer is passed to the function using ``VarPtr(buffer(0))``.

And as the final step ``errorMsg = StrConv(buffer(), vbUnicode)`` converts the C stype string to the Unicode String used by VBA.

#### Sending Strings

The ``StrConv`` function can also be used to convert from Unicode strings to C Style strings:

```VBA
buffer() = StrConv("StackOverflow", vbFromUnicode)
```

This statement converts a Unicode String to a C Style string stored in a Byte array. In the same way you can pass the pointer to this Byte array to a C function like so: ``VarPtr(buffer(0))``.

#### Using the ``ByVal str As String`` declaration

VBA can handle the unicode convertion automatically when you define the argument using a ``ByVal String``. This works on most cases when you are dealing with input only arguments, typically defined as ``const char *`` in C. As soon as you need to be able to pass a ``NULL`` pointer value you need to use the more generic ``LongPtr`` and Byte array declaration method wich is more typical when receiving strings.

## VBA Resources

- [Declare statement (VBA) | Microsoft Docs](https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/declare-statement)
- [VBA Reference | bytecomb](https://bytecomb.com/vba-reference/)
  - [VBA Internals: What's in a variable | bytecomb](https://bytecomb.com/vba-internals-whats-in-a-variable/)
  - [VBA Internals: Getting Pointers | bytecomb](https://bytecomb.com/vba-internals-getting-pointers/)
  - [VBA Internals: String Variables and Pointers in Depth | bytecomb](https://bytecomb.com/vba-internals-string-variables-and-pointers-in-depth/)
  - [VBA Internals: Array Variables and Pointers in Depth | bytecomb](https://bytecomb.com/vba-internals-array-variables-and-pointers-in-depth/)
- [excel - Convert an array of bytes into a string? - Stack Overflow](https://stackoverflow.com/questions/50449004/convert-an-array-of-bytes-into-a-string)
- [VBA Articles - Excel Macro Mastery](https://excelmacromastery.com/vba-articles/)
  - [VBA Class Modules - The Ultimate Guide - Excel Macro Mastery](https://excelmacromastery.com/vba-class-modules/)
  - [VBA Class Modules - The Ultimate Guide - Excel Macro Mastery - Class Module Events (Factory Method)](https://excelmacromastery.com/vba-class-modules/#Class_Module_Events)
