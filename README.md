# nimi-vba
VBA7 Win64 wrapper for NI Modular Instrument C Drivers framework.

## nimi-vba C DLL Wrapping
The nimi-vba uses the simularities between NI Modular instrunts drivers driven by IVI to implement a consistent VBA implementation.
Each driver wrapper consists of two files a Module and a Module Class. The framework depends on the VBA Declare statement to reference functions in the C Dlls

The module file contains the public constants and Factory method for the Class. The class module implements a session class that manages the a driver session. when created the a session is opened and when the object instance goes out of scope the session is automatically clodes.

Methods on the clase module wrap API functions and Properties wrap driver Attributes.

### Factory Method
To Do

### Error Handling
To Do

### Methods
To Do

### Properties & Attributes
To Do

## Data Types
How to map datatypes in VBA Declare statements and additional conversion needs

| IVI / VISA Datatype | Base C Type         | Visual Basic Type   | Conversion needs    |
| ------------------- | ------------------- | ------------------- | ------------------- |
| ``ViUInt64`` | ``unsigned __int64`` | ``LongLong`` | |
| ``ViInt64`` | ``signed __int64`` | ``LongLong`` | |
| ``ViUInt32`` | ``unsigned long`` | ``Long`` | |
| ``ViInt32`` | ``signed long`` | ``Long`` | |
| ``ViUInt16`` | ``unsigned short`` | ``Int`` | |
| ``ViInt16`` | ``signed short `` | ``Int`` | |
| ``ViUInt8`` | ``unsigned char`` | ``Byte`` | |
| ``ViInt8`` | ``signed char `` | ``Byte`` | |
| ``ViByte`` | ``unsigned char`` | ``Byte`` | |
| ``ViChar`` | ``char `` | ``Byte`` | |
| ``ViReal32`` | ``float`` | ``Single`` | |
| ``ViReal64`` | ``double`` | ``Double`` | |
| ``ViBoolean`` | ``unsigned short`` | ``Boolean`` | |
| ``ViString`` | ``char * `` | ``LongPtr`` or ``ByVal String`` | See [Using Pointers and Strings](#Using-Pointers-and-Strings) | 
| ``ViConstString`` | ``const char * `` | ``ByVal String`` | See [Using the ByVal str As String declaration](#Using-the-ByVal-str-As-String-declaration) | 
| ``ViRsrc`` | ``char * `` | ``LongPtr`` or ``ByVal String`` | See [Using Pointers and Strings](#Using-Pointers-and-Strings) | 
| ``ViStatus`` | ``signed long`` | ``Long`` | |
| ``ViSession`` | ``unsigned long`` | ``Long`` | |
| ``niRFmxInstrHandle`` | ``void *`` | ``LongPtr`` | |


## Using Pointers and Strings
In the VBA Declare statement you need to use ``ByVal`` to indicate a argument is passed **by value** and ``ByRef`` to when passed **by reference**. This means that in most cases you just need to use ``ByRef`` when in externally a pointer is required.

So this C function declaraction 
```C
ViStatus _VI_FUNC niDMM_GetAttributeViReal64(ViSession vi, ViConstString channelName, ViAttr attributeId, ViReal64 *value);
```
Will translate into the following VBA Declare statement,
```VBA
Private Declare PtrSafe Function niDMM_GetAttributeViReal64 Lib "niDMM_64" ( _
    ByVal vi As Long, _
    ByVal channelName As String, _
    ByVal attributeID As Long, _
    ByRef value As Double _
) As Long
```
*Note how ByRef is used on the ``value`` argument.*

The main exception is when using arrays or other complex data structions. In this case you need to pass the pointer by value and dereference the pointer in VBA.
Argument declaration for a C pointer to an array or complext data structure in the VBA Declare statement:
```VBA
    ByVal value As LongPtr
```

In C arrays are also just pointers to the first element of the array. you can use the following VBA statement to get access to this pointer:
```VBA
VarPtr( myArrayVarable( 0 ) )
```
Or when you don't know the exact array bounds use the more generic:
```VBA
VarPtr( myArrayVarable( LBound( myArrayVarable ) ) )
```

### Strings
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
