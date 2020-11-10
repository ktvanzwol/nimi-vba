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

![](https://github.com/ktvanzwol/nimi-vba/raw/master/doc/Excel%20VBA%20Project%20Option.png)

Once this is done open the **nimi-vba ExcelTool.xlsm** file then on the nimi-vba sheet set the **Target Workbook** to *\<Create new Workbook\>* and click the Import/Update button on the sheet. This will create a new sheet and import all the nimi-vba modules into it.

![](https://github.com/ktvanzwol/nimi-vba/raw/master/doc/nimi-vba%20Excel%20Tool.png)

Alternatively you can open an existing workbook first and then select this as the Target Workbook to Import/Update nimi-vba in an existing application.

> :warning: **After confirming the import any existing module with a matching name will be overwriten without notification.**

### Exporting

If you fixed issues and or extended nimi-vba you can export modules manually or use the same **nimi-vba ExcelTool.xlsm** file to automatically export modules back into the **src** folder structure. To automatically export the modules open **nimi-vba ExcelTool.xlsm** and the excel application with the updates. Next Select the excel application workbook with the updates as the **Target Workbook** and click the Export button.

> :warning: **After confirming the export any existing files in the src folder structure will be overwriten without notification.**

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

In general when passing C arguments By Value you need to specify arguments with ``ByVal`` in the VBA Declare statement. By default the VBA Declare statement will assume passing the the C argument by reference (e.g. using a pointer). However it is good practise to specify ``ByRef`` in the VBA declare statement in this case.

Example of a C Function followed by the corresponding VBA Declare statement.

```C
ViStatus _VI_FUNC niDMM_GetAttributeViReal64(
    ViSession vi,
    ViConstString channelName,
    ViAttr attributeId,
    ViReal64 *value);
```

```VBA
Private Declare PtrSafe Function niDMM_GetAttributeViReal64 Lib "niDMM_64" ( _
    ByVal vi As Long, _
    ByVal channelName As String, _
    ByVal attributeID As Long, _
    ByRef value As Double _
) As Long
```

*Note how ``ByRef`` is used on the ``value`` argument.*

There are a few special cases that requires handling the pointer values manually to correctly pass data. These are discussed in the next section

### Array Pointers

Arrays in VBA are treated semantically like value types but are implemented as reference types.Internally VBA stores Array types as a SAFEARRAY structure, as a result passing a array type ``ByRef`` in the declare statement will not work. C will expect the pointer to the start of the first array element which is reference in the SAFEARRAY.

We can actually get to this pointer value by using the ``VarPtr()`` function on the first element in the VBA Array. ``VarPtr()`` will return the pointer value as a ``LongPtr`` type, the ``LongPtr`` type is a 32 bits ``Long`` in 32 bits Office and a 64 bits ``LongLong`` in 64 bits Office.

We can also use the ``LongPtr`` with the VBA Declare statement to directly pass pointers by value by defining a C pointer argument as ``ByVal value As LongPtr`` in the VBA Declare statement. We can use this to pass VBA Arrays to C external code.

Here is an example of a C Function that contains array pointers (``voltageMeasurements[]`` and ``currentMeasurements[]``) and the corresponding VBA Declare statement using the ``LongPtr`` type.

```C
ViStatus niDCPower_MeasureMultiple(
    ViSession vi,
    ViConstString channelName,
    ViReal64 voltageMeasurements[],
    ViReal64 currentMeasurements[]);
```

```VBA
Private Declare PtrSafe Function niDCPower_MeasureMultiple Lib "niDCPower_64" ( _
    ByVal vi As Long, _
    ByVal channelName As String, _
    ByVal voltageMeasurements As LongPtr, _
    ByVal currentMeasurements As LongPtr _
) As Long
```

We can now pass the VBA array to C by using the ``VarPtr()`` funtion on the arrays first ellement to get the correct pointer value to pass to use in the C function call. Like in this example:

```VBA
Dim voltageMeasurements() As Double
Dim currentMeasurements() As Double
Dim numMeasurements As Long
Dim status As Long

numMeasurements = 4
ReDim voltageMeasurements(numMeasurements) As Double
ReDim currentMeasurements(numMeasurements) As Double

status = niDCPower_MeasureMultiple(m_Session, "", _
                VarPtr( voltageMeasurements(0) ), VarPtr( currentMeasurements(0) ))
```

### Strings

Similar to arrays, strings in VBA are treated semantically like value types but are implemented as reference types. VBA Strings are represented a BSTR's which are multi byte UNICODE encodes strings preceded by the string length. In C strings are represented as byte arrays using ASCII characters with a null terminating character. This means strings need to be converted when passed between VBA and C.

The VBA Declare statement is able to handle some of this conversion for us. In the case of passing static strings to C external code you can simple define the argument as ``ByVal String``. In this case the conversion to null terminated ASCII string is done automatically.

In nimi-vba this is can be used for the majority of use cases. The notible exceptions are reciving error messanege and reading string type attributes. More generically speaking these are the cases were you need to read a dynamic size string typically you need to call the function twich. Once with NULL pointer to retrive the size of the string followed by a call with the properly sized string.

In this situation you need to treat the string as a ``Byte`` array. This allows you to pass 0 as the pointer value to query the requires for the size of the ``Byte`` array. The second call would get the string as a byte array en then needs to be converted to a native VBA string using the ``StrConv`` function.

An example of this using the ``niDMM_GetError`` function's ``errMessage`` argument:

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
    ReDim buffer(size - 1) As Byte 'In VBA the upperbound is inclusive.

    status = niDMM_GetError(m_Session, errorCode, size, VarPtr(buffer(0)))
    'Remove \0 character with LeftB and convert to Unicode with StrConv
    errorMsg = StrConv(LeftB(buffer(), size - 1), vbUnicode)
End Sub
```

Note that the ``StrConv`` function can also be used to convert from Unicode strings to C Style ASCII strings. E.g. a ``Byte`` array:

```VBA
buffer() = StrConv("StackOverflow", vbFromUnicode)
```

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
