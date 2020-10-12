# nimi-vba
VBA7 Win64 wrapper for NI Modular Instrument C Drivers framework.

## nimi-vba C DLL Wrapping
The nimi-vba uses the simularities between NI Modular instrunts drivers driven by IVI to implement a consistent VBA implementation.
Each driver wrapper consists of two files a Module (.bas) and a Module Class (.cls). The framework depends on the VBA Declare statement to reference functions in the C Dlls

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

## nimi-vba Data Types
How to map datatypes in VBA Declare statements and additional conversion needs

| IVI / VISA Datatype | Base C Type         | Visual Basic Type   | Conversion needs    |
| ------------------- | ------------------- | ------------------- | ------------------- |
| ``ViUInt64`` | ``unsigned __int64`` | ``ByVal LongLong`` | |
| ``ViInt64`` | ``signed __int64`` | ``ByVal LongLong`` | |
| ``ViUInt32`` | ``unsigned long`` | ``ByVal Long`` | |
| ``ViInt32`` | ``signed long`` | ``ByVal Long`` | |
| ``ViUInt16`` | ``unsigned short`` | ``ByVal Int`` | |
| ``ViInt16`` | ``signed short `` | ``ByVal Int`` | |
| ``ViUInt8`` | ``unsigned char`` | ``ByVal Byte`` | |
| ``ViInt8`` | ``signed char `` | ``ByVal Byte`` | |
| ``ViByte`` | ``unsigned char`` | ``ByVal Byte`` | |
| ``ViChar`` | ``char `` | ``ByVal Byte`` | |
| ``ViReal32`` | ``float`` | ``ByVal Single`` | |
| ``ViReal64`` | ``double`` | ``ByVal Double`` | |
| ``ViBoolean`` | ``unsigned short`` | ``ByVal Boolean`` | |
| ``ViString`` | ``char \* `` | ``ByRef Byte()`` | Requires convertion To/From UNICODE | 
| ``ViConstString`` | ``const char \* `` | ``ByRef String`` | UNICODE conversion seems to happen automatically in VBA7 | 
| ``ViStatus`` | ``signed long`` | ``Long`` | |
| ``ViSession`` | ``unsigned long`` | ``Long`` | |

### Pointers
to do

### Arrays
to do 

### Strings
to do

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
