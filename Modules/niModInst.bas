Attribute VB_Name = "niModInst"
Option Explicit

' niModInst_GetInstalledDeviceAttributeViString Attribute IDs
Public Const NIMODINST_ATTR_DEVICE_NAME As Long = 0
Public Const NIMODINST_ATTR_DEVICE_MODEL As Long = 1
Public Const NIMODINST_ATTR_SERIAL_NUMBER As Long = 2

' niModInst_GetInstalledDeviceAttributeViInt32 Attribute IDs
Public Const NIMODINST_ATTR_SLOT_NUMBER As Long = 10
Public Const NIMODINST_ATTR_CHASSIS_NUMBER As Long = 11
Public Const NIMODINST_ATTR_BUS_NUMBER As Long = 12
Public Const NIMODINST_ATTR_SOCKET_NUMBER As Long = 13
Public Const NIMODINST_ATTR_PCIEXPRESS_LINK_WIDTH As Long = 17
Public Const NIMODINST_ATTR_MAX_PCIEXPRESS_LINK_WIDTH As Long = 18

' niModInst Factory Method
Public Function niModInst_CreateSession(driver As String) As niModInst_Session
    Dim session As niModInst_Session
    
    Set session = New niModInst_Session
    session.InitSession (driver)
    
    Set niModInst_CreateSession = session
End Function

' Utility function to get device names, most common use case for inModInst
Public Function niModInst_GetDeviceNames(driver As String) As Collection
    Dim mi As niModInst_Session
    Dim devNames As Collection
    Dim name As String
    Dim index As Long
    
    Set mi = niModInst_CreateSession(driver)
    Set devNames = New Collection
      
    If mi.Count > 0 Then
        For index = 0 To mi.Count - 1
            mi.GetInstalledDeviceAttributeString index, NIMODINST_ATTR_DEVICE_NAME, name
            devNames.Add name
        Next index
    End If
    
    Set niModInst_GetDeviceNames = devNames
End Function
