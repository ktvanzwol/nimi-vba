Attribute VB_Name = "example_niModInst"
Option Explicit

Public Sub ListAllModularInstruments()
    Dim ws As Worksheet
    Dim mi As niModInst_Session
    Dim devNames As Collection
    Dim sAttr As String
    Dim lAttr As Long
    Dim index As Long
    
    ' List all instruments from niModInst supported drivers with empty driver string
    Set mi = niModInst_CreateSession("")
    Set ws = Example_GetNewOutputSheet("niModInst")
    
    ws.range("A1").Value2 = "Resource Name"
    ws.range("B1").Value2 = "Device Model"
    ws.range("C1").Value2 = "Serial Number"
    ws.range("D1").Value2 = "Chassis Number"
    ws.range("E1").Value2 = "Slot Number"
      
    If mi.Count > 0 Then
        For index = 0 To mi.Count - 1
            mi.GetInstalledDeviceAttributeString index, NIMODINST_ATTR_DEVICE_NAME, sAttr
            ws.Cells(index + 2, 1).Value2 = sAttr
            
            mi.GetInstalledDeviceAttributeString index, NIMODINST_ATTR_DEVICE_MODEL, sAttr
            ws.Cells(index + 2, 2).Value2 = sAttr
            
            mi.GetInstalledDeviceAttributeString index, NIMODINST_ATTR_SERIAL_NUMBER, sAttr
            ws.Cells(index + 2, 3).Value2 = sAttr
            
            mi.GetInstalledDeviceAttributeLong index, NIMODINST_ATTR_CHASSIS_NUMBER, lAttr
            ws.Cells(index + 2, 4).Value2 = lAttr
            
            mi.GetInstalledDeviceAttributeLong index, NIMODINST_ATTR_SLOT_NUMBER, lAttr
            ws.Cells(index + 2, 5).Value2 = lAttr
        Next index
    End If
    
    ws.columns("A").AutoFit
    ws.columns("B").AutoFit
    ws.columns("C").AutoFit
    ws.columns("D").AutoFit
    ws.columns("E").AutoFit
End Sub
