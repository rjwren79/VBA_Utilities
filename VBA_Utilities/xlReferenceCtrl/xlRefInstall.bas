Attribute VB_Name = "xlRefInstall"
Option Explicit

Public Sub InstallReferences()

    Dim ThisWB As Workbook ' The current workbook
    Set ThisWB = ThisWorkbook
    
    AddRef ThisWB, "{420B2830-E718-11CF-893D-00A0C9054228}", "Scripting", 1, 0
'    AddRef ThisWB, "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}", "MSScriptControl", 1, 0
    AddRef ThisWB, "{F5078F18-C551-11D3-89B9-0000F81FE221}", "MSXML2", 6, 0
    AddRef ThisWB, "{0002E157-0000-0000-C000-000000000046}", "VBIDE", 5, 3
    
End Sub
