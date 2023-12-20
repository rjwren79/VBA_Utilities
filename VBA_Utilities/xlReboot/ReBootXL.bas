Attribute VB_Name = "ReBootXL"
Option Explicit

'namespace=vba-files\xlReboot

Sub ReOpenXL()
    Dim strCMD As String
    strCMD = "CMD /C PING 127.0.0.1 -n 1 -w 5000 >NUL & Excel.exe " & Chr(34) & ThisWorkbook.FullName & Chr(34)
    
    ThisWorkbook.Save
    Shell strCMD, vbHide
    If Application.Workbooks.Count = 1 Then
        Application.Quit
    Else
        ThisWorkbook.Close SaveChanges:=False
    End If

End Sub
