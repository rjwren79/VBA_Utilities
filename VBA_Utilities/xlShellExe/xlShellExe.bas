Attribute VB_Name = "xlShellExe"
Option Explicit
' Add namespace for XVBA file structure
'namespace=vba-files\VBA_Utilities\xlShellExe
'@Folder("VBA_Utilities\xlShellExe")
Public Sub Restart_xl()
    Dim batFile As String
    batFile = "Restart.bat"
    Dim batFilePath As String
    batFilePath = ThisWorkbook.Path & "\batfiles\" & batFile
    CreateBatFile batFilePath, ThisWorkbook.FullName
    ShellExe batFilePath
    ThisWorkbook.Close True
End Sub

Private Function CreateBatFile(batFilePath As String, TagetFile As String)

    Dim iFileNum As Long
    iFileNum = FreeFile
    Dim exeFile As String
    exeFile = "Excel.exe"
    
    Open batFilePath For Output As #iFileNum
    Print #iFileNum, "@Echo off"
    Print #iFileNum, Replace("SET ~TagetFile=" & TagetFile & "~", "~", Chr(34), 1)
    Print #iFileNum, Replace("SET ~exeFile=" & exeFile & "~", "~", Chr(34), 1)
    Print #iFileNum, "TIMEOUT 3"
    Print #iFileNum, Replace("START /I ~%exeFile%~ ~%TagetFile%~", "~", Chr(34), 1)
    Close #iFileNum

End Function

Private Function ShellExe(batFilePath As String) As Boolean
    ShellExe = Shell("cmd /c " & Chr(34) & batFilePath & Chr(34), vbHide)
End Function
