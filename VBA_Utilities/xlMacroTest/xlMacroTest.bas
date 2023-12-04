Attribute VB_Name = "xlMacroTest"
' Add namespace for XVBA file structure
'namespace=VBA_Utilities\xlMacroTest
'@Folder("VBA_Utilities\xlMacroTest")
Option Explicit
'Option Private Module

Sub TestModule(tMacro As String)
    On Error GoTo ErrCtrl
    
        Dim cmd As String
        Dim tReturn As String
        Dim Report As String
        Dim errNum As Long
        Dim errDes As String
        cmd = tMacro
        Report = vbNullString
        
        Debug.Print "Starting Test"
        Debug.Print "Running " & cmd
        
Action:
        tReturn = Application.Run(cmd)
        Report = "None"
        GoTo ExitSub
    
ExitSub:
        Debug.Print "=========================="
        Debug.Print "       Test Report        "
        Debug.Print "Run Result: " & tReturn
        Debug.Print "Encountered Errors: "
        Debug.Print Report
        Debug.Print "=========================="
        Debug.Print "Test Complete."
        Exit Sub
    
ErrCtrl:
        errNum = Err.Number
        errDes = Err.Description
        Err.Clear
        Report = "Description: " & errDes & vbCrLf & _
                 "Error: " & errNum
        GoTo ExitSub
        
    End Sub

