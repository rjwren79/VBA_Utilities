Attribute VB_Name = "ReportUtility"
' Add namespace for XVBA file structure
'namespace=vba-files\Utilities\Jarvis
'@Folder("vba-files\Utilities\Jarvis")
Option Explicit
'Option Private Module
Public Const MESSAGE_TITLE As String = "=========== REPORT ==========="
Public Const MESSAGE_TAB As String   = "    "
Public Const MESSAGE_SPACE As String = "                              "
Public Const MESSAGE_BAR As String   = "=============================="

Public Sub PrintToScreen(Message As String)

    Debug.Print Report(Message)

End Sub

Private Function Report(Message As String) As String

    Dim mText As String
    
    mText = vbCrLf & MESSAGE_TITLE
    mText = mText & vbCrLf & Message
    mText = mText & vbCrLf & MESSAGE_BAR
    mText = mText & vbCrLf

    Report = mText

End Function