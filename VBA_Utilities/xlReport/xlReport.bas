Attribute VB_Name = "xlReport"
' Add namespace for XVBA file structure
'namespace=vba-files\VBA_Utilities\xlReport
'@Folder("VBA_Utilities\xlReport")
Option Explicit
Option Private Module
Public msgDocument As String
Public Const MESSAGE_TTL As String = "==================================== REPORT ===================================="
Public Const MESSAGE_TAB As String = "    "
Public Const MESSAGE_SPC As String = "                                                                                "
Public Const MESSAGE_BAR As String = "--------------------------------------------------------------------------------"
Public Const MESSAGE_END As String = "================================== END REPORT =================================="

Public Sub mPrint(message As String)

        Debug.Print Report(message)
    
End Sub

Public Sub mBox(message As String)

    Debug.Print Report(message)

End Sub

Public Sub mFile(message As String)

    Debug.Print Report(message)

End Sub

Public Sub Document(message As String)

    msgDocument = msgDocument & vbCrLf & message

End Sub

Private Function Report(message As String) As String

    Dim mText As String
    Dim gFor As String
    Dim TabInsert As Integer
    
    gFor = "report generated for " & GetUserName & MESSAGE_TAB
    TabInsert = GetLenghString(MESSAGE_SPC) - GetLenghString(gFor)
    
    mText = vbCrLf & MESSAGE_TTL
    mText = mText & vbCrLf
    mText = mText & vbCrLf & message
    mText = mText & vbCrLf
    mText = mText & vbCrLf & MESSAGE_BAR
    mText = mText & vbCrLf
    mText = mText & vbCrLf & Space(TabInsert) & gFor
    mText = mText & vbCrLf & MESSAGE_END

    Report = mText

End Function

Private Function GetLenghString(message As String) As Long

    GetLenghString = Len(message)

End Function

Private Function FindSpace(message As String)
' Find Space in message
    
End Function
