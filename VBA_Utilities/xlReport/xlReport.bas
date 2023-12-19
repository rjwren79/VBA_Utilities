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

Public Sub mPrint(text As String)

        Debug.Print Report(text)
    
End Sub

Public Sub mBox(text As String)

    Debug.Print Report(text)

End Sub

Public Sub mFile(text As String)

    Dim fileName As String
    fileName = "Report" & Format(Date, "MMddyyyy") & Format(Time, "HHmmss") & ".txt"
    Dim overwrite As Boolean
    overwrite = True

    Call WriteToFile(Report(text), fileName, overwrite)

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

'write to a file
Private Sub WriteToFile(text As String, fileName As String, Optional overwrite As Boolean = False)
    Dim filePath As String
    Dim fileNo As Integer
    Dim data As String
    
    ' Get the current path of the project
    filePath = ThisWorkbook.Path & "\" & fileName
    
    ' Check if the file exists
    If FileExists(filePath) Then
        ' If overwrite is True, clear the content of the file
        If overwrite Then
            Open filePath For Output As #1
            Print #1, text
            Close #1
        Else
            ' If overwrite is False, append the text to the last line of the file
            Open filePath For Append As #1
            Print #1, text
            Close #1
        End If
    Else
        ' If the file doesn't exist, create it and write the text
        Open filePath For Output As #1
        Print #1, text
        Close #1
    End If
End Sub

' Function to check if a file exists
Function FileExists(filePath As String) As Boolean
    FileExists = Dir(filePath) <> ""
End Function