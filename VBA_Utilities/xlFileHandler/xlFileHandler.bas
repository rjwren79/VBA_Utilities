Attribute VB_Name = "xlFileHandler"
' Add namespace for XVBA file structure
'namespace=vba-files\VBA_Utilities\UpdateModules

Option Explicit

Public Sub PrintFileDetails(FilePath As String, TargetFile As String)
    
    Dim sFile As String
    sFile = FilePath & TargetFile
    '~~> File Path
    Debug.Print "File Location: "; FilePath
    '~~> Target File
    Debug.Print "File Name: "; TargetFile
    '~~> Created Date
    Debug.Print "Created Date: "; Created(sFile)
    '~~> Modified Date
    Debug.Print "Modified Date: "; Modified(sFile)

End Sub

Private Function Created(FullTargetFile As String) As String
    
    Dim sFile As String
    sFile = FullTargetFile
    Dim oFS As Object
    Set oFS = CreateObject("Scripting.FileSystemObject")
    '~~> Created Date
    Created = oFS.GetFile(sFile).DateCreated
    Set oFS = Nothing
    
End Function

Private Function Modified(FullTargetFile As String) As String
    
    Dim sFile As String
    sFile = FullTargetFile
    Dim oFS As Object
    Set oFS = CreateObject("Scripting.FileSystemObject")
    '~~> Created Date
    Modified = oFS.GetFile(sFile).Datelastmodified
    Set oFS = Nothing
    
End Function
