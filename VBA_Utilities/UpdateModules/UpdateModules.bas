Attribute VB_Name = "UpdateModules"
' Add namespace for XVBA file structure
'namespace=vba-files\VBA_Utilities\UpdateModules
'@Folder("VBA_Utilities\UpdateModules")
Option Explicit

Sub ImportModules()
    ' Create FileSystemObject
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Get the current path of the workbook
    Dim currentPath As String
    currentPath = ThisWorkbook.Path
    
    ' Get the import path of the workbook
    Dim importPath As String
    importPath = currentPath & "\vba-files\VBA_Utilities"
    
    ' Create a folder named "TempMod" in the current path
    Dim tempFolder As Object
    If Not fso.FolderExists(currentPath & "\TempMod") Then
        Set tempFolder = fso.CreateFolder(currentPath & "\TempMod")
    Else
        Set tempFolder = fso.GetFolder(currentPath & "\TempMod")
    End If
    
    ' Loop through all subfolders in the import path
    Dim subfolder As Object
    For Each subfolder In fso.GetFolder(importPath).SubFolders
        ' Loop through all files in the subfolder
        Dim file As Object
        For Each file In subfolder.Files
            ' Check if the file is a component type
            If Right(file.Name, 4) = ".cls" Or Right(file.Name, 4) = ".bas" Or Right(file.Name, 4) = ".frm" Then
                ' Export the existing module to the TempMod folder
                Dim moduleName As String
                moduleName = Left(file.Name, Len(file.Name) - 4)
                If ModuleExists(moduleName) Then
                    ThisWorkbook.VBProject.VBComponents(moduleName).Export tempFolder.Path & "\" & file.Name
                    ThisWorkbook.VBProject.VBComponents.Remove VBComponent:=ThisWorkbook.VBProject.VBComponents(moduleName)
                End If
                
                ' Import the module from the file
                ThisWorkbook.VBProject.VBComponents.Import file.Path
            End If
        Next file
    Next subfolder
End Sub

Function ModuleExists(moduleName As String) As Boolean
    ' Check if a module exists
    Dim component As Object
    For Each component In ThisWorkbook.VBProject.VBComponents
        If component.Name = moduleName Then
            ModuleExists = True
            Exit Function
        End If
    Next component
    ModuleExists = False
End Function