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
            Dim moduleComp As String
            moduleComp = Right(file.Name, 4)
            If moduleComp = ".cls" Or moduleComp = ".bas" Or moduleComp = ".frm" Then
                ' Export the existing module to the TempMod folder
                Dim moduleName As String
                moduleName = Left(file.Name, Len(file.Name) - 4)
                If ModuleExists(moduleName) Then
'                    Dim moduleRenamed As String
'                    moduleRenamed = moduleName & "_old"
'                    file.Name = moduleRenamed & moduleComp
'                    ThisWorkbook.VBProject.VBComponents(moduleName).Name = moduleRenamed
'                    ThisWorkbook.VBProject.VBComponents(moduleRenamed).Export tempFolder.Path & "\" & file.Name
'                    ThisWorkbook.VBProject.VBComponents.Remove VBComponent:=ThisWorkbook.VBProject.VBComponents(moduleRenamed)
                    ExportVBComponent ThisWorkbook.VBProject.VBComponents(moduleName), tempFolder.Path, file.Name, True
                    ' Delete the existing after export
                    DeleteModule moduleName
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

Function GetFileExtension(VBComp As VBIDE.VBComponent) As String
    ' This returns the appropriate file extension based on the Type of the VBComponent.
        Select Case VBComp.Type
            Case vbext_ct_ClassModule
                GetFileExtension = ".cls"
            Case vbext_ct_Document
                GetFileExtension = ".cls"
            Case vbext_ct_MSForm
                GetFileExtension = ".frm"
            Case vbext_ct_StdModule
                GetFileExtension = ".bas"
            Case Else
                GetFileExtension = ".bas"
        End Select
        
End Function

Public Function ExportVBComponent(VBComp As VBIDE.VBComponent, _
                FolderName As String, _
                Optional FileName As String, _
                Optional OverwriteExisting As Boolean = True) As Boolean
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' This function exports the code module of a VBComponent to a text
    ' file. If FileName is missing, the code will be exported to
    ' a file with the same name as the VBComponent followed by the
    ' appropriate extension.
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim Extension As String
    Dim FName As String
    Extension = GetFileExtension(VBComp:=VBComp)
    If Trim(FileName) = vbNullString Then
        FName = VBComp.Name & Extension
    Else
        FName = FileName
        If InStr(1, FName, ".", vbBinaryCompare) = 0 Then
            FName = FName & Extension
        End If
    End If
    
    If StrComp(Right(FolderName, 1), "\", vbBinaryCompare) = 0 Then
        FName = FolderName & FName
    Else
        FName = FolderName & "\" & FName
    End If
    
    If Dir(FName, vbNormal + vbHidden + vbSystem) <> vbNullString Then
        If OverwriteExisting = True Then
            Kill FName
        Else
            ExportVBComponent = False
            Exit Function
        End If
    End If
    
    VBComp.Export FileName:=FName
    ExportVBComponent = True
    
End Function

Sub DeleteModule(moduleName As String)
    Dim VBProj As VBIDE.VBProject
    Dim VBComp As VBIDE.VBComponent

    Set VBProj = ActiveWorkbook.VBProject
    Set VBComp = VBProj.VBComponents(moduleName)
    VBProj.VBComponents.Remove VBComp
End Sub

Sub RenameModule(currentName As String, newName As String)
    ThisWorkbook.VBProject.VBComponents(currentName).Name = newName
End Sub

Sub reNameThisWorkbook()
    RenameModule "ThisWorkbook_old", "ThisWorkbook"
End Sub

Sub reNameSheet1()
    RenameModule "Sheet1_old", "Sheet1"
End Sub
