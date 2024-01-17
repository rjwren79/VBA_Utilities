Attribute VB_Name = "ReferenceCTRL"
Option Explicit
Option Private Module ' Prevent the module from being visible in the macro dialog box

'namespace=vba-files\VBA_Utilities\UpdateModules

Sub AddRef(CurrentWB As Workbook, sGuid As String, sRefName As String, Optional varMajor As Variant, Optional varMinor As Variant)
    Dim i As Integer
    On Error GoTo EH
    With CurrentWB.VBProject.References
        If IsMissing(varMajor) Or IsMissing(varMinor) Then
           For i = 1 To .Count
               If .Item(i).Name = sRefName Then
                  Exit For
               End If
           Next i
           If i > .Count Then
              .AddFromGuid sGuid, 0, 0 ' 0,0 should pick the latest version installed on the computer
           End If
        Else
           For i = 1 To .Count
               If .Item(i).guid = sGuid Then
                  If .Item(i).major = varMajor And .Item(i).minor = varMinor Then
                     Exit For
                  Else
                     If vbYes = MsgBox(.Item(i).Name & " v. " & .Item(i).major & "." & .Item(i).minor & " is currently installed," & vbCrLf & "do you want to replace it with v. " & varMajor & "." & varMinor, vbQuestion + vbYesNo, "Reference already exists") Then
                        DelRef CurrentWB, sGuid
                     Else
                        i = 0
                        Exit For
                     End If
                  End If
               End If
           Next i
           If i > .Count Then
              .AddFromGuid sGuid, varMajor, varMinor
           End If
        End If
    End With
EX: Exit Sub
EH: MsgBox "Error in 'AddRef' for guid:" & sGuid & " " & vbCrLf & vbCrLf & Err.description
    Resume EX
    Resume ' debug code
End Sub

Public Sub DelRef(wbk As Workbook, sGuid As String)
    Dim oRef As Object
    For Each oRef In wbk.VBProject.References
        If oRef.guid = sGuid Then
           Debug.Print "The reference to " & oRef.FullPath & " was removed."
           Call wbk.VBProject.References.Remove(oRef)
        End If
    Next
End Sub

Public Sub DebugPrintExistingRefsWithVersion()
    Dim wbk As Workbook
    Dim i As Integer
    Set wbk = Application.ThisWorkbook
    With wbk.VBProject.References
        For i = 1 To .Count
            Debug.Print "   'AddRef wbk, """ & .Item(i).guid & """, """ & .Item(i).Name & """" & Space(30 - Len(vbNullString & .Item(i).Name)) & " ' install the latest version"
            Debug.Print "    AddRef """ & wbk.Name & """, """ & .Item(i).guid & """, """ & .Item(i).Name & """, " & .Item(i).major & ", " & .Item(i).minor & Space(30 - Len(", " & .Item(i).major & ", " & .Item(i).minor) - Len(vbNullString & .Item(i).Name)) & " ' install v. " & .Item(i).major & "." & .Item(i).minor
        Next i
    End With
End Sub


Sub ListReferencePaths()
 'Macro purpose:  To determine full path and Globally Unique Identifier (GUID)
 'to each referenced library.  Select the reference in the Tools\References
 'window, then run this code to get the information on the reference's library

On Error Resume Next
Dim i As Long

Debug.Print "Reference name" & " | " & "Full path to reference" & " | " & "Reference GUID" & " | " & "Major" & " | " & "Minor"

For i = 1 To ThisWorkbook.VBProject.References.Count
  With ThisWorkbook.VBProject.References(i)
    Debug.Print .Name & " | " & .FullPath & " | " & .guid & " | " & .major & " | " & .minor
  End With
Next i
On Error GoTo 0
End Sub


Sub AddReference()
    'Macro purpose:  To add a reference to the project using the GUID for the
    'reference library
     
    Dim strGUID As String, theRef As Variant, i As Long
     
    'Update the GUID you need below.
    strGUID = "{B691E011-1797-432E-907A-4D8C69339129}"
     
    'Set to continue in case of error
    On Error Resume Next
     
    'Remove any missing references
    For i = ThisWorkbook.VBProject.References.Count To 1 Step -1
        Set theRef = ThisWorkbook.VBProject.References.Item(i)
        If theRef.isbroken = True Then
            ThisWorkbook.VBProject.References.Remove theRef
        End If
    Next i
     
    'Clear any errors so that error trapping for GUID additions can be evaluated
    Err.Clear
     
    'Add the reference
    ThisWorkbook.VBProject.References.AddFromGuid _
    guid:=strGUID, major:=1, minor:=0
     
    'If an error was encountered, inform the user
    Select Case Err.Number
    Case Is = 32813
        'Reference already in use.
        Debug.Print "Reference already in use.  No action necessary."
    Case Is = vbNullString
         'Reference added without issue
         Debug.Print "Reference added without issue."
    Case Else
         'An unknown error was encountered, so alert the user
         Debug.Print "An unknown error was encountered."
        MsgBox "A problem was encountered trying to" & vbNewLine _
        & "add or remove a reference in this file" & vbNewLine & "Please check the " _
        & "references in your VBA project!", vbCritical + vbOKOnly, "Error!"
    End Select
    On Error GoTo 0
End Sub


