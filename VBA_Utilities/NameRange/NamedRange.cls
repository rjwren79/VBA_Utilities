VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "NamedRange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'namespace=VBA_Utilities\NamedRange

' A class to add named ranges in VBA
' The class name is NamedRange
' The class has two properties: Name and Range
' The class has one method: Add_NameToRange

' Declare the properties of the class
Private pName As String
Private pRange As Range

' Create a property procedure for Name
Public Property Get Name() As String
    Name = pName
End Property

Public Property Let Name(ByVal value As String)
    pName = value
End Property

' Create a property procedure for Range
Public Property Get Range() As Range
    Set Range = pRange
End Property

Public Property Set Range(ByVal value As Range)
    Set pRange = value
End Property

' Create a method to add a name to range
Public Sub Add_NameToRange()
    ' Check if the name and range are valid
    If pName <> "" And Not pRange Is Nothing Then
        ' Add the name to the workbook's names collection
        ThisWorkbook.Names.Add Name:=pName, RefersTo:=pRange
        ' Display a message to confirm the operation
        'MsgBox "Named range " & pName & " has been set to " & pRange.Address, vbInformation, "Success"
    Else
        ' Display a message to indicate the error
        MsgBox "Please provide a valid name and range", vbCritical, "Error"
    End If
End Sub
