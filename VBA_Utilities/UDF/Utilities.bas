Attribute VB_Name = "Utilities"
' Add namespace for XVBA file structure
'namespace=vba-files\VBA_Utilities\UDF
'@Folder("VBA_Utilities\UDF")
Option Explicit
'Option Private Module

Enum gUserName 'Enum for GetUserName
    from_System
    from_Application
End Enum

Function GetUserName(Optional from As gUserName = from_Application)

    Dim fnDescription As String
    fnDescription = "Get application or system username"

    Select Case from
        Case from_Application
            GetUserName = Application.UserName
        Case from_System
            GetUserName = Environ("username")
    End Select
            
End Function

Function Show_Window(sw As Boolean) As String
    
    Dim fnDescription As String
    fnDescription = "Show or hide window"
    
    Dim message As String
    ThisWorkbook.Activate
    ActiveWindow.Visible = sw
    message = "Done."
    GoTo ExitSub

ExitSub:
    Show_Window = message
    Exit Function

End Function

Function GetLastRow(ByVal ws As Worksheet, ByVal column As String) As Long

    Dim fnDescription As String
    fnDescription = "Get last row of column " & column & " of worksheet " & ws.Name

    GetLastRow = ws.Range(column & ws.Rows.Count).End(xlUp).Row

End Function

Function IsNullOrEmpty(s As String) As Boolean

    Dim fnDescription As String
    fnDescription = "Check if string is Null or Empty"
    
    IsNullOrEmpty = False

    If s = "" Or s = Empty Or s = Null Then IsNullOrEmpty = True

End Function

Function ISZERRO(s As Long) As Boolean

    Dim fnDescription As String
    fnDescription = "Check if value is 0."
    
    ISZERRO = False

    If s = 0 Then ISZERRO = True

End Function

Function DoesSheetExist(SheetName As String) As Boolean

    Dim fnDescription As String
    fnDescription = "Check if sheet '" & SheetName & "' exists"
    
    DoesSheetExist = Evaluate("ISREF('" & SheetName & "'!A1)")
    
End Function

Function CreateNewWorksheet(ByVal SheetName As String) As Worksheet

    Dim fnDescription As String
    fnDescription = "Check if sheet '" & SheetName & "' exists"

    Dim DSE As Boolean
    DSE = DoesSheetExist(SheetName)
    If DSE Then Exit Function

     'create new worksheet at the end of the workbook
    Set CreateNewWorksheet = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    CreateNewWorksheet.Name = SheetName 'set the worksheet name
    With ThisWorkbook.VBProject.VBComponents(CreateNewWorksheet.CodeName) 'access the worksheet codename
        .Properties("_CodeName") = SheetName 'set the worksheet codename
    End With

End Function

Function CollectionToArray(c As Collection) As Variant()
    Dim a() As Variant: ReDim a(0 To c.Count - 1)
    Dim i As Integer
    For i = 1 To c.Count
        a(i - 1) = c.Item(i)
    Next
    CollectionToArray = a
End Function

'This code will highlight the cells that have misspelled words
Sub HighlightMisspelledCells()
    Dim cl As Range
    For Each cl In ActiveSheet.UsedRange
        If Not Application.CheckSpelling(word:=cl.Text) Then
            cl.Interior.Color = vbRed
        End If
    Next cl
End Sub

Sub PrintArray(Data, SheetName, StartRow, StartCol)

    Dim Row As Integer
    Dim Col As Integer

    Row = StartRow

    For i = LBound(Data, 1) To UBound(Data, 1)
        Col = StartCol
        For j = LBound(Data, 2) To UBound(Data, 2)
            Sheets(SheetName).Cells(Row, Col).Value = Data(i, j)
            Col = Col + 1
        Next j
            Row = Row + 1
    Next i

End Sub

Sub PrintArray(Data)

    Dim Row As Integer
    Dim Col As Integer

    Row = StartRow

    For i = LBound(Data, 1) To UBound(Data, 1)
        Col = StartCol
        For j = LBound(Data, 2) To UBound(Data, 2)
            Debug.print Data(i, j)
            Col = Col + 1
        Next j
            Row = Row + 1
    Next i

End Sub
