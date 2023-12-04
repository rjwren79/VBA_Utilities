Attribute VB_Name = "NamedRanges"
Option Explicit
'Option Private Module
'namespace=vba-files\module

Private s_name As String
Private s_address As String

Public Sub RangeName_New()
    
    Dim rngAddress As String
    
    Dim rng As Range
    Set rng = Application.Selection

    Dim text As String
    text = "Please enter name for range." & vbCrLf
    text = text & Chr(149) & " Must start with a letter or underscore." & vbCrLf
    text = text & Chr(149) & " Certain characters aren't allowed." & vbCrLf
    text = text & Chr(149) & " Must be unique." & vbCrLf
    text = text & Chr(149) & " Name cannot be a cell reference." & vbCrLf
    
    Dim rngName As String
    rngName = InputBox(text, "RangeName", "RangeName")    
    'Did user cancel
    If vbCancel =  True then Exit Sub
    'Need to add a test for valid NameRange



    'Set group name
    Call NameRange_Add(rngName, rng) 
    
    'Set name for each cell
    Dim answer As Integer
    answer = MsgBox("Do you want to name each cell in range?", vbQuestion + vbYesNo + vbDefaultButton2, "Group Range")

    If answer = vbYes then Call NameCells(rngName, rng)
    'Set cell name
    Exit Sub
End Sub

Public Sub RangeName_DeleteAll()

    Call NamedRanges_Delete_All()

End Sub

Public Sub NamedRange()
    ' 'Clear the previous entries
    ' Call NamedRanges_Delete_All
    ' 'Set developer columns
    ' s_name = "Developer"
    ' s_address = "A:B"
    ' Call s_set
    ' 'Set Operating Keys
    ' s_name = "oKeys"
    ' s_address = "C5:L10"
    ' Call s_set
    ' 'Set Bottom Pins
    ' s_name = "bPins"
    ' s_address = "C12:L14"
    ' Call s_set
    ' 'Set Master Split
    ' s_name = "mSplit"
    ' s_address = "C16:L21"
    ' Call s_set
    ' 'Set Control
    ' s_name = "Control"
    ' s_address = "N5:W11"
    ' Call s_set
    ' 'Set Top / Driver
    ' s_name = "tdPins"
    ' s_address = "N13:W17"
    ' Call s_set
    ' 'Set Core Pinning Matrix
    ' s_name = "pMatrix"
    ' s_address = "Y5:AH13"
    ' Call s_set
    ' 'Set Validate Stack Height
    ' s_name = "pValidate"
    ' s_address = "Y15:AH17"
    ' Call s_set
    
End Sub

Private Sub NameRange_Add(Name As String, Selection As Range)
    Dim nr As New NamedRange
    nr.Name = Name
    Set nr.Range = Range(Selection.address)
    nr.Add_NameToRange
End Sub

Private Sub NamedRanges_Delete_All()
    Dim MyName As Name
    For Each MyName In Names
    ActiveWorkbook.Names(MyName.Name).Delete
    Next
End Sub

Private Sub NameCells(Name As String, Selection As Range)
    Dim i As Long
    i = 0
    Dim rng As Range
    Set rng = Selection
    Dim cell As Range
    For Each cell In rng
        i = i + 1
        Dim CellName As String
        CellName = Name & Chr(46) & i
        Call NameRange_Add(CellName, cell)
    Next cell
End Sub