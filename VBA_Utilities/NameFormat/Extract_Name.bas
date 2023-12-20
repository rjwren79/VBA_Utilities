Attribute VB_Name = "Extract_Name"
' Add namespace for XVBA file structure
'namespace=vba-files\VBA_Utilities\NameFormat
'@Folder("VBA_Utilities\NameFormat")
Option Explicit
Private sNameSub As Variant, IsInitialized As Boolean, LastRow As Long
Private ws As Worksheet, waRange As Range, icRange As Range, oaRange As Range, ocRange As Range

Private Sub Initialize()
    
    pfOptimize.pfEnable
    
    Set ws = Sheet1
    
    With ws
    ' Clear existing data
        Debug.Print "Clear existing data."
        LastRow = GetLastRow("F")
        If Not LastRow < 2 Then
            Set oaRange = .Range("B2:F" & LastRow)
            oaRange.Clear
            Debug.Print "Cleared."
        Else
            Debug.Print "Nothing to clear."
        End If
        Debug.Print "Done."

    ' Set new parameters
        LastRow = GetLastRow("A")
        Set waRange = .Range("A2:F" & LastRow) 'Set working area
        Set icRange = .Range("A2:A" & LastRow) 'Set input column
        Set oaRange = .Range("B2:F" & LastRow) 'Set output area
        Set ocRange = .Range("F2:F" & LastRow) 'Set output column
    End With
    
    IsInitialized = True

End Sub

Private Function GetLastRow(ByVal column As String) As Long

    GetLastRow = ws.Range(column & ws.Rows.Count).End(xlUp).Row

End Function

Private Sub Terminate()

    Set waRange = Nothing
    Set icRange = Nothing
    Set oaRange = Nothing
    Set ocRange = Nothing
    LastRow = Empty
    IsInitialized = Empty
    
    

End Sub

Private Sub ExtractName()

Dim lRow As Long, i As Long, employee As clsEmployee, entry As String, SplitName, SplitArrLast As Integer, _
    hasSuffix As Boolean, hasXtraName As Boolean, SuffixNamePos As Integer, LastNamePos As Integer, _
    FirstNamePos As Integer, MiddleNamePos As Integer, j As Integer, a As Integer, eList As New Collection, _
    SuffixArr
    
    lRow = LastRow
    SuffixArr = "JR, SR, II, III, IV, V"
    SuffixArr = Split(SuffixArr, ", ")

    With ws
        For i = 2 To lRow
            Set employee = New clsEmployee
            entry = StrConv(Range("A" & i), vbProperCase)
            If InStr(1, entry, ",") > 0 Then
                sNameSub = Split(entry, ", ")
                Call ReverseList
                entry = Join(sNameSub, " ")
            End If
            SplitName = Split(entry, " ")
            SplitArrLast = UBound(SplitName)
            With employee
                .entry = entry
                a = 0
                For a = LBound(SuffixArr) To UBound(SuffixArr)
                    If InStr(1, " " & UCase(entry) & " ", " " & Trim(SuffixArr(a)) & " ") Then
                        SuffixNamePos = SplitArrLast
                        If Not InStr(1, SplitName(SuffixNamePos), ".") > 0 Then
                            SplitName(SuffixNamePos) = Replace(SplitName(SuffixNamePos), "R", "r.")
                            SplitName(SuffixNamePos) = Replace(SplitName(SuffixNamePos), "i", "I")
                            SplitName(SuffixNamePos) = Replace(SplitName(SuffixNamePos), "v", "V.")
                        End If
                        hasSuffix = True
                        Exit For
                    End If
                Next a
                    
                If hasSuffix Then
                    SuffixNamePos = SplitArrLast 'Add Suffix Position
                    .Suffix = SplitName(SuffixNamePos) 'Add Suffix
                    LastNamePos = SuffixNamePos - 1
                Else
                    LastNamePos = SplitArrLast
                End If
                
                '---------------------- Extra names go to middle name
                Dim xNameCount
                xNameCount = LastNamePos - 1
                If xNameCount > 0 Then
                    j = 0
                    For j = 1 To xNameCount 'Add extra names
                    If Not j < 2 Then
                        Dim joinMiddle As Variant
                        joinMiddle = Array(SplitName(1), SplitName(j))
                        SplitName(1) = Join(joinMiddle, " ")
                        SplitName(j) = ""
                        Erase joinMiddle
                    End If
                    Next j
                    .MiddleName = SplitName(1) 'Add extra names
                End If
                .SurName = SplitName(LastNamePos) 'Add Surname
                .FirstName = SplitName(0) 'Add First given
            End With

NextLoop:
            eList.Add employee
            hasSuffix = Empty
            hasXtraName = Empty
            Next i
            Call WriteNames(eList)
        End With

ExitProcdure:
    'On Error GoTo 0
    Exit Sub

End Sub

Private Sub WriteNames(ByVal Name As Collection)
    
    Dim i As Long, l As Long, employeeOut As clsEmployee
    Dim ProgressBar As New ProgressBar, progBar As Long
    
    For i = 1 To Name.Count
        progBar = WorksheetFunction.RoundUp((i * 100) / Name.Count, 0)
        'Debug.Print progBar
        l = 0
        Set employeeOut = Name(i)
        l = i + 1
        With Sheet1
            .Cells(l, "B").value = employeeOut.FirstName
            .Cells(l, "C").value = employeeOut.MiddleName
            .Cells(l, "D").value = employeeOut.SurName
            .Cells(l, "E").value = employeeOut.Suffix
            .Cells(l, "F").value = employeeOut.FileAs
        End With
        
        Call ProgressBar.Update(progBar, 100, "Writting Names", True)
        'Application.Wait (Now + TimeValue("0:00:01")) 'Remove before production
        
    Next i
    
End Sub

Private Sub ReverseList()

    Dim Arr As Variant
    Dim Count As Integer
    Dim a As Long
    Dim i As Long
    Dim j As Long
    
    Arr = sNameSub
    
    Count = UBound(Arr) - LBound(Arr) + 1

    If Count = 0 Then Exit Sub
    i = Count - 1
    j = i
    
    ReDim Preserve Arr(i)

    For a = 0 To i
        Arr(a) = sNameSub(j)
        j = j - 1
    Next a

    sNameSub = Arr

End Sub

Sub Raw_Ascending()
On Error GoTo eCtrl

    Dim LastRow As Long
    Dim rng As Range
    If Not IsInitialized Then GoTo eCtrl
    
sProcedure:
    Set rng = icRange
    rng.Sort key1:=rng, order1:=xlAscending, Header:=xlNo

eProcedure:
    On Error GoTo 0
    Exit Sub

eCtrl:
    If Not Err.Number = 0 Then Debug.Print Err.Number & " - " & Err.Description
    Err.Clear
    Call Start_List
    GoTo sProcedure
    
End Sub

Sub List_Ascending()
On Error GoTo eCtrl

    Dim LastRow As Long
    Dim rng As Range
    If Not IsInitialized Then GoTo eCtrl
    
sProcedure:
    Set rng = ocRange
    rng.Sort key1:=rng, order1:=xlAscending, Header:=xlNo

eProcedure:
    On Error GoTo 0
    Exit Sub

eCtrl:
    If Not Err.Number = 0 Then Debug.Print Err.Number & " - " & Err.Description
    Err.Clear
    Call Initialize
    GoTo sProcedure
    
End Sub

Sub Start_List()
    
    If Not IsInitialized Then Call Initialize
    Call ExtractName
    Call Terminate    
    
End Sub
