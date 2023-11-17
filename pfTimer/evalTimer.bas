Attribute VB_Name = "evalTimer"
' Add namespace for XVBA file structure
'namespace=vba-files\pfTimer
'@Folder("vba-files\pfTimer")
Option Explicit
Option Private Module

 Private Type initTest
     Title As String
     ProcessInit As String
     ProcessTerm As String
     Count As Long
 End Type

' Private Type Calculate
'     Max As Double
'     Min As Double
'     Avg As Double
' End Type

Private IsInitialized As Boolean, modTime As pfTimer, modTest As pfTest, _
        ProgressBar As ProgressBar, eResults As New Collection
Private nTest As initTest

Private Sub Initialize() ' Prepare evalTimer
    
    Debug.Print "Starting Test"
    pfOpti.Enable
    Set modTest = New pfTest
    
    With modTest
        .Title = "Test01"
    '     .ProcessInit = "pfOpti.Enable"
    '     .ProcessTerm = "pfOpti.Disable"
        .Count = 2
    End With

ExitSub:
    IsInitialized = True
    Exit Sub

End Sub

Private Sub Terminate() ' Quit evalTimer

    Set eResults = Nothing
    IsInitialized = Empty

ExitSub:
    pfOpti.Disable
    Debug.Print "Test Complete"
    Exit Sub

End Sub

Private Function Eval() As Double ' Run

    ' Start Timer
    With modTime
        .Clock_Start
        'Debug.Print "Starting Timer"
    End With

    ' Run test code
    With modTest
        If IsNullOrEmpty(.ProcessInit) Then GoTo ExitSub '
        Application.Run .ProcessInit ' Initialize procedure
        If IsNullOrEmpty(.ProcessTerm) Then GoTo ExitSub '
        Application.Run .ProcessTerm ' Terminate procedure
    End With

ExitSub:
    ' Stop timer
    With modTime
        Eval = .TimeElapsed
        Debug.Print "Stopping Timer"
        Debug.Print "Time Elapsed: " & Eval
    End With
    Exit Function

End Function

Private Sub LoopCode() ' Loop

    Debug.Print "Loop Code"
    'Set loop iterations
    Dim LoopCount As Long, iLoop
    LoopCount = nTest.Count
    'Start ProgressBar
    Dim progBar As Long, i As Long
    Set ProgressBar = New ProgressBar
    For i = 1 To LoopCount
        progBar = WorksheetFunction.RoundUp((i * 100) / LoopCount, 0)
        Set modTime = New pfTimer
        eResults.Add Eval
        Call ProgressBar.Update(progBar, 100, "Running Test", True) 'Call ProgressBar.Update
    Next i
    
ExitSub:
    Application.StatusBar = False
    Exit Sub

End Sub

Private Sub Get_Results() ' Get results
    
    Debug.Print "Get results"
    
    Dim i As Long
    Dim Item
    ReDim TimeArr(nTest.Count - 1)
        
    Set modTest = New pfTest
    
    For Each Item In eResults
        TimeArr(i) = CDbl(Item)
        i = i + 1
    Next Item
    
    With modTest
        .Title = nTest.Title
        .ProcessInit = nTest.ProcessInit
        .ProcessTerm = nTest.ProcessTerm
        .Count = nTest.Count
        .Data = TimeArr
     End With

ExitSub:
    Exit Sub

End Sub

Private Sub Write_Results() ' Write results

    Debug.Print "Write Results"
    Dim pWrite As String
    With modTest
        pWrite = "Name of Test: " & .Title
        pWrite = pWrite & vbCrLf & "No. of Tests: " & .Count
        pWrite = pWrite & vbCrLf & "Highest Time: " & .Max
        pWrite = pWrite & vbCrLf & "Lowest  Time: " & .Min
        pWrite = pWrite & vbCrLf & "Average Time: " & .Avg
    End With

    xlReport.mPrint pWrite

'    With pfResults
'        c = 2
'        .Cells(c, "A").value = .Name
'        .Cells(c, "B").value = eResults.Description
'        .Cells(c, "C").value = .Count
'        .Cells(c, "D").value = .Low
'        .Cells(c, "E").value = .High
'        .Cells(c, "F").value = .Avg
'    End With

End Sub

Private Sub SheetFormat(ByVal SheetName As Worksheet)
    
    Debug.Print "Format new sheet"
    With SheetName
    ' Disable page breaks
        .DisplayPageBreaks = False
    ' Format cells
        With .Cells
            .Clear
            .Clear
            .RowHeight = "15"
            .ColumnWidth = "8.43"
            With .Font
                .Name = "Calibri"
                .Size = "11"
                .Bold = False
                .Color = vbBlack
            End With
        End With
    End With
    
End Sub

Private Sub AddResultsSheet(ByVal SheetName As String)

    Debug.Print "Add sheet for results"
    Dim DSE As Boolean
    DSE = DoesSheetExist(SheetName)
    If Not DSE Then CreateNewWorksheet (SheetName)
    If Not DSE Then CreateNewWorksheet (SheetName)
    
End Sub

Sub Test()

    Initialize 'Call Initialize
    LoopCode 'Call LoopCode
    Get_Results 'Call Get_Results
    AddResultsSheet ("pfResults") 'Call AddResultsSheet
    AddResultsSheet ("pfResults") 'Call AddResultsSheet
    Write_Results 'Call Write_Results
    Terminate 'Call Terminate
    
End Sub