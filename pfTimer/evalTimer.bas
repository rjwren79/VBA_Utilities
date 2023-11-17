Attribute VB_Name = "evalTimer"
' Add namespace for XVBA file structure
'namespace=vba-files\Modules
'@Folder("vba-files\Modules")
Option Explicit
Option Private Module

' Private Type Setup
'     Title As String
'     Proc_Init As String
'     Proc_Term As String
'     iCount As Long
' End Type

' Private Type Calculate
'     Max As Double
'     Min As Double
'     Avg As Double
' End Type

Private IsInitialized As Boolean, modTime As pfTimer, modTe
        ProgressBar As ProgressBar, eResults As New Collect

Private Sub Initialize() ' Prepare evalTimer
    
    Debug.Print "Starting Test"
    Optimization.Enable
    Set modTest = New pfTest
    
    With modTest
        .Title = "Test01"
    '     .ProcessInit = "Optimization.Enable"
    '     .ProcessTerm = "Optimization.Disable"
        .Count = 2
    End With

ExitSub:
    IsInitialized = True
    Exit Sub

End Sub

Private Sub Terminate() ' Quit evalTimer

    ' Set eResults = Nothing
    IsInitialized = Empty

ExitSub:
    Optimization.Disable
    Debug.Print "Test Complete"
    Exit Sub

End Sub

Private Function Eval() As Double ' Run

    ' Start Timer
    With modTime 
        .Clock_Start
        Debug.Print "Starting Timer"
    End With

    ' Run test code
    With modTest
        If IsNullOrEmpty(.ProcessInit) Then GoTo ExitSub ' 
        Application.Run .ProcessInit ' Initialize procdure
        If IsNullOrEmpty(.ProcessTerm) Then GoTo ExitSub ' 
        Application.Run .ProcessTerm ' Terminate procdure
    End With

ExitSub:
    ' Stop timer
    With modTime
        Eval = .TimeElapsed
        Debug.Print "Stoping Timer"
        Debug.Print "Time Elapsed: " & Eval
    End With
    Exit Function

End Function

Private Sub LoopCode() ' Loop

    Debug.Print "Loop Code"
    'Set loop iterations
    Dim LoopCount As Long, iLoop
    LoopCount = modTest.Count
    'Start ProgressBar
    Dim progBar As Long, i As Long
    Set ProgressBar = New ProgressBar
    For i = 1 To LoopCount
        progBar = WorksheetFunction.RoundUp((i * 100) / Loo
        Set modTime = New pfTimer
        eResults.Add Eval
        Call ProgressBar.Update(progBar, 100, "Running Test
    Next i
    
ExitSub:
    Application.StatusBar = False
    Exit Sub

End Sub
' Private Function Calc(Arr As Variant) As Calculate ' Calc

'     With WorksheetFunction
'         Calc.Max = .Max(Arr)
'         Calc.Min = .Min(Arr)
'         Calc.Avg = .Average(Arr)
'     End With

' End Function

Private Sub Get_Results() ' Get results
    
    Debug.Print "Get results"
    ' Dim i As Long, r As Calculate
    ' Dim Item
    ' ReDim TimeArr(init.iCount - 1)
    ' i = 0
    ' For Each Item In eResults
    '     TimeArr(i) = CDbl(Item)
    '     i = i + 1
    ' Next Item
    
    ' r = Calc(TimeArr)
    
    ' With modTest
    '     .Name = init.Title
    '     .Count = init.iCount
    '     .Data = TimeArr
    '     .High = r.Max
    '     .Low = r.Min
    '     .Avg = r.Avg
    ' End With

ExitSub:
    Exit Sub

End Sub

Private Sub Write_Results() ' Write results

    Debug.Print "Write Results"
    Dim pWrite As String
    With modTest
        pWrite = "Name of Test: " & .Title
        pWrite = pWrite & vbCrLf & "No. of Tests: " & .Coun
        pWrite = pWrite & vbCrLf &  "Highest Time: " & .Max
        pWrite = pWrite & vbCrLf &  "Lowest  Time: " & .Min
        pWrite = pWrite & vbCrLf &  "Average Time: " & .Avg
    End With

    PrintToScreen pWrite

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
            .clear
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
    If not DSE then CreateNewWorksheet(SheetName)
    
End Sub

Sub Test()

    Initialize 'Call Initialize
    LoopCode 'Call LoopCode
    Get_Results 'Call Get_Results
    AddResultsSheet("pfResults") 'Call AddResultsSheet
    Write_Results 'Call Write_Results
    Terminate 'Call Terminate
    
End Sub