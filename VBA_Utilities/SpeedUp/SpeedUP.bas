Attribute VB_Name = "SpeedUP"
' Add namespace for XVBA file structure
'namespace=vba-files\VBA_Utilities\SpeedUP
Option Explicit
Option Private Module

Sub SpeedUp(Optional DoIt As Boolean = True)

    'This macro will set most properties for speeding up macro execution.
    '  Before running your macros, you put 'SpeedUp' as one of the first
    '  commands in your code, and, as one of the last lines in your code,
    '  you put 'SpeedUp (False)' to reset the properties.
    
    'Retrieved from www.excelguard.dk


    'Initialize
    With Application
          .Cursor = xlWait
          .DisplayStatusBar = True
          .WindowState = xlMaximized
          '.VBE.MainWindow.Visible = False
          .EnableEvents = False
          .DisplayAlerts = False
          .ScreenUpdating = False
          .Calculation = xlCalculationManual
        ' .Interactive = False
          .AskToUpdateLinks = False
          .IgnoreRemoteRequests = False
       If ThisWorkbook.IsAddin Then .EnableCancelKey = xlDisabled
    End With

    On Error Resume Next


    'Define variables
    Dim ws As Worksheet
    Dim WB As Workbook
    Set WB = ActiveWorkbook


    'Don't display pagebreaks
    ActiveSheet.DisplayPageBreaks = False
    ActiveSheet.DisplayAutomaticPageBreaks = False

    For Each ws In WB.Worksheets
          ws.DisplayPageBreaks = False
          ws.DisplayAutomaticPageBreaks = False
    Next


    'Set workbook properties
    With WB
          .AcceptAllChanges
          .SaveLinkValues = False
          .UpdateRemoteReferences = True
          .UpdateLinks = xlUpdateLinksAlways
          .ConflictResolution = xlUserResolution
          .Colors(14) = RGB(0, 153, 153)
    End With


    'Skip to speedup
    Set WB = Nothing
    Set ws = Nothing

    If DoIt = True Then Exit Sub


ES: ' End of Sub
    With Application
          .Calculation = xlCalculationAutomatic
          .ScreenUpdating = True
          .DisplayAlerts = True
          .EnableEvents = True
          .EnableCancelKey = xlInterrupt
          .CutCopyMode = False
          .Interactive = True
          .Cursor = xlDefault
          .StatusBar = False
    End With

End Sub
