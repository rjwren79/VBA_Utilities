Attribute VB_Name = "pfOpti"
' Add namespace for XVBA file structure
'namespace=vba-files\pfOptimize
Option Explicit
Option Private Module
Private sCalcMode As XlCalculation
Private sScreenUpdating As Boolean, sEnableEvents As Boolean
Private sPageBrakes As Boolean, sEnableAnimations As Boolean
Private sStatusBar As Boolean, sPrintCommunication As Boolean
Private swStatus As Boolean

Public Sub Enable() ' Turn off everything but the essentials

    SwitchOff (True)
    Debug.Print oStat
    
End Sub

Public Sub Disable() ' Recover previous state

    SwitchOff (False)
    Debug.Print oStat

End Sub

Public Sub Reset() ' Reset to default state

    ResetSwitch
    PrintToScreen oStatus
    'PrintToScreen Optimization.oStatus
End Sub

Private Sub oState() ' Store current optimization state

    sCalcMode = Application.Calculation
    sScreenUpdating = Application.ScreenUpdating
    sEnableEvents = Application.EnableEvents
    sPageBrakes = ActiveSheet.DisplayPageBreaks
    sEnableAnimations = Application.EnableAnimations
    sStatusBar = Application.DisplayStatusBar
    sPrintCommunication = Application.PrintCommunication

End Sub

Private Function oStatus() As String ' Get current status
    
    Dim Msg As String
    
    Msg = oStat
    Msg = Msg & vbCrLf & "Calculation: " & GetEnum(Application.Calculation)
    Msg = Msg & vbCrLf & "Screen Updating: " & Application.ScreenUpdating
    Msg = Msg & vbCrLf & "Enable Events: " & Application.EnableEvents
    Msg = Msg & vbCrLf & "Display Page Breaks: " & ActiveSheet.DisplayPageBreaks
    Msg = Msg & vbCrLf & "Enable Animations: " & Application.EnableAnimations
    Msg = Msg & vbCrLf & "Display Status Bar: " & Application.DisplayStatusBar
    Msg = Msg & vbCrLf & "Print Communication: " & Application.PrintCommunication
    
    oStatus = Msg
    
End Function

Private Sub SwitchOff(bSwitchOff As Boolean) ' Toggle state
      
    swStatus = bSwitchOff
    
    ActiveSheet.DisplayPageBreaks = False
    
    With Application
        If bSwitchOff Then
        ' OFF
            
            oState 'Save current state
            
            .Calculation = xlCalculationManual
            .ScreenUpdating = False
            .EnableEvents = False
            .EnableAnimations = False
            .DisplayStatusBar = False
            .PrintCommunication = False
        Else
        
        ' ON
        ' Restore state from save
        
            .Calculation = sCalcMode
            .ScreenUpdating = sScreenUpdating
            .EnableEvents = sEnableEvents
            .EnableAnimations = sEnableAnimations
            .DisplayStatusBar = sStatusBar
            .PrintCommunication = sPrintCommunication
    
        End If
    
    End With
    
    ActiveSheet.DisplayPageBreaks = sPageBrakes
    
End Sub

Private Sub ResetSwitch()

    With Application
        .Calculation = xlCalculationAutomatic
        .ScreenUpdating = True
        .EnableEvents = True
        .EnableAnimations = True
        .DisplayStatusBar = True
        .PrintCommunication = True
    End With

End Sub

Private Function GetEnum(ByVal value As Long) As String ' Get calculation value string

    Select Case value
        Case -4105
            GetEnum = "xlCalculationAutomatic"
        Case -4135
            GetEnum = "xlCalculationManual"
        Case Else
            GetEnum = "Not Defined"
    End Select

End Function

Private Function oStat() As String
    
    oStat = "Optimization: Disabled"
    If swStatus = True Then oStat = "Optimization: Enabled"

End Function
