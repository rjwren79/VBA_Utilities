Attribute VB_Name = "Optimization"
' Add namespace for XVBA file structure
'namespace=vba-files\Optimization
Option Explicit
Option Private Module
Private sCalcMode As XlCalculation, sScreenUpdating As Boolean, sEnableEvents As Boolean
Private sPageBrakes As Boolean, sEnableAnimations As Boolean, sStatusBar As Boolean
Private sPrintCommunication As Boolean, swStatus As Boolean


Public Sub Enable()

    SwitchOff (True) 'Turn off everything but the essentials
    Status 'Call Status
    
End Sub

Public Sub Disable()

    SwitchOff (False) 'turn features back on
    Status 'Call Status

End Sub

Private Sub State()
    
    sCalcMode = Application.Calculation
    sScreenUpdating = Application.ScreenUpdating
    sEnableEvents = Application.EnableEvents
    sPageBrakes = ActiveSheet.DisplayPageBreaks
    sEnableAnimations = Application.EnableAnimations
    sStatusBar = Application.DisplayStatusBar
    sPrintCommunication = Application.PrintCommunication

End Sub

Private Sub Status()
    
    Dim oStat As String, Msg As String
    
    oStat = "Disabled"
    If swStatus = True Then oStat = "Active"
    
    Msg = "Optimization is " & oStat
    Msg = Msg & vbCrLf & "Calculation: " & GetEnum(Application.Calculation)
    Msg = Msg & vbCrLf & "Screen Updating: " & Application.ScreenUpdating
    Msg = Msg & vbCrLf & "Enable Events: " & Application.EnableEvents
    Msg = Msg & vbCrLf & "Display Page Breaks: " & ActiveSheet.DisplayPageBreaks
    Msg = Msg & vbCrLf & "Enable Animations: " & Application.EnableAnimations
    Msg = Msg & vbCrLf & "Display Status Bar: " & Application.DisplayStatusBar
    Msg = Msg & vbCrLf & "Print Communication: " & Application.PrintCommunication
    Msg = Msg & vbCrLf & "======================================================="
    
    #If ImWindow = 1 Then
        Debug.Print Msg
    #End If
    
End Sub

Private Sub SwitchOff(bSwitchOff As Boolean)
    
    State 'Call State
    
    swStatus = bSwitchOff
    
    ActiveSheet.DisplayPageBreaks = False
    
    With Application
        If bSwitchOff Then
        
        ' OFF
            .Calculation = xlCalculationManual
            .ScreenUpdating = False
            .EnableEvents = False
            .EnableAnimations = False
            '.DisplayStatusBar = False
            .PrintCommunication = False
        Else
    
        ' ON
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

Private Function GetEnum(ByVal value As Long) As String

    Select Case value
        Case -4105
            GetEnum = "xlCalculationAutomatic"
        Case -4135
            GetEnum = "xlCalculationManual"
        Case Else
            GetEnum = "Not Defined"
    End Select

End Function