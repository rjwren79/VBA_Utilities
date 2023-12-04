VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "pfTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Add namespace for XVBA file structure
'namespace=VBA_Utilities\pfTimer
'@Folder("VBA_Utilities\pfTimer")
Option Explicit

'==============================
'Properties
'==============================

' Private Type Calculate
'     Max As Double
'     Min As Double
'     Avg As Double
' End Type

Private pTitle As String, pCount As Long, pData As Variant, _
    pInit As String, pTerm As String

Public Property Get Title() As String
    Title = pTitle
End Property

Public Property Let Title(ByVal value As String)
    pTitle = value
End Property

Public Property Get Count() As Long
    Count = pCount
End Property

Public Property Let Count(ByVal value As Long)
    pCount = value
End Property

Public Property Get Data() As Variant
    Data = pData
End Property

Public Property Let Data(ByVal value As Variant)
    pData = value
End Property

Public Property Get Max() As Double
    Max = WorksheetFunction.Max(pData)
End Property

' Public Property Let High(ByVal value As Double)
'     pHigh = value
' End Property

Public Property Get Min() As Double
    Min = WorksheetFunction.Min(pData)
End Property

' Public Property Let Low(ByVal value As Double)
'     pLow = value
' End Property

Public Property Get Avg() As Double
    Avg = WorksheetFunction.Average(pData)
End Property

' Public Property Let Avg(ByVal value As Double)
'     pAvg = value
' End Property

Public Property Get ProcessInit() As String
    ProcessInit = pInit
End Property

Public Property Let ProcessInit(ByVal value As String)
    pInit = value
End Property

Public Property Get ProcessTerm() As String
    ProcessTerm = pTerm
End Property

Public Property Let ProcessTerm(ByVal value As String)
    pTerm = value
End Property

'==============================
'Public Methods
'==============================

'==============================
'Private Methods
'==============================
' Private Function Calc(Arr As Variant) As Calculate ' Calculate results
'
'         With WorksheetFunction
'             Calc.Max = .Max(Arr)
'             Calc.Min = .Min(Arr)
'             Calc.Avg = .Average(Arr)
'         End With
'
'  End Function

