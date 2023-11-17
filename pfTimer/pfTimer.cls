VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "pfTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Add namespace for XVBA file structure
'namespace=vba-files\pfTimer
'@Folder("vba-files\pfTimer")
Option Explicit
 
Private Type LARGE_INTEGER
    lowpart As Long
    highpart As Long
End Type

    ' Declare the API functions as private members of the class
#If VBA7 Then
    Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As LARGE_INTEGER) As Long
    Private Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As LARGE_INTEGER) As Long
    
#Else
    '  Code is running in 32-bit version of Microsoft Office
    Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As LARGE_INTEGER) As Long
    Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As LARGE_INTEGER) As Long
    
#End If

Private m_CounterStart As LARGE_INTEGER, m_CounterEnd As LARGE_INTEGER, m_crFrequency As Double
'Private pTestID As String, pTestResult, pResultHigh, pResultLow, pResultAvg

Private Const TWO_32 = 4294967296# ' = 256# * 256# * 256# * 256#

Private Function LI2Double(LI As LARGE_INTEGER) As Double

    Dim Low As Double
    Low = LI.lowpart
    If Low < 0 Then
        Low = Low + TWO_32
    End If
    LI2Double = LI.highpart * TWO_32 + Low
    
End Function

Private Sub class_initialize()

    Dim PerfFrequency As LARGE_INTEGER
    QueryPerformanceFrequency PerfFrequency
    m_crFrequency = LI2Double(PerfFrequency)
    
End Sub

Public Sub Clock_Start()

    QueryPerformanceCounter m_CounterStart
    
End Sub

'Public Sub Clock_Stop()
'
'    pTestResult = TimeElapsed
'
'End Sub

Property Get TimeElapsed() As Double

    Dim crStart As Double
    Dim crStop As Double
    QueryPerformanceCounter m_CounterEnd
    crStart = LI2Double(m_CounterStart)
    crStop = LI2Double(m_CounterEnd)
    TimeElapsed = 1000# * (crStop - crStart) / m_crFrequency
    
End Property

' Public Property Get TestID() As String
'     TestID = pTestID
' End Property

' Public Property Let TestID(ByVal value As String)
'     pTestID = value
' End Property

' Public Property Get TestResult() As String
'     TestResult = pTestResult
' End Property

' Public Property Let TestResult(ByVal value As String)
'     pTestResult = value
' End Property

'Public Property Get ResultHigh() As String
'    ResultHigh = pResultHigh
'End Property
'
'Public Property Let ResultHigh(ByVal value As String)
'    pResultHigh = value
'End Property
'
'Public Property Get ResultLow() As String
'    ResultLow = pResultLow
'End Property
'
'Public Property Let ResultLow(ByVal value As String)
'    pResultLow = value
'End Property
'
'Public Property Get ResultAvg() As String
'    ResultAvg = pResultAvg
'End Property
'
'Public Property Let ResultAvg(ByVal value As String)
'    pResultAvg = value
'End Property