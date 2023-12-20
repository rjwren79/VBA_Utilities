VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsEmployee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Add namespace for XVBA file structure
'namespace=vba-files\VBA_Utilities\NameFormat
'@Folder("VBA_Utilities\NameFormat")
Option Explicit

Private pFullName As String
Private pFileAs As String
Private pFirstName As String
Private pMiddleName As String
Private pSurName As String
Private pSuffix As String
Private pEntry As String

'==============================
'Properties
'==============================

Public Property Get FirstName() As String
    FirstName = pFirstName
End Property

Public Property Let FirstName(ByVal value As String)
    pFirstName = value
End Property

Public Property Get MiddleName() As String
    MiddleName = pMiddleName
End Property

Public Property Let MiddleName(ByVal value As String)
    pMiddleName = value
End Property

Public Property Get SurName() As String
    SurName = pSurName
End Property

Public Property Let SurName(ByVal value As String)
    pSurName = value
End Property

Public Property Get Suffix() As String
    Suffix = pSuffix
End Property

Public Property Let Suffix(ByVal value As String)
    pSuffix = value
End Property

Public Property Get FullName() As String
    Call CompileName
    FullName = pFullName
End Property

Public Property Get FileAs() As String
    Call CompileName
    FileAs = pFileAs
End Property

Public Property Get entry() As String
    entry = pEntry
End Property

Public Property Let entry(ByVal value As String)
    pEntry = value
End Property

'==============================
'Public Methods
'==============================

'==============================
'Private Methods
'==============================

Private Sub CompileName()
    If pEntry = "" Then Exit Sub
    
    Dim mSurName As String, mGivenName As String
    
    mSurName = pSurName
    mGivenName = pFirstName
    
    If Not pSuffix = "" Then
        mSurName = pSurName & " " & pSuffix
    End If
    
    If Not pMiddleName = "" Then
        mGivenName = pFirstName & " " & pMiddleName
    End If
    pFullName = mGivenName & " " & mSurName
    pFileAs = mSurName & ", " & mGivenName
        
End Sub


