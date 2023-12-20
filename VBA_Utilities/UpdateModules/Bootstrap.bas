' I do not own the following code and is modified to fit for my use case.
' The original can be found at https://gist.github.com/3403537
'
' IMPORTANT - CHANGE gtLoad_XML() to include gtDoit() for each Gist you want to load
' bootstrap code to update VBA modules from gists
' all code is in this module - so no classes etc.
' latebinding is used to avoid need for any references
' can be found at https://gist.github.com/3403537
'
Attribute VB_Name = "Bootstrap"
' Add namespace for XVBA file structure
'namespace=vba-files\VBA_Utilities\UpdateModules
'@Folder("VBA_Utilities\UpdateModules")
Option Explicit
'
' v2.04 - 3403537
' if you are using your own gists - change this
Const gistOwner = "rjwren-brd"

Public Function gtLoad_XML()
' this is an example of how you would load your VBE with a particular manifest
' you could set the 2nd parameter to override conflict checking the first time used-

' we are going to need fpcode.xml
  gtDoit "b2b2857ae5ab39388a5a2b419d3768d9", True

  
End Function

Public Function gtMakeManifest()
    ' this is an example of how you would create a manifest to be loaded up as a Gist
    '
    Dim dom As Object ' DOMDocument
    Set dom = gtInitManifest("Fingerprint and associated classes and modules", "richard.wren79@brdnest.net")
    '
    ' call this for each required gist of the manifest
    '---cDataSet
    '---Fingerprint
    gtAddToManifest dom, "6b4abf42c9d668a6a9b53a08a6532269", "module", "AppQuit", "AppQuit.bas"
    gtAddToManifest dom, "6b4abf42c9d668a6a9b53a08a6532269", "module", "ChangeLog", "ChangeLog.bas"
    gtAddToManifest dom, "6b4abf42c9d668a6a9b53a08a6532269", "module", "TermsAndConditions", "TermsAndConditions.bas"
    gtAddToManifest dom, "6b4abf42c9d668a6a9b53a08a6532269", "class", "WorkbookCode", "WorkbookCode.cls"

    
    ' cut and paste the result of this into a gist - this will be your manifest
    Debug.Print dom.XML
    
End Function

Private Function gtCreateReferences(dom As Object) 'DOMDocument)
    ' adds all current references to an xml
    Dim r As Object ' Reference
    
    With ActiveWorkbook.VBProject
        For Each r In .References
            gtAddRefToManifest dom, r
        Next r
    End With

End Function
Public Function gtUpdateAll()
    ' do all updates for known manifests in this project
    Dim modle As Object 'VBComponent
    Dim stampLine As Long, co As New Collection, manifest As String, s As String, v As Variant
    For Each modle In ThisWorkbook.VBProject.VBComponents
        ' do we know this module?
        stampLine = gtManageable(modle)
        If stampLine > 0 Then
            ' yes we do - get the manifest
            manifest = gtStampManifest(modle, stampLine)
            If (manifest = vbNullString) Then
                MsgBox ("Stamp line in module " & modle.name & " fiddled with. Run again as greenField")
            Else
            ' add it to the collection of manifests we need to process
              If gtCoIndex(manifest, co) = 0 Then
                co.Add manifest, manifest
              End If
            End If
        End If
    Next modle
    ' todo check if versions are up to date
    If co.Count > 0 Then
        s = vbNullString
        For Each v In co
            gtDoit CStr(v)
            If s <> vbNullString Then s = s & ","
            s = s & CStr(v)
        Next v
        ReWrite_TWb
        MsgBox ("updated " & co.Count & " manifests: (" & s & ")")
    End If
End Function
Private Function gtCoIndex(sid As Variant, co As Collection) As Long
    ' find index in a collection
    Dim i As Long
    gtCoIndex = 0
    For i = 1 To co.Count
        If co(i) = sid Then
          gtCoIndex = i
          Exit Function
        End If
    Next i
End Function
Private Function gtPreventCaching(url As String) As String
    ' this will tweak the url with an extra random parameter to prevent any accidental caching
    Dim p As String
    If (InStr(1, url, "?") > 0) Then
        p = "&"
    Else
        p = "?"
    End If
    Randomize
    gtPreventCaching = url & p & "gtPreventCaching=" & CStr(Int(10000 * Rnd()))
End Function
Public Function gtDoit(gtDoitmanifestID As String, Optional greenField As Boolean = False) As Boolean
    Dim dom As Object ' DOMDocument
    Dim rawUrl As String, t As String, n As String, g As String
    Dim xNode As Object ' IXMLDOMNode
    Dim attrib As Object 'IXMLDOMAttribute
    Dim vbCom As Object 'VBComponent
    ' get the requested manifest
    Set dom = gtRecreateManifest(gtDoitmanifestID)

    ' now we know which gists are needed here
    If (gtWillItWork(dom, greenField)) Then
        'theres a good chance it will work
        ' for each module
        For Each xNode In dom.SelectSingleNode("//gists").ChildNodes
            t = xNode.Attributes.getNamedItem("type").Text
            Select Case t
                Case "class", "module"
                    ' get the gist
                    rawUrl = gtConstructRawUrl(xNode.Attributes.getNamedItem("gistid").Text, _
                                            xNode.Attributes.getNamedItem("filename").Text)
                    ' prevent caching will make it look like a different request
                    g = gtHttpGet(gtPreventCaching(rawUrl))
                    ' module name
                    n = xNode.Attributes.getNamedItem("module").Text
                    ' does it exist - if so then delete it
                    Set vbCom = gtModuleExists(n, ThisWorkbook)
                    If (Not vbCom Is Nothing) Then
                        ' delete everything in it
                        With vbCom.CodeModule
                            .DeleteLines 1, .CountOfLines
                        End With
                    Else
                        Set vbCom = gtAddModule(n, ThisWorkbook, xNode.Attributes.getNamedItem("type").Text)
                    End If
        
                    ' add in the new code
                    With vbCom.CodeModule
                        .AddFromString g
                    End With
        
                    ' stamp it
                    gtInsertStamp vbCom, gtDoitmanifestID, rawUrl
                
                Case "reference"
                    gtAddReference xNode.Attributes.getNamedItem("name").Text, _
                                   xNode.Attributes.getNamedItem("guid").Text, _
                                   xNode.Attributes.getNamedItem("major").Text, _
                                   xNode.Attributes.getNamedItem("minor").Text
                Case Else
                    Debug.Assert False
            
            End Select
        Next xNode
        gtDoit = True
    Else

    End If
End Function

Private Function gtAddReference(name As String, guid As String, major As String, minor As String) As Object ' Reference
    ' add a reference (if its not already there)
    Dim r As Object ' Reference
    On Error GoTo handle
    With ActiveWorkbook.VBProject
        For Each r In .References
            If (r.name = name) Then
                If (r.major < major Or r.major = major And r.minor < minor And Not r.BuiltIn) Then
                    .References.AddFromGuid guid, major, minor
                    .References.Remove (r)
                End If
                Exit Function
            End If
        Next r
    ' if we get here then we need to add it
      Set gtAddReference = .References.AddFromGuid(guid, major, minor)
      Exit Function
    End With
    
handle:
    MsgBox ("warning - tried and failed to add reference to " & name & "v" & major & "." & minor)
    Exit Function
    
End Function
Private Function gtStampManifest(vbCom As Object, line As Long) As String 'VBComponent
    ' the manifest should be on the given line
    Dim s As String, n As Long, p As Long, marker As String
    marker = "manifest: "
    s = vbNullString
    With vbCom.CodeModule
       n = InStr(1, LCase(.Lines(line, 1)), marker)
       If (n > 0) Then
        s = Mid(.Lines(line, 1), n + Len(marker))
        p = InStr(1, s, " ")
        s = Left(s, p - 1)
       End If
    End With
    gtStampManifest = s
End Function
Private Function gtInsertStamp(vbCom As Object, manifest As String, rawUrl As String) As Long 'VBComponent
    Dim stampLine As Long
    stampLine = gtManageable(vbCom)
    ' if it wasnt found then insert at line 1
    With vbCom.CodeModule
        If stampLine <> 0 Then
            .DeleteLines stampLine, 1
        Else
            stampLine = 1
        End If
        .InsertLines stampLine, gtStampLog(manifest, rawUrl)
    End With
    gtInsertStamp = stampLine
    
End Function
Private Function gtWillItWork(dom As Object, _
                Optional greenField As Boolean = False) As Boolean 'DOMDocument
    
    Dim xNode As Object ' IXMLDOMNode
    Dim attrib As Object 'IXMLDOMAttribute
    Dim n As String, s As String, t As String
    Dim modle As Object 'VBComponent
    
    ' check we have something to do
    gtWillItWork = Not dom Is Nothing
    If Not gtWillItWork Then
        Exit Function
    End If
    ' first we check if these are new modules
    s = vbNullString
    For Each xNode In dom.SelectSingleNode("//gists").ChildNodes
        ' the target module
        t = xNode.Attributes.getNamedItem("type").Text
        Select Case t
            Case "class", "module"
                n = xNode.Attributes.getNamedItem("module").Text
                Set modle = gtModuleExists(n, ThisWorkbook)
                
                If (Not modle Is Nothing) Then
                    ' it exists - validate its not somethig else with the same name
                    If (gtManageable(modle) = 0 And Not greenField) Then
                        s = gtAddStr(s, n)
                    End If
                End If
            Case "reference"
            Case Else
                s = gtAddStr(s, "unknown type " & t)
        End Select
    Next xNode
    
    If (s <> vbNullString) Then
        MsgBox ("there may be a conflict with these modules names (" & s & _
            ") and some others in your project. " & _
            "If this is the first time you have run this - run with greenfield set to true to override this check")
        gtWillItWork = False
        Exit Function
    End If
   
   ' now check all gists are getable
   ' todo
   
End Function
Private Function gtAddStr(t As String, n As String) As String
    Dim s As String
    s = t
    If (s <> vbNullString) Then s = s & ","
    gtAddStr = s & n
End Function

Private Function gtRecreateManifest(manifestID As String) As Object 'DOMDocument
    Dim dom As Object 'DOMDocument
    Dim manifest As String
    
    ' get the xml string
    manifest = gtHttpGet(gtPreventCaching(gtConstructRawUrl(manifestID)))
    
    If manifest <> vbNullString Then
    ' parse the xml
        Set dom = CreateObject("MSXML.DOMDocument")
        dom.LoadXML (manifest)
        Set gtRecreateManifest = dom
    Else
        MsgBox ("Could not get manifest for " & manifestID)
    End If
End Function

Private Function gtModuleExists(name As String, wb As Workbook) As Object 'VBComponent
    ' determine whether this module exists in the given workbook
    Dim modle As Object 'VBComponent
    For Each modle In wb.VBProject.VBComponents
       If Trim(LCase(modle.name)) = Trim(LCase(name)) Then
        Set gtModuleExists = modle
        Exit Function
       End If
    Next modle
End Function

Private Function gtAddModule(name As String, wb As Workbook, modType As String) As Object ' VBComponent
    ' determine whether this module exists in the given workbook
    Dim modle As Object, t As String ' VBComponent, t As Long

    Select Case LCase(modType)
        Case "class"
            t = 2
        Case "module"
            t = 1
        Case Else
            MsgBox ("unknown module type " & modType)
    End Select
        
    Set modle = wb.VBProject.VBComponents.Add(t)
    modle.name = name
    
    ' added by andypope.info
    If modle.CodeModule.CountOfLines > 1 Then
        ' remove Option Explict lines if it was added automatically
        modle.CodeModule.DeleteLines 1, modle.CodeModule.CountOfLines
    End If
    
    Set gtAddModule = modle
End Function

Private Function gtConstructRawUrl(gistID As String, _
                Optional gistFileName As String = vbNullString) As String
    ' given a gist, where is it?
    Dim s As String
    ' raw URL
    s = "https://gist.github.com/" & gistOwner & "/" & gistID & "/raw"

    ' a gist can have multiple files in it
    If gistFileName <> vbNullString Then s = s & "/" & gistFileName
    ' TODO - specific versions
    gtConstructRawUrl = s
End Function

Private Function gtAddToManifest(dom As Object, _
                                 gistID As String, _
                                 modType As String, _
                                 modle As String, _
                                 Optional Filename As String = vbNullString, _
                                 Optional version As String = vbNullString _
                        ) As Object ' DOMDocument
                                 
    Dim Element As Object 'IXMLDOMElement
    Dim attrib As Object 'IXMLDOMAttribute
    Dim elements As Object 'IXMLDOMNodeList
    Dim head As Object 'IXMLDOMElement
    ' add an item to the manifest element - returns the dom for chaining
    Set elements = dom.getElementsByTagName("gists")
    Set head = elements.NextNode
    Set Element = dom.createElement("item" & CStr(head.ChildNodes.Length + 1))
    head.appendChild Element
    
    Set attrib = dom.createAttribute("gistid")
    attrib.NodeValue = gistID
    Element.setAttributeNode attrib
    
    Set attrib = dom.createAttribute("version")
    attrib.NodeValue = version
    Element.setAttributeNode attrib
    
    Set attrib = dom.createAttribute("filename")
    attrib.NodeValue = Filename
    Element.setAttributeNode attrib
    
    Set attrib = dom.createAttribute("module")
    attrib.NodeValue = modle
    Element.setAttributeNode attrib
    
    Set attrib = dom.createAttribute("type")
    attrib.NodeValue = modType
    Element.setAttributeNode attrib
    
    Set gtAddToManifest = dom
End Function
Private Function gtAddRefToManifest(dom As Object, r As Object) As Object   ' DOMDocument, reference, domdocument
                                 
    Dim Element As Object 'IXMLDOMElement
    Dim attrib As Object 'IXMLDOMAttribute
    Dim elements As Object 'IXMLDOMNodeList
    Dim head As Object 'IXMLDOMElement
    
    ' add an item to the manifest element - returns the dom for chaining
    Set elements = dom.getElementsByTagName("gists")
    Set head = elements.NextNode
    Set Element = dom.createElement("item" & CStr(head.ChildNodes.Length + 1))
    head.appendChild Element
    'r.GUID, r.name, r.Major, r.Minor, r.description
    Set attrib = dom.createAttribute("guid")
    attrib.NodeValue = r.guid
    Element.setAttributeNode attrib
    
    Set attrib = dom.createAttribute("name")
    attrib.NodeValue = r.name
    Element.setAttributeNode attrib
    
    Set attrib = dom.createAttribute("major")
    attrib.NodeValue = r.major
    Element.setAttributeNode attrib
    
    Set attrib = dom.createAttribute("minor")
    attrib.NodeValue = r.minor
    Element.setAttributeNode attrib

    Set attrib = dom.createAttribute("description")
    attrib.NodeValue = r.description
    Element.setAttributeNode attrib
    
    Set attrib = dom.createAttribute("type")
    attrib.NodeValue = "reference"
    Element.setAttributeNode attrib
    
    Set gtAddRefToManifest = dom
End Function
Private Function gtInitManifest(Optional description As String = vbNullString, _
                                 Optional contact As String = vbNullString) As Object ' DOMDocument
    Dim Element As Object 'IXMLDOMElement
    Dim attrib As Object 'IXMLDOMAttribute
    Dim dom As Object ' DOMDocument
    Dim e2 As Object ' IXMLDOMElement
   
 ' creates an xml manifest of required gists
   Set dom = CreateObject("MSXML.DOMDocument")

    Set Element = dom.createElement("fpcode")
    Set attrib = dom.createAttribute("info")
    attrib.NodeValue = _
            "this is a manifest for fpcode VBA code distribution " & _
            " - email for details"
    Element.setAttributeNode attrib

    
    With dom.appendChild(Element)
        Set Element = dom.createElement("manifest")
        .appendChild Element
        
        Set attrib = dom.createAttribute("description")
        attrib.NodeValue = description
        Element.setAttributeNode attrib

        Set attrib = dom.createAttribute("contact")
        attrib.NodeValue = contact
        Element.setAttributeNode attrib

        Element.appendChild dom.createElement("gists")

    End With
    Set gtInitManifest = dom
   
End Function
 
Private Function gtHttpGet(url As String) As String
    ' TODO oAuth
    Dim ohttp As Object
    Set ohttp = CreateObject("Msxml2.ServerXMLHTTP.6.0")
    Call ohttp.Open("GET", url, False)
    Call ohttp.Send("")
    gtHttpGet = ohttp.ResponseText
    Set ohttp = Nothing
End Function
Private Function gtStampLog(manifest As String, rawUrl As String) As String
    ' create a comment to identify this as manageable
    'Dim mName As
    'Set mName =
    gtStampLog = gtStamp & " updated on: " & Now() & _
        " from manifest: " & manifest & " " & rawUrl
        
End Function
Private Function gtStamp() As String
' this marks a module as manageable
    gtStamp = "' developer@brdnest.net :do not modify following lines" & _
    " - email for details:"
End Function
Private Function gtManageable(vbCom As Object) As Long  ' VBComponent
    ' return the line number of the gtStamp
    ' parameters as passed by ref in .find method
    Dim startLine As Long, startColumn As Long, endLine As Long, endColumn As Long
    startLine = 1: endLine = vbCom.CodeModule.CountOfLines: startColumn = 1: endColumn = 255
    
    If (vbCom.CodeModule.Find(gtStamp(), startLine, startColumn, endLine, endColumn)) Then
        gtManageable = startLine
    End If
    
End Function

Function ReWrite_TWb()
    
    'OptimizeVBA True
    Dim wb As Workbook
    Set wb = ActiveWorkbook
    
    Debug.Print "Importing lines from " & "WorkbookCode" & "."
    
    ' At this point compilation errors may cause a crash, so we ignore those.
    On Error Resume Next
    With ThisWorkbook
        .VBProject.VBComponents("WorkbookCode").Export .Path & "\WkBkCode.txt"
        AltText
        .VBProject.VBComponents("ThisWorkbook").CodeModule.DeleteLines 1, _
            .VBProject.VBComponents("ThisWorkbook").CodeModule.CountOfLines
        wb.VBProject.VBComponents("ThisWorkbook").CodeModule.AddFromFile .Path & "\WkBkCode.txt"
        Kill .Path & "\WkBkCode.txt"
        '.VBProject.VBComponents("ThisWorkbook").CodeModule.DeleteLines 8, 4
    End With
    'On Error GoTo 0
    
    'OptimizeVBA False
    
End Function

Sub AltText()
    Dim File As String
    Dim VecFile() As String, Aux As String
    Dim i As Long, j As Long
    Dim SizeNewFile As Long
    
    File = ThisWorkbook.Path & "\WkBkCode.txt"
    
    'Import file lines to array excluding first 3 lines and
    'lines starting with "-"
    Open File For Input As 1
    i = 0
    j = 0
    Do Until EOF(1)
        j = j + 1
        Line Input #1, Aux
        If j > 9 And InStr(1, Aux, "-") <> 1 Then
            i = i + 1
            ReDim Preserve VecFile(1 To i)
            VecFile(i) = Aux
        End If
    Loop
    Close #1
    SizeNewFile = i
    
    'Write array to file
    Open File For Output As 1
    For i = 1 To SizeNewFile
        Print #1, VecFile(i)
    Next i
    Close #1
    
    'MsgBox "File alteration completed!"
    
End Sub