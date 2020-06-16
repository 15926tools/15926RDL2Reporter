Attribute VB_Name = "Module1"
Option Explicit

Public Const WS_NAME_PROGRAM = "Program"
Public Const WS_NAME_REPORT = "Result1"

Public gEndpoint As String
Public WS_PROGRAM, WS_REPORT As Worksheet


Sub init()
    Set WS_PROGRAM = Worksheets(WS_NAME_PROGRAM)
    gEndpoint = WS_PROGRAM.Cells(2, 2)
    If Len(gEndpoint) < 2 Then
        MsgBox ("Set an endpoint to query")
    End If
    
    createNEWPage
    
End Sub
Sub createNEWPage()
    Dim i

    Application.DisplayAlerts = False
    On Error Resume Next
    Sheets(WS_NAME_REPORT).Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    Set WS_REPORT = Sheets.Add(After:=Worksheets(Worksheets.Count))
    WS_REPORT.Name = WS_NAME_REPORT
    
    WS_REPORT.Columns("A:Q").WrapText = True
    WS_REPORT.Columns("A:Q").VerticalAlignment = xlVAlignTop


End Sub


Sub Run_the_report()
    
    Dim sparql_text, Line, results, el, y, bindings, t, c, n, x, TargetAddress, text, id, mainLabel, _
        ClassOfClass, mainID, sourceLine, Query, Spreadsheet, SuperClassLevel, EntityLevel, _
        columnName, multiColumns, reportColumnNo, ColumnWidth, lastPK, newValue, colPK
    Dim Label
    Dim htmQry As Object
    
    Dim list_columns As Dictionary
    Set list_columns = New Dictionary
    
    Dim list_values As Dictionary
    Set list_values = New Dictionary

    init
    
    'header
    reportColumnNo = 1
    For n = 9 To 60
        columnName = WS_PROGRAM.Cells(n, 4)
        multiColumns = WS_PROGRAM.Cells(n, 2)
        ColumnWidth = WS_PROGRAM.Cells(n, 5)
        If columnName > " " Then
            For x = reportColumnNo To reportColumnNo + multiColumns - 1
                WS_REPORT.Cells(1, x) = columnName
                WS_REPORT.Cells(1, x).ColumnWidth = ColumnWidth
            Next
            reportColumnNo = reportColumnNo + multiColumns
        End If
    Next
    DoEvents
    
    'list of columns
    reportColumnNo = 0
    For n = 9 To 50
        columnName = WS_PROGRAM.Cells(n, 1)
        multiColumns = WS_PROGRAM.Cells(n, 2)
        If columnName > " " Then
            reportColumnNo = reportColumnNo + 1
            list_columns.Add columnName, reportColumnNo
            reportColumnNo = reportColumnNo + multiColumns - 1
        End If
    Next
    
    'set primary key
    colPK = ""
    For x = 9 To 50
        If WS_PROGRAM.Cells(x, 3) = "yes" Then
            colPK = WS_PROGRAM.Cells(x, 1)
            Exit For
        End If
    Next
    If colPK = "" Then
        MsgBox "error. No primary key given (need 1 ""yes"" in column C"
    Else
    
        Line = 1
        lastPK = "x"
        sparql_text = WS_PROGRAM.Cells(3, 2)
        If Mid(sparql_text, 1, 6) = "prefix" Or Mid(sparql_text, 1, 6) = "select" Then
            Set htmQry = getHtm(Replace(sparql_text, "{mainID}", mainID))
            
            Set results = htmQry.getElementsByTagName("result")
            
            For Each el In results
                If el.innerHTML > " " Then
                    Set bindings = el.getElementsByTagName("binding")
                    For n = 9 To 50
                        columnName = WS_PROGRAM.Cells(n, 1)
                        multiColumns = WS_PROGRAM.Cells(n, 2)
                        reportColumnNo = 0
                        If columnName > " " Then
                            reportColumnNo = list_columns(columnName)
                            For x = reportColumnNo To reportColumnNo + multiColumns - 1
                                For Each t In bindings
                                    If t.Name = columnName Then
                                        text = namespacehandler(Trim(t.innertext))
                                        If text > " " Then
                                            'check if this must be a new line
                                            If t.Name = colPK Then
                                                If text <> lastPK Then
                                                    WS_REPORT.Cells(Line, 1).Select
                                                    DoEvents
                                                    If lastPK <> "x" Then
                                                        For y = 1 To list_values.Count
                                                            If Not IsEmpty(list_values(CStr(y))) Then
                                                                WS_REPORT.Cells(Line, y) = list_values(CStr(y))
                                                            End If
                                                        Next
                                                    End If
                                                    list_values.RemoveAll
                                                    lastPK = text
                                                    Line = Line + 1
                                                End If
                                            End If
                                            
                                            
                                            'fill in the list_values with only values that are unique
                                            newValue = True
                                            For y = reportColumnNo To reportColumnNo + multiColumns - 1
                                                If list_values(CStr(y)) > " " Then
                                                    If text = list_values(CStr(y)) Then
                                                        newValue = False
                                                    End If
                                                End If
                                            Next
                                            'if found a new value enter it in list_values, unless there is no space any more
                                            If newValue Then
                                                For y = reportColumnNo To reportColumnNo + multiColumns - 1
                                                    If list_values(CStr(y)) = "" And text > " " Then
                                                        list_values(CStr(y)) = text
                                                        text = ""
                                                    End If
                                                Next
                                            End If
                                        End If
                                        
                                    End If
                                Next
                            Next
                            reportColumnNo = reportColumnNo + multiColumns
    
                        End If
                    Next
                End If
            Next
        Else
            MsgBox "error: found query does not start with ""prefix"" or with ""select"""
        End If
        
        ' do the last one
        For x = 1 To list_values.Count
            If Not IsEmpty(list_values(CStr(x))) Then
                WS_REPORT.Cells(Line, x) = list_values(CStr(x))
            End If
        Next
    End If
    
    Set htmQry = Nothing

End Sub
Function transform(ByVal text As String) As String

    transform = Trim(text)
    transform = Replace(transform, "&lt;", "<")
    transform = Replace(transform, "&gt;", ">")
    
End Function

Function getHtm(ByVal sparql_text As String, Optional ByVal altEndpoint As String) As Object

Dim oXML, aErr, sURL, ok
Set getHtm = CreateObject("htmlFile")
If altEndpoint > " " Then
    sURL = altEndpoint & "?query=" & URLEncode(sparql_text) & "&output=xml"
Else
    sURL = gEndpoint & "?query=" & URLEncode(sparql_text) & "&output=xml"
End If
 
ok = 24 'do this for a full 2 minutes before raising the alarm
While ok > 0
    Set oXML = CreateObject("msxml2.XMLHTTP.6.0")
    On Error Resume Next
     oXML.Open "GET", sURL, False
     aErr = Array(Err.Number, Err.Description)
    On Error GoTo 0
     If 0 = aErr(0) Then
       On Error Resume Next
        oXML.Send
        aErr = Array(Err.Number, Err.Description)
       On Error GoTo 0
        Select Case True
          Case 0 <> aErr(0)
            ok = ok - 1 'try connecting many times
            
            If ok <= 0 Then
                MsgBox "send failed:" & vbCrLf & aErr(0) & vbCrLf & aErr(1) & vbCrLf & oXML.responsetext & vbCrLf & gEndpoint & vbCrLf & sparql_text
            End If
            DoEvents
            Application.Wait (Now + TimeValue("0:00:05")) 'calmly press the refresh button every 5 seconds
            
          Case 200 = oXML.Status
            ok = 0
            'MsgBox sURL, oXML.Status, oXML.statusText
          Case Else
            MsgBox "further work needed:"
            MsgBox sURL, oXML.Status, oXML.statusText
        End Select
     Else
        MsgBox "open failed:", aErr(0), aErr(1)
     End If
     getHtm.body.innerHTML = oXML.responsetext
Wend

'    'get the sparql done
'    Set getHtm = CreateObject("htmlFile")
'
'    On Error GoTo error:
'
'    With CreateObject("msxml2.xmlhttp")
'        .Open "GET", gEndpoint & "?query=" & sparql_text & "&output=xml", False
'        .Send
'        getHtm.body.innerHTML = .responsetext
'    End With
    
'    If InStr(Mid(getHtm.body.innerHTML, 1, 30), "Error") > 0 Then
'    'error
'        MsgBox getHtm.body.innerHTML
'    End If
'
'    Exit Function

'error:
' MsgBox "Error 404 (address not found): " & gEndpoint & vbCrLf & "-----------------------------" & vbCrLf & sparql_text

    
    
End Function
Public Function URLEncode(ByVal StringVal As String) As String

  Dim StringLen As Long: StringLen = Len(StringVal)
  Dim SpaceAsPlus As Boolean
    SpaceAsPlus = True
    


  If StringLen > 0 Then
    ReDim result(StringLen) As String
    Dim i As Long, CharCode As Integer
    Dim Char As String, Space As String

    If SpaceAsPlus Then Space = "+" Else Space = "%20"

    For i = 1 To StringLen
      Char = Mid$(StringVal, i, 1)
      CharCode = Asc(Char)
      Select Case CharCode
        Case 97 To 122, 65 To 90, 48 To 57, 45, 46, 95, 126
          result(i) = Char
        Case 32
          result(i) = Space
        
        Case 0 To 15
          result(i) = "%0" & Hex(CharCode)
        Case Else
          result(i) = "%" & Hex(CharCode)
      End Select
    Next i
    URLEncode = Join(result, "")
  End If
End Function

Function namespacehandler(obj)
    namespacehandler = obj
    namespacehandler = Replace(namespacehandler, "http://rds.posccaesar.org/2008/02/OWL/ISO-15926-2_2003#", "dm:")
    namespacehandler = Replace(namespacehandler, "http://www.w3.org/1999/02/22-rdf-syntax-ns#", "rdf:")
    namespacehandler = Replace(namespacehandler, "http://www.w3.org/2000/01/rdf-schema#", "rdfs:")
    namespacehandler = Replace(namespacehandler, "http://www.w3.org/2002/07/owl#", "owl:")
    namespacehandler = Replace(namespacehandler, "http://purl.org/dc/elements/1.1/", "dc:")
    namespacehandler = Replace(namespacehandler, "http://data.posccaesar.org/rdl/", "pcardl:")
    namespacehandler = Replace(namespacehandler, "http://www.w3.org/2004/02/skos/core#", "skos:")
    namespacehandler = Replace(namespacehandler, "http://data.15926.org/meta/", "meta:")
    namespacehandler = Replace(namespacehandler, "http://rds.15926.org/2008/02/OWL/ISO-15926-2_2003#", "dm:")
    namespacehandler = Replace(namespacehandler, "http://data.15926.org/dm/", "dm:")
    namespacehandler = Replace(namespacehandler, "http://data.posccaesar.org/dm/", "dm:")
    namespacehandler = Replace(namespacehandler, "http://data.posccaesar.org/lci/", "pcalci:")
    namespacehandler = Replace(namespacehandler, "http://data.15926.org/lci/", "lci:")
    
    namespacehandler = Replace(namespacehandler, "http://data.15926.org/aisi/", "aisi:")
    namespacehandler = Replace(namespacehandler, "http://data.15926.org/ansi/", "ansi:")
    namespacehandler = Replace(namespacehandler, "http://data.15926.org/api/", "api:")
    namespacehandler = Replace(namespacehandler, "http://data.15926.org/asme/", "asme:")
    namespacehandler = Replace(namespacehandler, "http://data.15926.org/astm/", "astm:")
    namespacehandler = Replace(namespacehandler, "http://data.15926.org/atex/", "atex:")
    namespacehandler = Replace(namespacehandler, "http://data.15926.org/bs/", "bs:")
    namespacehandler = Replace(namespacehandler, "http://data.15926.org/cfihos/", "cfihos:")
    namespacehandler = Replace(namespacehandler, "http://data.15926.org/din/", "din:")
    namespacehandler = Replace(namespacehandler, "http://data.15926.org/dm/", "dm:")
    namespacehandler = Replace(namespacehandler, "http://data.15926.org/dnv/", "dnv:")
    namespacehandler = Replace(namespacehandler, "http://data.15926.org/edm/", "edm:")
    namespacehandler = Replace(namespacehandler, "http://data.15926.org/en/", "en:")
    namespacehandler = Replace(namespacehandler, "http://data.15926.org/eu/", "eu:")
    namespacehandler = Replace(namespacehandler, "http://data.15926.org/iec/", "iec:")
    namespacehandler = Replace(namespacehandler, "http://data.15926.org/iso/", "iso:")
    namespacehandler = Replace(namespacehandler, "http://data.15926.org/mss/", "mss:")
    namespacehandler = Replace(namespacehandler, "http://data.15926.org/nace/", "nace:")
    namespacehandler = Replace(namespacehandler, "http://data.15926.org/nec/", "nec:")
    namespacehandler = Replace(namespacehandler, "http://data.15926.org/nema/", "nema:")
    namespacehandler = Replace(namespacehandler, "http://data.15926.org/nfpa/", "nfpa:")
    namespacehandler = Replace(namespacehandler, "http://data.15926.org/nor/", "nor:")
    namespacehandler = Replace(namespacehandler, "http://data.15926.org/rdl/", "rdl:")
    namespacehandler = Replace(namespacehandler, "http://data.15926.org/sae/", "sae:")
    namespacehandler = Replace(namespacehandler, "http://data.15926.org/tema/", "tema:")
    namespacehandler = Replace(namespacehandler, "http://data.15926.org/twr/", "twr:")
    namespacehandler = Replace(namespacehandler, "http://data.15926.org/uns/", "uns:")
    namespacehandler = Replace(namespacehandler, "http://data.15926.org/wits/", "wits:")

    
End Function

