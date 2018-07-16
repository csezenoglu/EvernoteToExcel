Option Explicit

Sub OutputNotesXML()

Dim iRow As Long

Close #1
With ActiveSheet
    'For iRow = 2 To 2
    Open ThisWorkbook.Path & "\evernote-import.enex" For Output As #1
        Print #1, "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "UTF-8" & Chr(34) & "?>"
        Print #1, "<!DOCTYPE en-export SYSTEM " & Chr(34) & "http://xml.evernote.com/pub/evernote-export.dtd" & Chr(34) & ">"
        Print #1, "<en-export export-date=" & Chr(34) & "20120202T073208Z" & Chr(34) & " application=" & Chr(34) & "Evernote/Windows" & Chr(34) & " version=" & Chr(34) & "4.x" & Chr(34) & ">"
    For iRow = 2 To .Cells(.Rows.Count, "A").End(xlUp).row
        Print #1, "<note><title>"
        Print #1, .Cells(iRow, "A").Value 'Title
        Print #1, "</title><content><![CDATA[<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "UTF-8" & Chr(34) & "?>"
        Print #1, "<!DOCTYPE en-note SYSTEM " & Chr(34) & "http://xml.evernote.com/pub/enml2.dtd" & Chr(34) & ">"
        Print #1, "<en-note style=" & Chr(34) & "word-wrap: break-word; -webkit-nbsp-mode: space; -webkit-line-break: after-white-space;" & Chr(34) & ">"
        Print #1, CBr(.Cells(iRow, "B").Value) 'Note
        Print #1, "</en-note>]]></content><created>"
        Print #1, .Cells(iRow, "D").Text 'Created Date in Evernote Time Format...
        'To get the evernote time, first convert your time to Zulu/UTC time.
        'Put this formula in Column D: =C2+TIME(6,0,0) where 6 is the hours UTC is ahead of you.
        'Then right click on your date column, select format, then select custom. Use this custom code: yyyymmddThhmmssZ
        Print #1, "</created><updated>201206025T000001Z</updated></note>"
    Next iRow
    Print #1, "</en-export>"
    Close #1
    
End With

End Sub

Function CBr(val) As String
    'parse hard breaks into to HTML breaks
    CBr = Replace(val, Chr(13), "")
    CBr = Replace(CBr, "&", "&amp;")
End Function

'I modified this code from Marty Zigman's post here: http://blog.prolecto.com/2012/01/31/importing-excel-data-into-evernote-without-a-premium-account/
' This will read ENEX file (Evernote export file) into Excel worksheet
Sub ReadBCNotesXML()
    Dim fdgOpen As FileDialog
    Dim fp As Integer
    Dim i As Integer
    Dim DataLine As String, WholeFileContent As String
    Dim RE As Object, allMatches As Object
    Set RE = CreateObject("vbscript.regexp")
    
    Set fdgOpen = Application.FileDialog(msoFileDialogOpen)
    With fdgOpen
        .Filters.Add "Evernote files", "*.enex", 1
        .Title = "Please open Evernote file..."
        .InitialFileName = "."
        .InitialView = msoFileDialogViewDetails
        .Show
    End With
    ' MsgBox fdgOpen.SelectedItems(1)
    fp = FreeFile()
    WholeFileContent = ""
    Open fdgOpen.SelectedItems(1) For Input As #fp
        WholeFileContent = Input$(LOF(fp), fp)
    Close #fp
    
    ' Removing CR&LF line endings
    WholeFileContent = Replace(WholeFileContent, Chr(10), "")
    WholeFileContent = Replace(WholeFileContent, Chr(13), "")
    
    
    ' First line
    Worksheets(1).Cells(1, 1) = "Note Title"
    Worksheets(1).Cells(1, 2) = "Name"
    Worksheets(1).Cells(1, 3) = "Title"
    Worksheets(1).Cells(1, 4) = "Organization"
    Worksheets(1).Cells(1, 5) = "Email"
    Worksheets(1).Cells(1, 6) = "Phone"
    Worksheets(1).Cells(1, 7) = "Notes"
    Worksheets(1).Cells(1, 8) = "Created"
    Worksheets(1).Cells(1, 9) = "Updated"
    
    
    ' Filter for title
    RE.Pattern = "<title>(.*?)<\/title>"
    RE.IgnoreCase = True
    RE.Global = True
    RE.MultiLine = True
    Set allMatches = RE.Execute(WholeFileContent)
    For i = 0 To allMatches.Count - 1
        Worksheets(1).Cells(2 + i, 1) = UTF8_Decode(allMatches(i).submatches(0))
    Next
    
    
'    ' Filter for content
    RE.Pattern = "<content>(.*?)<\/content>"
    'RE.IgnoreCase = True
    'RE.Global = True
    Set allMatches = RE.Execute(WholeFileContent)
    For i = 0 To allMatches.Count - 1
        ParseBCContent allMatches(i).submatches(0), i
    Next

    
    ' Filter for created
    RE.Pattern = "<created>(.*?)<\/created>"
    'RE.IgnoreCase = True
    'RE.Global = True
    Set allMatches = RE.Execute(WholeFileContent)
    For i = 0 To allMatches.Count - 1
        Worksheets(1).Cells(2 + i, 8) = UTF8_Decode(allMatches(i).submatches(0))
    Next
    ' Filter for updated
    RE.Pattern = "<updated>(.*?)<\/updated>"
    'RE.IgnoreCase = True
    'RE.Global = True
    Set allMatches = RE.Execute(WholeFileContent)
    For i = 0 To allMatches.Count - 1
        Worksheets(1).Cells(2 + i, 9) = UTF8_Decode(allMatches(i).submatches(0))
    Next
    ' Free
    Set RE = Nothing
    Set allMatches = Nothing
End Sub

'This function is created espacially for parsing Evernote Business Card notes.
Function ParseBCContent(WholeFileContent As String, i As Integer)

Dim http As Object, html As HTMLDocument, topics As Object, titleElem As Object, detailsElem As Object, topic As Object
Dim nEmail, nPhone, k As Integer, Item As Object, Email, Phone As String

'Dim WholeFileContent As String

Set html = New MSHTML.HTMLDocument

html.body.innerHTML = WholeFileContent

If InStr(1, WholeFileContent, "x-evernote:contact", vbTextCompare) Then

Set topics = html.getElementsByTagName("span")

k = 1
For Each titleElem In topics

    If InStr(1, titleElem.Style.cssText, "x-evernote: display-as", vbTextCompare) Then
        Worksheets(1).Cells(2 + i, 2) = UTF8_Decode(titleElem.innerText)
    End If
    
    If InStr(1, titleElem.Style.cssText, "x-evernote: contact-title", vbTextCompare) Then
        Worksheets(1).Cells(2 + i, 3) = UTF8_Decode(titleElem.innerText)
    End If
    
    If InStr(1, titleElem.Style.cssText, "x-evernote: contact-org", vbTextCompare) Then
        Worksheets(1).Cells(2 + i, 4) = UTF8_Decode(titleElem.innerText)
    End If
    
    k = k + 1
Next
Set topics = Nothing


Set topics = html.getElementsByTagName("div")
k = 1
nEmail = 0
nPhone = 0

For Each topic In topics


If InStr(1, topic.Style.cssText, "x-evernote: email", vbTextCompare) Then
    If nEmail > 0 Then
        Email = Email & vbCrLf & topic.getElementsByTagName("div")(1).innerText
        nEmail = nEmail + 1
    Else
        Email = topic.getElementsByTagName("div")(1).innerText
        nEmail = nEmail + 1
    End If

ElseIf InStr(1, topic.Style.cssText, "x-evernote: phone", vbTextCompare) Then
    
    If nPhone > 0 Then
        Phone = Phone & vbCrLf & topic.getElementsByTagName("span")(1).innerText & ": " & topic.getElementsByTagName("span")(2).innerText
        nPhone = nPhone + 1
    Else
        Phone = topic.getElementsByTagName("span")(1).innerText & ": " & topic.getElementsByTagName("span")(2).innerText
        nPhone = nPhone + 1
    End If

ElseIf InStr(1, topic.Style.cssText, "x-evernote: note-body", vbTextCompare) Then

    Worksheets(1).Cells(2 + i, 7) = UTF8_Decode(topic.innerText)

End If
    
    k = k + 1
Next

    Worksheets(1).Cells(2 + i, 5) = Email
    Worksheets(1).Cells(2 + i, 6) = Phone

Else
    'Error
    MsgBox ("This is not a Business Card")
End If



End Function
' http://p2p.wrox.com/vbscript/29099-unicode-utf-8-system-text-utf8encoding-vba.html
Function UTF8_Decode(ByVal sStr As String)
    Dim l As Long, sUTF8 As String, iChar As Integer, iChar2 As Integer
    For l = 1 To Len(sStr)
        iChar = Asc(Mid(sStr, l, 1))
        If iChar > 127 Then
            If Not iChar And 32 Then ' 2 chars
            iChar2 = Asc(Mid(sStr, l + 1, 1))
            sUTF8 = sUTF8 & ChrW$(((31 And iChar) * 64 + (63 And iChar2)))
            l = l + 1
        Else
            Dim iChar3 As Integer
            iChar2 = Asc(Mid(sStr, l + 1, 1))
            iChar3 = Asc(Mid(sStr, l + 2, 1))
            sUTF8 = sUTF8 & ChrW$(((iChar And 15) * 16 * 256) + ((iChar2 And 63) * 64) + (iChar3 And 63))
            l = l + 2
        End If
            Else
            sUTF8 = sUTF8 & Chr$(iChar)
        End If
    Next l
    UTF8_Decode = sUTF8
End Function

Function StripTags(inString As String) As String
    Dim RE As Object, allMatches As Object
    Set RE = CreateObject("vbscript.regexp")
    ' Keeping enters
    ' inString = Replace(inString, "</div>", " ")
    ' Removing   other <tag>-s
    ' RE.Pattern = "<[^>]+>"
    RE.Pattern = "<\S[^>]*>"
    RE.IgnoreCase = True
    RE.Global = True
    StripTags = RE.Replace(inString, "")
    ' Cleaning up strange things
    StripTags = Replace(StripTags, "]]>", "")
    StripTags = Replace(StripTags, "&apos;", "'")
    StripTags = Replace(StripTags, "&nbsp;", "  ")
    ' Free
    Set RE = Nothing
    Set allMatches = Nothing
End Function




