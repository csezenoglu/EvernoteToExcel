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
    For iRow = 2 To .Cells(.Rows.Count, "A").End(xlUp).Row
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
Sub ReadNotesXML()
    Dim fdgOpen As FileDialog
    Dim fp As Integer
    Dim i As Integer
    Dim DataLine As String, WholeFileContent As String
    Dim RE As Object, allMatches As Object
    Set RE = CreateObject("vbscript.regexp")
    
    Set fdgOpen = Application.FileDialog(msoFileDialogOpen)
    With fdgOpen
        .Filters.Add "Evernote files", "*.enex", 1
        .TITLE = "Please open Evernote file..."
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
    ' Worksheets(1).Cells(5, 5) = WholeFileContent
    ' First line
    Worksheets(1).Cells(1, 1) = "Title"
    Worksheets(1).Cells(1, 2) = "Content"
    Worksheets(1).Cells(1, 3) = "Created"
    Worksheets(1).Cells(1, 4) = "Updated"
    ' Filter for title
    RE.Pattern = "<title>(.*?)<\/title>"
    RE.IgnoreCase = True
    RE.Global = True
    RE.MultiLine = True
    Set allMatches = RE.Execute(WholeFileContent)
    For i = 0 To allMatches.Count - 1
        Worksheets(1).Cells(2 + i, 1) = allMatches(i).submatches(0)
    Next
    ' Filter for content
    RE.Pattern = "<content>(.*?)<\/content>"
    'RE.IgnoreCase = True
    'RE.Global = True
    Set allMatches = RE.Execute(WholeFileContent)
    For i = 0 To allMatches.Count - 1
        Worksheets(1).Cells(2 + i, 2) = StripTags(allMatches(i).submatches(0))
    Next
    ' Filter for created
    RE.Pattern = "<created>(.*?)<\/created>"
    'RE.IgnoreCase = True
    'RE.Global = True
    Set allMatches = RE.Execute(WholeFileContent)
    For i = 0 To allMatches.Count - 1
        Worksheets(1).Cells(2 + i, 3) = allMatches(i).submatches(0)
    Next
    ' Filter for updated
    RE.Pattern = "<updated>(.*?)<\/updated>"
    'RE.IgnoreCase = True
    'RE.Global = True
    Set allMatches = RE.Execute(WholeFileContent)
    For i = 0 To allMatches.Count - 1
        Worksheets(1).Cells(2 + i, 4) = allMatches(i).submatches(0)
    Next
    ' Free
    Set RE = Nothing
    Set allMatches = Nothing
End Sub

Function StripTags(inString As String) As String
    Dim RE As Object, allMatches As Object
    Set RE = CreateObject("vbscript.regexp")
    ' Keeping enters
    inString = Replace(inString, "</div>", " ")
    ' Removing   other <tag>-s
    RE.Pattern = "<[^>]+>"
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

Sub r2i()
    Dim lLastRow As Long
    Dim lLastCol As Long
    Dim rgLast As Range
    Dim rgSrc As Range
    Dim rgDst As Range
    Dim i, j As Integer
    Dim RE As Object, allMatches As Object
    Set RE = CreateObject("vbscript.regexp")
    Dim m As String
    
    Set rgLast = Range("A1").SpecialCells(xlCellTypeLastCell)
    lLastRow = rgLast.Row
    lLastCol = rgLast.Column
    
    Set rgSrc = Range(Cells(2, 2), Cells(lLastRow, 2))
    Set rgDst = Range(Cells(2, 1), Cells(lLastRow, 1))
    RE.Pattern = "\((.*?)\)"
    RE.IgnoreCase = True
    RE.Global = True
    
    For i = 1 To rgSrc.Count
        ' Getting stuff in brackets
        Set allMatches = RE.Execute(rgSrc.Cells(i, 1))
        m = ""
        If allMatches.Count > 0 Then
            For j = 0 To allMatches.Count - 1
                If allMatches.Count = 1 Then
                  m = allMatches(j).submatches(0)
                Else
                  m = m & allMatches(j).submatches(0) & ";"
                End If
            Next
            rgDst.Cells(i, 1) = m
        Else
            m = rgDst.Cells(i, 1)
            rgDst.Cells(i, 1) = rgSrc.Cells(i, 1)
            rgSrc.Cells(i, 1) = m
        End If
    Next
    Set RE = Nothing
    Set allMatches = Nothing
End Sub

Sub i2r()
    Dim lLastRow As Long
    Dim lLastCol As Long
    Dim rgLast As Range
    Dim rgSrc As Range
    Dim rgDst As Range
    Dim i As Integer
    Dim RE As Object, allMatches As Object
    Set RE = CreateObject("vbscript.regexp")
    Dim m As String
    
    Set rgLast = Range("A1").SpecialCells(xlCellTypeLastCell)
    lLastRow = rgLast.Row
    lLastCol = rgLast.Column
    
    Set rgSrc = Range(Cells(2, 2), Cells(lLastRow, 2))
    Set rgDst = Range(Cells(2, 1), Cells(lLastRow, 1))
    RE.Pattern = "^(.*?)\s+\(.*"
    RE.IgnoreCase = True
    RE.Global = True
    
    For i = 1 To rgSrc.Count
        ' Getting stuff in brackets
        Set allMatches = RE.Execute(rgSrc.Cells(i, 1))
        If allMatches.Count > 0 Then
            rgDst.Cells(i, 1) = allMatches(0).submatches(0)
        Else
            m = rgDst.Cells(i, 1)
            rgDst.Cells(i, 1) = rgSrc.Cells(i, 1)
            rgSrc.Cells(i, 1) = m
        End If
    Next
    Set RE = Nothing
    Set allMatches = Nothing
End Sub


