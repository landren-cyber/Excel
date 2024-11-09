Sub CopyToWordAreas()
    Dim wdApp As Object
    Dim wdDoc As Object
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim rng As Range
    Dim strFile As String
    Dim strDocName As String
    Dim FinalRow As Long
    Dim areaCount As Integer
    Dim currentArea As Range

    ' File name
    strFile = ActiveWorkbook.FullName
    strDocName = Right(strFile, Len(strFile) - InStrRev(strFile, "."))

    ' Opening and creating a new document
    On Error Resume Next
    Set wdApp = GetObject(, "Word.Application")
    If wdApp Is Nothing Then
       Set wdApp = CreateObject("Word.Application")
           wdApp.Visible = True
       Set wdDoc = wdApp.Documents.Add
    Else
       Set wdDoc = wdApp.ActiveDocument
    End If
    On Error GoTo 0

    ' Getting data from Excel
    Set wb = ActiveWorkbook
    Set ws = wb.Sheets("Sheet1")
    Set rng = Selection

    ' Processing each selected range individually
    areaCount = rng.Areas.Count
    For Each currentArea In rng.Areas
        ' Adding paragraph break between ranges if not first one
        If areaCount > 1 And Not currentArea Is rng.Areas(1) Then
            wdDoc.Content.InsertParagraphAfter
        End If

        ' Pasting the table into Word
        currentArea.Copy
        wdDoc.Paragraphs.Add.Range.PasteExcelTable _
            LinkedToExcel:=False, _
            WordFormatting:=False, _
            RTF:=False
    Next currentArea

    ' Formatting font in all paragraphs
    With wdDoc.Paragraphs
        For i = 1 To wdDoc.Paragraphs.Count
            With wdDoc.Paragraphs(i).Range.Font
                .Name = "Times New Roman"
                .Size = 12
            End With
        Next i
    End With

    ' Removing extra spaces
    With wdDoc.Content.Find
        .ClearFormatting
        .Text = "^13"
        .Replacement.Text = " "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .Execute Replace:=wdReplaceAll
    End With

    ' Saving a document
    wdDoc.SaveAs "export.docx"

    ' Clearing the cache
    Set wdDoc = Nothing
    Set wdApp = Nothing
    Set ws = Nothing
    Set wb = Nothing
End Sub
