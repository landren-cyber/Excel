Sub CopyToWordAndAppend()
    Dim wdApp As Object
    Dim wdDoc As Object
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim rng As Range
    Dim strFile As String
    Dim strDocName As String
    Dim FinalRow As Long
    
    ' File name
    strFile = ActiveWorkbook.FullName
    strDocName = Right(strFile, Len(strFile) - InStrRev(strFile, "."))
    
    ' Open file & create new document
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
    
    ' Get data from Excel
    Set wb = ActiveWorkbook
    Set ws = wb.Sheets("Sheet1")
    Set rng = Selection
    
    ' Checking for data availability
    If wdDoc.Paragraphs.Count > 0 Then
        ' Adding the data to the end of the document
        rng.Copy
        wdDoc.Paragraphs(wdDoc.Paragraphs.Count).Range.PasteExcelTable LinkedToExcel:=False, WordFormatting:=False, RTF:=False
    Else
        rng.Copy
        wdDoc.Paragraphs(1).Range.PasteExcelTable LinkedToExcel:=False, WordFormatting:=False, RTF:=False
    End If
    
    ' New font
    With wdDoc.Paragraphs(1).Range.Font
        .Name = "Times New Roman"
        .Size = 14
    End With
    
    ' Saving document
    wdDoc.SaveAs "export.docx"
    
    ' Clearing the cache
    Set wdDoc = Nothing
    Set wdApp = Nothing
    Set ws = Nothing
    Set wb = Nothing
    
End Sub

