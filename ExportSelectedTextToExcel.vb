Sub ExportSelectedTextToExcel()    Dim objExcel As Object
    Dim WordDoc As Document    Dim CurrRange As Range
    Dim objExcelWorkbook As Workbook    Dim objExcelWorksheet As Worksheet
    Dim TargetCellAddress As String
    On Error GoTo ErrHandler
    ' Проверяем наличие выделения    
If Selection.Type <> wdSelectionNormal Then
        MsgBox "Пожалуйста, выделите текст для экспорта.", vbExclamation        
Exit Sub
    End If
    ' Открываем документ Word    
Set WordDoc = ActiveDocument
    ' Получаем выделенный текст
    Set CurrRange = Selection.Range
    ' Запрашиваем у пользователя адрес ячейки для экспорта    TargetCellAddress = InputBox("Укажите адрес ячейки для экспорта:", "Выбор ячейки", "A1")
        If Trim(TargetCellAddress) = "" Then
        MsgBox "Вы не указали адрес ячейки. Операция отменена.", vbInformation        Exit Sub
    End If
    ' Создаем объект Excel и новую книгу    Set objExcel = CreateObject("Excel.Application")
    Set objExcelWorkbook = objExcel.Workbooks.Add    
Set objExcelWorksheet = objExcelWorkbook.Sheets(1)
    ' Копируем выделенный текст в указанную ячейку Excel
    objExcelWorksheet.Range(TargetCellAddress).Value = CurrRange.FormattedText
    ' Делаем Excel видимым    objExcel.Visible = True

ErrExit:    ' Освобождаем ресурсы
    Set objExcelWorkbook = Nothing    Set objExcel = Nothing
    Set WordDoc = Nothing    Set CurrRange = Nothing
    Exit Sub
ErrHandler:    ' Обработка ошибки
    MsgBox "Ошибка: " & Err.Description, vbCritical
    Resume ErrExit
End Sub