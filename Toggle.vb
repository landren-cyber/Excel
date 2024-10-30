Sub ToggleMacro()
    If ActiveSheet.Shapes("Кнопка1").ControlFormat.Value = 1 Then
        ActiveSheet.Shapes("Кнопка1").ControlFormat.Value = 0
        MsgBox "Макрос выключен."
    Else
        ActiveSheet.Shapes("Кнопка1").ControlFormat.Value = 1
        MsgBox "Макрос включен."
    End If
End Sub
