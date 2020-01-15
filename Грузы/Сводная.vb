Public Class Сводная
    Private Sub Сводная_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
    Private Sub rev()
        Dim strsql As String = "SELECT РейсыКлиента.НомерРейса, РейсыКлиента.НазвОрганизации, РейсыКлиента.СтоимостьФрахта, РейсыКлиента.Валюта, РейсыКлиента.ВалютаПлатежа, РейсыКлиента.СрокОплаты
FROM РейсыКлиента WHERE Год LIKE * #" & ComboBox1.Text & "#"
    End Sub
End Class