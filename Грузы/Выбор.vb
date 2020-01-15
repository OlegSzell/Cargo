Public Class Выбор
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Отчет.Кли()
        Me.Close()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Отчет.Перевоз()
        Me.Close()
    End Sub

    Private Sub Выбор_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
End Class