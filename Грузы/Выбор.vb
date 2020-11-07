Public Class Выбор
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim f As New Отчет
        f.Кли()
        Me.Close()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim f As New Отчет
        f.Перевоз()
        Me.Close()
    End Sub

    Private Sub Выбор_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
End Class