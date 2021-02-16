Public Class ЖурналДобавитьГрузДопДанные
    Public Rezult As String = Nothing




    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If RichTextBox1.Text.Length > 0 Then
            Rezult = Trim(RichTextBox1.Text)
        End If
        MessageBox.Show("Данные приняты!", Рик)
    End Sub
End Class