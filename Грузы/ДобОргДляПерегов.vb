Public Class ДобОргДляПерегов
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim strsql As String = "SELECT ДатаПереговоров FROM ПереговорыКлиент WHERE Клиент='" & RichTextBox1.Text & "'"
        Dim ds As DataTable = Selects(strsql)
        If errds = 1 Then
            новая()
            MessageBox.Show("Организация добавлена!", Рик)
            RichTextBox1.Text = ""
            RichTextBox2.Text = ""
            ПереговорыКлиент.комб1()
            Me.Close()
        Else
            MessageBox.Show("Данная организация уже существует!", Рик)
        End If

    End Sub
    Private Sub новая()
        Dim strsql As String = "INSERT INTO ПереговорыКлиент(Клиент,КонтДанные) VALUES('" & RichTextBox1.Text & "','" & RichTextBox2.Text & "')"
        Updates(strsql)
    End Sub
End Class