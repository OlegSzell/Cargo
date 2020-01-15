Public Class Паролик
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If TextBox1.Text = "oleg1389925" Then
            ПодтверждениеПароля = True
            Me.Close()
        Else
            MessageBox.Show("Введите пароль правильно!")
            Exit Sub

        End If

    End Sub

    Private Sub Паролик_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ПодтверждениеПароля = False
    End Sub

    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown
        If e.KeyCode = Keys.Enter Then
            Button1.PerformClick()
        End If
    End Sub
End Class