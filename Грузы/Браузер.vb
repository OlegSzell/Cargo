Imports AxSHDocVw

Public Class Браузер
    Private gog As String = Nothing
    Private Sub Браузер_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        AxWebBrowser1.Navigate("google.com")
        gog = "https://www.google.com/search?q="
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        If TextBox1.Text.Length = 0 Then Return
        ProgressBar1.Value = ProgressBar1.Minimum
        AxWebBrowser1.Navigate2(gog & TextBox1.Text)
        ProgressBar1.Value = ProgressBar1.Maximum
    End Sub

    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown
        If e.KeyCode = Keys.Enter Then
            If TextBox1.Text.Length > 0 Then
                AxWebBrowser1.Navigate(TextBox1.Text)
                e.SuppressKeyPress = True
            End If
        End If
    End Sub

    Private Sub AxWebBrowser1_NavigateComplete2(sender As Object, e As DWebBrowserEvents2_NavigateComplete2Event) Handles AxWebBrowser1.NavigateComplete2
        Try
            TextBox1.Text = AxWebBrowser1.LocationURL.ToString
        Catch ex As Exception

        End Try
    End Sub
End Class