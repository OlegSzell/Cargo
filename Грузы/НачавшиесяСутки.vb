Public Class НачавшиесяСутки
    Public Rez As String = Nothing

    Private Sub НачавшиесяСутки_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        RichTextBox1.Text = "8. Нормативное время на погрузку составляет 24 часа, на выгрузку 24 часа. " _
& "Штраф за простой производится согласно карты простоя, за каждые начавшиеся сутки, в размере 100 Евро на территории стран Европы или 50 Евро на территории стран СНГ. " _
& "Выходные и праздничные дни при расчёте нормативного времени не учитываются и не считаются простойными."

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If RichTextBox1.Text.Length = 0 Then
            MessageBox.Show("Заполните данные!", Рик)
            Return
        End If
        MessageBox.Show("Данные приняты!", Рик)
        Rez = Trim(RichTextBox1.Text)
        Close()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Close()
    End Sub
End Class