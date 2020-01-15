Public Class ДляСкайпаРабочаяФорма
    Private Sub ДляСкайпаРабочаяФорма_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        RichTextBox1.BackColor = Color.LightBlue
        RichTextBox2.BackColor = Color.LightGreen
        '        Dim ds As DataTable = Selects3(StrSql:="SELECT ДляСкайпа FROM ГрузыКлиентов WHERE Организация='" & ДанныеДляВставкиСкайпа.Item1 & "' AND 
        'ГородВыгрузки='" & ДанныеДляВставкиСкайпа.Item2 & "' AND Груз='" & ДанныеДляВставкиСкайпа.Item3 & "' AND СтранаЗагрузки='" & ДанныеДляВставкиСкайпа.Item4 & "'")
        Dim ds As DataTable = Selects3(StrSql:="SELECT ДляСкайпа FROM ГрузыКлиентов WHERE Организация='" & ДанныеДляВставкиСкайпа.Item1 & "' AND 
ГородВыгрузки='" & ДанныеДляВставкиСкайпа.Item2 & "' AND Груз='" & ДанныеДляВставкиСкайпа.Item3 & "' AND СтранаЗагрузки='" & ДанныеДляВставкиСкайпа.Item4 & "'")

        If ds.Rows(0).Item(0).ToString = "" Then
            RichTextBox1.Text = "Нет данных"
        Else
            RichTextBox1.Text = ds.Rows(0).Item(0).ToString
        End If

        RichTextBox2.Text = "Добрый день.
BY>D
Груз готов
Лельчицы(Гом.обл) - D 12277 + D 87700
затаможка на месте, растаможка (Swiecko(PL) или Görlitz (D))
брикет на подддонах, до 23 т- тент - 950е,
оплата безнал, евро, предложите авто."
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim f As Integer = ДанныеДляВставкиСкайпа.Item5
        Dim ds As String = "UPDATE ГрузыКлиентов SET ДляСкайпа='" & RichTextBox1.Text & "' WHERE Код=" & f & ""
        'Updates(ds)
        Updates3(ds)
        RichTextBox2.Text = ""
        RichTextBox1.Text = ""
        Me.Close()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        RichTextBox1.Text = RichTextBox2.Text
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim ds As DataTable = Selects3(StrSql:="SELECT Груз FROM ГрузыКлиентов WHERE Код=" & ДанныеДляВставкиСкайпа.Item5 & "")
        RichTextBox3.Text = ds.Rows(0).Item(0).ToString

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        RichTextBox1.Text = ""
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        My.Computer.Clipboard.SetText(RichTextBox1.Text)
    End Sub
End Class