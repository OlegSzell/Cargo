Public Class ПредоплатаКлиент
    Private Sub ПредоплатаКлиент_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox4.Text = ""

    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        If TextBox1.Text = "" Or TextBox2.Text = "" Or TextBox4.Text = "" Then
            MessageBox.Show("Заполните все данные!")
            Exit Sub
        End If
        'If ComboBox1.Text = "" Or ComboBox2.Text = "" Then
        '    MessageBox.Show("Заполните все данные!")
        '    Exit Sub
        'End If

        'If ComboBox1.Text = "БанкПоДок" Then
        '    df = "банковских дней"
        'ElseIf ComboBox1.Text = "КалПоДок" Then
        '    df = "календарных дней"
        'ElseIf ComboBox1.Text = "БанкПоВыг" Then
        '    df = "банковских дней"
        'Else
        '    df = "календарных дней"
        'End If

        'If ComboBox2.Text = "БанкПоДок" Then
        '    df1 = "банковских дней"
        'ElseIf ComboBox2.Text = "КалПоДок" Then
        '    df1 = "календарных дней"
        'ElseIf ComboBox2.Text = "БанкПоВыг" Then
        '    df1 = "банковских дней"
        'Else
        '    df1 = "календарных дней"
        'End If

        Рейс.ЧастичнаяОплатаКлиент = "Предоплата в размере " & TextBox1.Text & "% в течении " & TextBox2.Text & " дня(дней) после загрузки транспортного средства,
остаток в течении " & TextBox4.Text & " дня(дней) после получения документов по перевозке (CMR, счет, акт выполненных работ)."

        MessageBox.Show("Данные сохранены!", Рик)
        Me.Close()
    End Sub

    Private Sub RichTextBox1_TextChanged(sender As Object, e As EventArgs) Handles RichTextBox1.TextChanged

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Рейс.ЧастичнаяОплатаКлиент = ""
        MessageBox.Show("Данные внесены, предоплата отменена!", Рик)
        Рейс.Button7.BackColor = Color.LightBlue
        Me.Close()
    End Sub
End Class