Public Class ПредоплатаПеревозчик
    Private Sub Предоплата_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ComboBox1.Items.Clear()
        Dim d1() As String

        d1 = {"БанкПоДок", "КалПоДок", "БанкПоВыг", "КалПоВыг"}
        ComboBox1.Items.AddRange(d1)
        ComboBox2.Items.AddRange(d1)
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox4.Text = ""
        ComboBox1.Text = ""
        ComboBox2.Text = ""
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim df, df1 As String
        If TextBox1.Text = "" Or TextBox2.Text = "" Or TextBox4.Text = "" Then
            MessageBox.Show("Заполните все данные!")
            Exit Sub
        End If
        If ComboBox1.Text = "" Or ComboBox2.Text = "" Then
            MessageBox.Show("Заполните все данные!")
            Exit Sub
        End If

        If ComboBox1.Text = "БанкПоДок" Then
            df = "банковских дней"
        ElseIf ComboBox1.Text = "КалПоДок" Then
            df = "календарных дней"
        ElseIf ComboBox1.Text = "БанкПоВыг" Then
            df = "банковских дней"
        Else
            df = "календарных дней"
        End If

        If ComboBox2.Text = "БанкПоДок" Then
            df1 = "банковских дней"
        ElseIf ComboBox2.Text = "КалПоДок" Then
            df1 = "календарных дней"
        ElseIf ComboBox2.Text = "БанкПоВыг" Then
            df1 = "банковских дней"
        Else
            df1 = "календарных дней"
        End If

        Рейс.ЧастичнаяОплатаПеревозчик = "Оплата в размере " & TextBox1.Text & "% в течении " & TextBox2.Text & "" & df1 & " после получения оформленных должным образом документов по перевозке (CMR, счет, акт выполненных работ),
остаток в течении " & TextBox4.Text & "" & df & " после получения оформленных должным образом документов по перевозке (CMR, счет, акт выполненных работ)."




        MessageBox.Show("Данные сохранены!", Рик)
        Me.Close()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Рейс.ЧастичнаяОплатаПеревозчик = ""
        MessageBox.Show("Данные внесены, частичная оплата отменена!", Рик)
        Рейс.Button2.BackColor = Color.LightBlue
        Me.Close()
    End Sub
End Class