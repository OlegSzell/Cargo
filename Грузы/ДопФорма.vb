Option Explicit On
Imports System.Data.OleDb
Public Class ДопФорма


    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown

        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            TextBox1.Text = Replace(TextBox1.Text, ".", ",")
            Dim a As Double = CDbl(TextBox1.Text)
            TextBox2.Text = Math.Round((100 - a), 2)
            TextBox4.Text = Math.Round(CType((Рейс.TextBox1.Text), Integer) * a / 100)
            TextBox3.Text = CType(CType((Рейс.TextBox1.Text), Integer) - CType((TextBox4.Text), Integer), String)
            MaskedTextBox2.Focus()
        End If
    End Sub

    Private Sub ДопФорма_Closed(sender As Object, e As EventArgs) Handles Me.Closed
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        TextBox6.Text = ""
        TextBox10.Text = ""
        MaskedTextBox1.Text = ""
        MaskedTextBox2.Text = ""

    End Sub

    Private Sub TextBox6_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox6.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.TextBox5.Focus()
        End If
    End Sub

    Private Sub TextBox5_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox5.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.MaskedTextBox1.Focus()
        End If
    End Sub

    Private Sub MaskedTextBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles MaskedTextBox2.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.TextBox10.Focus()
        End If
    End Sub

    Private Sub TextBox10_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox10.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.Button1.Focus()
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        TextBox6.Text = ""
        TextBox10.Text = ""
        TextBox7.Text = ""

        MaskedTextBox1.Clear()
        MaskedTextBox2.Clear()


    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Cursor = Cursors.WaitCursor
        Using db As New dbAllDataContext()
            Dim var = db.РейсыКлиента.Where(Function(x) x.НомерРейса = Рейс.НомРес).Select(Function(x) x).FirstOrDefault()
            If var IsNot Nothing Then

                var.ПоИнотерр = TextBox3.Text
                var.ПоТеррРБ = TextBox4.Text
                var.ДатаАкта = MaskedTextBox2.Text
                var.НомерСМР = TextBox10.Text
                var.ЗаявкаКлиента = TextBox6.Text
                var.НомерЗаявки = TextBox5.Text
                var.ДатаЗаявки = MaskedTextBox1.Text

                If TextBox7.Visible = True Then
                    var.ОплатаПоКурсуКурс = TextBox7.Text
                End If
                db.SubmitChanges()

                End If
        End Using


        '        Dim strsql As String = "UPDATE РейсыКлиента SET ПоИнотерр='" & TextBox3.Text & "', ПоТеррРБ='" & TextBox4.Text & "',
        'ДатаАкта='" & MaskedTextBox2.Text & "', НомерСМР='" & TextBox10.Text & "', ЗаявкаКлиента='" & TextBox6.Text & "', НомерЗаявки='" & TextBox5.Text & "',
        'ДатаЗаявки='" & MaskedTextBox1.Text & "' 
        'WHERE НомерРейса=" & Рейс.НомРес & ""
        '        Updates3(strsql)
        MessageBox.Show("Данные внесены в базу!", Рик)
        If MessageBox.Show("Изменить данные в файле эксель?", Рик, MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
            Рейс.СлРейс = Nothing
            Рейс.СлРейс = Рейс.НомРес

            Рейс.ДокиОбновление()
            Me.Cursor = Cursors.Default
            MessageBox.Show("Данные в файле эксель изменены!", Рик)
            Me.Close()
            Exit Sub
        End If
        Me.Cursor = Cursors.Default
        Me.Close()

    End Sub

    Private Sub ДопФорма_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
End Class