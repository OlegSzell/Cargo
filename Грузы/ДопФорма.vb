Option Explicit On
Imports System.Data.OleDb
Public Class ДопФорма
    Public Num As Integer
    Private txt1 As String
    Public ОбнвлExcel As Boolean = False
    Sub New(ByVal _num As Integer, ByVal _txt1 As String)
        Num = _num
        txt1 = _txt1
        ' Этот вызов является обязательным для конструктора.
        InitializeComponent()

        ' Добавить код инициализации после вызова InitializeComponent().

    End Sub

    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown

        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            TextBox1.Text = Replace(TextBox1.Text, ".", ",")
            Dim a As Double = CDbl(TextBox1.Text)
            TextBox2.Text = Math.Round((100 - a), 2)
            TextBox4.Text = Math.Round(CType((txt1), Integer) * a / 100)
            TextBox3.Text = CType(CType((IIf(txt1 = String.Empty, 0, txt1)), Integer) - CType((TextBox4.Text), Integer), String)
            MaskedTextBox2.Focus()
        End If
    End Sub

    Private Sub ДопФорма_Closed(sender As Object, e As EventArgs) Handles Me.Closed
        TextBox1.Text = String.Empty
        TextBox2.Text = String.Empty
        TextBox3.Text = String.Empty
        TextBox4.Text = String.Empty
        TextBox5.Text = String.Empty
        TextBox6.Text = String.Empty
        TextBox10.Text = String.Empty
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
        TextBox1.Text = String.Empty
        TextBox2.Text = String.Empty
        TextBox3.Text = String.Empty
        TextBox4.Text = String.Empty
        TextBox5.Text = String.Empty
        TextBox6.Text = String.Empty
        TextBox10.Text = String.Empty
        TextBox7.Text = String.Empty

        MaskedTextBox1.Clear()
        MaskedTextBox2.Clear()


    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Cursor = Cursors.WaitCursor
        Dim mo As New AllUpd
        Using db As New dbAllDataContext()
            Dim var = db.РейсыКлиента.Where(Function(x) x.НомерРейса = Num).Select(Function(x) x).FirstOrDefault()
            If var IsNot Nothing Then
                var.ПоИнотерр = TextBox3.Text
                var.ПоТеррРБ = TextBox4.Text
                var.ПоТеррРБПроц = TextBox1.Text
                var.ПоИнотерПроц = TextBox2.Text
                var.ДатаАкта = MaskedTextBox2.Text
                var.НомерСМР = TextBox10.Text
                var.ЗаявкаКлиента = TextBox6.Text
                var.НомерЗаявки = TextBox5.Text
                var.ДатаЗаявки = MaskedTextBox1.Text

                If TextBox7.Visible = True Then
                    var.ОплатаПоКурсуКурс = TextBox7.Text
                End If
                db.SubmitChanges()
                mo.РейсыКлиентаAllAsync()
            End If
        End Using


        MessageBox.Show("Данные внесены в базу!", Рик)
        If MessageBox.Show("Изменить данные в файле эксель?", Рик, MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then

            ОбнвлExcel = True
            Me.Cursor = Cursors.Default
            MessageBox.Show("Данные в файле эксель изменены!", Рик)
            Me.Close()
            Exit Sub
        End If
        Me.Cursor = Cursors.Default
        Me.Close()

    End Sub

    Private Sub ДопФорма_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim mo As New AllUpd
        Do While AllClass.РейсыКлиента Is Nothing
            mo.РейсыКлиентаAll()
        Loop
        Do While AllClass.РейсыПеревозчика Is Nothing
            mo.РейсыПеревозчикаAll()
        Loop

        Dim f1 = (From x In AllClass.РейсыКлиента
                  Where x.НомерРейса = Num
                  Select x).FirstOrDefault()

        Dim f2 = (From x In AllClass.РейсыПеревозчика
                  Where x.НомерРейса = Num
                  Select x).FirstOrDefault()

        TextBox4.Text = f1.ПоТеррРБ
        TextBox3.Text = f1.ПоИнотерр
        TextBox1.Text = f1.ПоТеррРБПроц
        TextBox2.Text = f1.ПоИнотерПроц

        TextBox6.Text = f1.ЗаявкаКлиента
        TextBox5.Text = f1.НомерЗаявки
        MaskedTextBox1.Text = f1.ДатаЗаявки

        MaskedTextBox2.Text = f1.ДатаАкта
        TextBox10.Text = f1.НомерСМР

        If f1.ОплатаПоКурсу = "True" Then
            Label5.Visible = True
            TextBox7.Visible = True
            TextBox7.Text = f1.ОплатаПоКурсуКурс
        Else
            Label5.Visible = False
            TextBox7.Visible = False

        End If



    End Sub
End Class