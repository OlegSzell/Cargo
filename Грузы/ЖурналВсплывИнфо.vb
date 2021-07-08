Public Class ЖурналВсплывИнфо
    Private D As New Grid2ЖурналClass
    Private Загр As String
    Private Выгр As String
    Private Став As String
    Private Доп As String
    Sub New(_D As Grid2ЖурналClass)
        D = _D
        Using db As New dbAllDataContext(_cn3)
            Dim ms = (From x In db.ЖурналКлиентДаты
                      Where x.ID = D.IdКодЖурнДаты
                      Select New With {.загр = x.ДатаЗагрузки, .выгр = x.ДатаДоставки, .ct = x.Ставка, .dop = x.ДопУсловия}).FirstOrDefault()
            Загр = ms.загр
            Выгр = ms.выгр
            Став = ms.ct
            Доп = ms.dop
        End Using




        ' Этот вызов является обязательным для конструктора.
        InitializeComponent()
        Timer1.Interval = 1000
        Timer1.Enabled = True
        ProgressBar1.Maximum = 40
        ProgressBar1.Step = 2

        ' Добавить код инициализации после вызова InitializeComponent().

    End Sub
    Private Sub ЖурналВсплывИнфо_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Label1.Text = D.Номер
        Timer1.Start()

        If D IsNot Nothing Then

            TextBox4.Text = D.Клиент
            TextBox5.Text = D.Страны
            RichTextBox1.Text = D.Груз

            TextBox1.Text = Загр
            TextBox2.Text = Выгр
            TextBox3.Text = Став
            RichTextBox2.Text = Доп

        End If
    End Sub
    Private ip As Integer = 0
    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick

        ProgressBar1.PerformStep()
        ip += 1
        If ip = 20 Then
            Timer1.Stop()
            Close()
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Timer1.Stop()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Timer1.Interval = 1000
        Timer1.Start()
    End Sub


    Private Async Sub txt1UpdAsync(dg As String, txt As String)
        Await Task.Run(Sub() txt1Upd(dg, txt))
    End Sub
    Private Sub txt1Upd(dg As String, txt As String)
        Using db As New dbAllDataContext(_cn3)
            Dim f = db.ЖурналКлиентДаты.Where(Function(x) x.ID = D.IdКодЖурнДаты).FirstOrDefault()
            If f IsNot Nothing Then
                Select Case txt
                    Case 1
                        f.ДатаЗагрузки = dg
                    Case 2
                        f.ДатаДоставки = dg
                    Case 3
                        f.Ставка = dg
                    Case Else
                        f.ДопУсловия = dg
                End Select

                db.SubmitChanges()
                Dim mo As New AllUpd
                mo.ЖурналКлиентДатыAll()
            End If
        End Using
    End Sub

    Private Sub TextBox1_LostFocus(sender As Object, e As EventArgs) Handles TextBox1.LostFocus
        If Not Загр = TextBox1.Text Then
            txt1UpdAsync(TextBox1.Text, 1)
        End If

    End Sub



    Private Sub TextBox2_LostFocus(sender As Object, e As EventArgs) Handles TextBox2.LostFocus
        If Not Выгр = TextBox2.Text Then
            txt1UpdAsync(TextBox2.Text, 2)
        End If
    End Sub

    Private Sub TextBox3_LostFocus(sender As Object, e As EventArgs) Handles TextBox3.LostFocus
        If Not Став = TextBox3.Text Then
            txt1UpdAsync(TextBox3.Text, 3)
        End If
    End Sub

    Private Sub RichTextBox2_LostFocus(sender As Object, e As EventArgs) Handles RichTextBox2.LostFocus
        If Not Доп = RichTextBox2.Text Then
            txt1UpdAsync(RichTextBox2.Text, 4)
        End If
    End Sub
End Class