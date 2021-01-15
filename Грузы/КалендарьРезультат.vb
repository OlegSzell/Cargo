Public Class КалендарьРезультат
    Public Property Rezul As String
    Public D As String = Nothing
    Sub New()

        ' Этот вызов является обязательным для конструктора.
        InitializeComponent()

        ' Добавить код инициализации после вызова InitializeComponent().

    End Sub
    Sub New(ByVal _d As String)
        D = _d
        ' Этот вызов является обязательным для конструктора.
        InitializeComponent()

        ' Добавить код инициализации после вызова InitializeComponent().

    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Close()
    End Sub

    Private Sub КалендарьРезультат_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If D IsNot Nothing Then
            RichTextBox1.Text = D
        End If

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If MessageBox.Show("Сохранить результат?", Рик, MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.No Then
            Return
        End If

        If RichTextBox1.Text.Length = 0 Then
            MessageBox.Show("Не заполнено поле текста!", Рик)
            Return
        End If

        Rezul = RichTextBox1.Text

        MessageBox.Show("Данные приняты!", Рик)
        Close()

    End Sub
End Class