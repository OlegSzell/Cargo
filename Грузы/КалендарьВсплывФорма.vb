Public Class КалендарьВсплывФорма
    Private _D As String
    Private _F As String
    Sub New(d As String, f As String)
        _F = f
        _D = d
        ' Этот вызов является обязательным для конструктора.
        InitializeComponent()
        Label1.Text = Now
        Label2.Text = _D
        RichTextBox1.Text = _F
        ' Добавить код инициализации после вызова InitializeComponent().

    End Sub
    Private Sub КалендарьВсплывФорма_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
End Class