Public Class КалендарьМодалноеОкно
    Dim df As String
    Private Sub КалендарьМодалноеОкно_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        RichTextBox1.Text = df
        Label1.Text = Now.ToShortTimeString
    End Sub
    Public Sub New(ByVal d As String)

        ' Этот вызов является обязательным для конструктора.
        InitializeComponent()
        df = d
        ' Добавить код инициализации после вызова InitializeComponent().

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        КалендарьПовтор = False
        Close()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        КалендарьПовтор = True
        Close()
    End Sub
End Class