Public Class РейсВыборГодаНовый
    Private lst As New List(Of String)
    Public Выбор As String = Nothing
    Sub New(ByVal _lst As List(Of String))
        lst = _lst
        ' Этот вызов является обязательным для конструктора.
        InitializeComponent()

        ' Добавить код инициализации после вызова InitializeComponent().

    End Sub

    Private Sub РейсВыборГодаНовый_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        For Each b In lst
            ListBox1.Items.Add(b)
        Next



    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Close()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If ListBox1.SelectedItem Is Nothing Then
            MessageBox.Show("Выберите год!", Рик)
            Return
        End If
        Выбор = ListBox1.SelectedItem.ToString
        Close()
    End Sub
End Class