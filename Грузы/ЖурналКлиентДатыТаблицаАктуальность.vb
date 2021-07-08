Public Class ЖурналКлиентДатыТаблицаАктуальность
    Private ID As Integer
    Public Rez As String = Nothing
    Sub New(ByVal _Id As Integer)

        ID = _Id

        ' Этот вызов является обязательным для конструктора.
        InitializeComponent()

        ' Добавить код инициализации после вызова InitializeComponent().

    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Close()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If CheckBox1.Checked = False And CheckBox2.Checked = False Then
            MessageBox.Show("Выберите вариант!", Рик)
            Return
        End If

        If CheckBox1.Checked = True Then
            Rez = "Актуально! " & vbCrLf & Now
        Else
            Rez = "Закрыт груз! " & vbCrLf & Now
        End If

        UpdAsync(Rez)

        MessageBox.Show("Данные приняты!", Рик)
        Close()

    End Sub
    Private Async Sub UpdAsync(ByVal rs As String)
        Await Task.Run(Sub() Upd(rs))
    End Sub
    Private Sub Upd(ByVal rs As String)
        Using db As New dbAllDataContext(_cn3)
            Dim f = db.ЖурналКлиентДаты.Where(Function(x) x.ID = ID).FirstOrDefault
            If f IsNot Nothing Then
                f.Состояние = rs
                db.SubmitChanges()
                Dim mo As New AllUpd
                mo.ЖурналКлиентДатыAll()
            End If
        End Using

    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        CheckBox2.Checked = False
    End Sub

    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox2.CheckedChanged
        CheckBox1.Checked = False
    End Sub
End Class