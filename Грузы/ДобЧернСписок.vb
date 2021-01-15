

Public Class ДобЧернСписок
    Private ReadOnly D As Boolean = False
    Private Fd As New ЧерныйСписок
    Public Rez As String = Nothing
    Sub New(_d As Boolean, Optional _Fd As ЧерныйСписок = Nothing)
        D = _d
        Fd = _Fd
        ' Этот вызов является обязательным для конструктора.
        InitializeComponent()
        If Fd IsNot Nothing Then
            TextBox1.Text = Fd.Примечание
            RichTextBox1.Text = Fd.Организация
        End If

        ' Добавить код инициализации после вызова InitializeComponent().

    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim rch1, rch2 As String
        rch1 = Trim(RichTextBox1.Text)
        rch2 = Trim(TextBox1.Text)
        Dim mo As New AllUpd
        If D = True Then
            Using db As New dbAllDataContext
                Dim f = db.ЧерныйСписок.Where(Function(x) x.Код = Fd.Код).FirstOrDefault()
                If f IsNot Nothing Then
                    f.Примечание = rch2
                    db.SubmitChanges()
                End If
            End Using

            mo.ЧерныйСписокAllAsync()
            Rez = rch2
            'Dim strsql2 As String = "UPDATE ЧерныйСписок SET Примечание = '" & Trim(TextBox1.Text) & "' WHERE Организация='" & RichTextBox1.Text & "'"
            'Updates3(strsql2)
            MessageBox.Show("Данные обновлены!", Рик)
        Else
            Using db As New dbAllDataContext()
                Dim f As New ЧерныйСписок
                f.Организация = rch1
                f.Примечание = rch2
                db.ЧерныйСписок.InsertOnSubmit(f)
                db.SubmitChanges()

            End Using
            'Dim strsql As String = "INSERT INTO ЧерныйСписок(Организация,Примечание) VALUES('" & rch1 & "','" & rch2 & "')"
            'Updates3(strsql)
            mo.ЧерныйСписокAllAsync()
            MessageBox.Show("Данные добавлены в базу!", Рик)
        End If

        RichTextBox1.Text = ""
        TextBox1.Text = ""

        Me.Close()

    End Sub



End Class