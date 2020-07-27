Public Class ВсплывФормаПереговоры
    Dim dgh As List(Of ПереговорыКлиент)
    Dim f As Integer
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If CType(Label5.Text, Integer) = 1 Then Exit Sub
        Dim d As Integer = CType(Label5.Text, Integer)
        Label5.Text = d - 1
        БолееОдногоНапомин(d)
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If CType(Label5.Text, Integer) = CType(Label4.Text, Integer) Then Exit Sub
        Dim d As Integer = CType(Label5.Text, Integer)
        Label5.Text = d + 1
        БолееОдногоНапомин(d)
    End Sub
    Public Sub Загр(ByVal ds As List(Of ПереговорыКлиент))
        RichTextBox1.Text = ""
        RichTextBox2.Text = ""
        RichTextBox3.Text = ""
        RichTextBox1.Text = ds(0).Клиент
        RichTextBox2.Text = ds(0).ТекстНапоминания
        If ds(0).ДатаНапоминания IsNot Nothing Then
            RichTextBox3.Text = Strings.Left(ds(0).ДатаНапоминания.ToString, 10)
        End If

        dgh = ds
    End Sub
    Public Sub БолееОдногоНапомин(ByVal int As Integer)
        f = Nothing
        f = dgh(int - 1).Код
        RichTextBox1.Text = ""
        RichTextBox2.Text = ""
        RichTextBox3.Text = ""
        RichTextBox4.Text = ""
        RichTextBox1.Text = dgh(int - 1).Клиент
        RichTextBox2.Text = dgh(int - 1).ТекстНапоминания
        If dgh(int - 1).ДатаНапоминания IsNot Nothing Then
            RichTextBox3.Text = Strings.Left(dgh(int - 1).ДатаНапоминания.ToString, 10)
        End If

        RichTextBox4.Text = dgh(int - 1).ОЧемДоговорВсплывФорма
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Try
            Using db As New dbAllDataContext()
                Dim var = db.ПереговорыКлиент.Where(Function(x) x.Код = f).Select(Function(x) x).FirstOrDefault()
                If var IsNot Nothing Then
                    var.ТекстНапоминания = RichTextBox2.Text
                    var.ОЧемДоговорВсплывФорма = RichTextBox4.Text
                    var.ДатаОчемДоговорилис = RichTextBox3.Text
                    db.SubmitChanges()
                End If

                Dim var1 = db.ПереговорыКлиент.Where(Function(x) x.ДатаНапоминания = Now.Date And x.Экспедитор = Экспедитор).Select(Function(x) x).ToList()
                dgh = var1
            End Using


            '            Dim strsql As String = "UPDATE ПереговорыКлиент SET ТекстНапоминания='" & RichTextBox2.Text & "', ОЧемДоговорВсплывФорма='" & RichTextBox4.Text & "', 
            'ДатаОчемДоговорилис='" & RichTextBox3.Text & "'
            'WHERE Код=" & f & ""
            '            Updates3(strsql)
            '            Dim df As DataTable = MDIParent1.Провер()
            '            dgh = df
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
        MessageBox.Show("Данные внесены!", Рик)

    End Sub

End Class