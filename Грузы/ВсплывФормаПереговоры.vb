Public Class ВсплывФормаПереговоры
    Dim dgh As DataTable
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
    Public Sub Загр(ByVal ds As DataTable)
        RichTextBox1.Text = ""
        RichTextBox2.Text = ""
        RichTextBox3.Text = ""
        RichTextBox1.Text = ds.Rows(0).Item(1).ToString
        RichTextBox2.Text = ds.Rows(0).Item(5).ToString
        RichTextBox3.Text = Strings.Left(ds.Rows(0).Item(4).ToString, 10)
        dgh = ds
    End Sub
    Public Sub БолееОдногоНапомин(ByVal int As Integer)
        f = Nothing
        f = dgh.Rows(int - 1).Item(0)
        RichTextBox1.Text = ""
        RichTextBox2.Text = ""
        RichTextBox3.Text = ""
        RichTextBox4.Text = ""
        RichTextBox1.Text = dgh.Rows(int - 1).Item(1).ToString
        RichTextBox2.Text = dgh.Rows(int - 1).Item(5).ToString
        RichTextBox3.Text = Strings.Left(dgh.Rows(int - 1).Item(4).ToString, 10)
        RichTextBox4.Text = dgh.Rows(int - 1).Item(8).ToString
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Try
            Dim strsql As String = "UPDATE ПереговорыКлиент SET ТекстНапоминания='" & RichTextBox2.Text & "', ОЧемДоговорВсплывФорма='" & RichTextBox4.Text & "', 
ДатаОчемДоговорилис='" & RichTextBox3.Text & "'
WHERE Код=" & f & ""
            Updates3(strsql)
            Dim df As DataTable = MDIParent1.Провер()
            dgh = df
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
        MessageBox.Show("Данные внесены!", Рик)

    End Sub


End Class