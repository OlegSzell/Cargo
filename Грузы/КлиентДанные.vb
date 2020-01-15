Option Explicit On
Imports System.Data.OleDb
Public Class КлиентДанные
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If ComboBox1.Text = "" Or RichTextBox2.Text = "" Then
            MessageBox.Show("Нет данных для сохранения!", Рик)
            Exit Sub
        End If

        Dim strsql As String = "SELECT * FROM ГрузыКлиентов WHERE Организация='" & ComboBox1.Text & "'"
        Dim ds As DataTable = Selects3(strsql)

        If errds = 1 Then
            Dim strsql1 As String = "INSERT INTO ГрузыКлиентов(Организация,ОрганизКонтакт) VALUES('" & ComboBox1.Text & "', '" & RichTextBox2.Text & "')"
            Updates3(strsql1)
            MessageBox.Show("Новый клиент добавлен!", Рик)
        Else
            Dim strsql2 As String = "UPDATE ГрузыКлиентов SET ОрганизКонтакт='" & RichTextBox2.Text & "' WHERE Организация='" & ComboBox1.Text & "'"
            Updates3(strsql2)
            MessageBox.Show("Данные обновлены!", Рик)
        End If
        RichTextBox2.Text = ""
        ComboBox1.Text = ""
        Me.Close()
    End Sub

    Private Sub КлиентДанные_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim strsql As String = "SELECT DISTINCT Организация FROM ГрузыКлиентов"
        Dim ds As DataTable = Selects3(strsql)
        ComboBox1.AutoCompleteCustomSource.Clear()
        ComboBox1.Items.Clear()
        For Each r As DataRow In ds.Rows
            ComboBox1.AutoCompleteCustomSource.Clear()
            ComboBox1.Items.Add(r(0).ToString)
        Next

    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        Dim strsql As String = "SELECT ОрганизКонтакт FROM ГрузыКлиентов WHERE Организация='" & ComboBox1.Text & "'"
        Dim ds As DataTable = Selects3(strsql)
        RichTextBox2.Text = ""
        RichTextBox2.Text = ds.Rows(0).Item(0).ToString
    End Sub
End Class