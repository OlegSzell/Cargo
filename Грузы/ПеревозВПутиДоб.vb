Option Explicit On
Imports System.Data.OleDb
Public Class ПеревозВПутиДоб
    Private Sub ПеревозВПутиДоб_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim strsql As String = "INSERT INTO ПеревозчикиВПути(IDПеревозчика, Перевозчик, ДатаВыгр,ГдеВыгр,Примечание)
VALUES(" & IDПеревоза & ", '" & NameПеревоза & "', '" & MaskedTextBox1.Text & "','" & RichTextBox1.Text & "','" & TextBox5.Text & "')"
        Dim c As New OleDbCommand
        c.Connection = conn
        c.CommandText = strsql
        Try
            c.ExecuteNonQuery()
        Catch ex As Exception
            MessageBox.Show("Не Сохранено!")
            Exit Sub
        End Try
        MessageBox.Show("Сохранено!")
        Me.Close()
        MaskedTextBox1.Text = ""
        RichTextBox1.Text = ""
        TextBox5.Text = ""
    End Sub
End Class