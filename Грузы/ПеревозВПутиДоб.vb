Option Explicit On
Imports System.Data.OleDb
Imports System.Data.SqlClient
Public Class ПеревозВПутиДоб
    Private Sub ПеревозВПутиДоб_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim strsql As String = "INSERT INTO ПеревозчикиВПути(IDПеревозчика, Перевозчик, ДатаВыгр,ГдеВыгр,Примечание)
VALUES(" & IDПеревоза & ", '" & NameПеревоза & "', '" & MaskedTextBox1.Text & "','" & RichTextBox1.Text & "','" & TextBox5.Text & "')"
        Dim conn As New SqlConnection
        conn.ConnectionString = ConString

        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If


        Dim c As New SqlCommand
        c.Connection = conn
        c.CommandText = strsql
        Try
            c.ExecuteNonQuery()
        Catch ex As Exception
            MessageBox.Show("Не Сохранено!")
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
            Exit Sub
        End Try
        MessageBox.Show("Сохранено!")
        If conn.State = ConnectionState.Open Then
            conn.Close()
        End If
        Me.Close()
        MaskedTextBox1.Text = ""
        RichTextBox1.Text = ""
        TextBox5.Text = ""
    End Sub

    Private Sub GroupBox2_Enter(sender As Object, e As EventArgs) Handles GroupBox2.Enter

    End Sub
End Class