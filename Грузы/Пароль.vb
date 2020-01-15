Option Explicit On
Imports System.Data.OleDb
Public Class Пароль
    Dim strsql As String
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        ClickOK()

    End Sub
    Private Sub ClickOK()
        strsql = "SELECT Парол FROM Пароли WHERE Логин='" & ComboBox1.Text & "'"
        Dim c1 As New OleDbCommand With {
            .Connection = conn,
            .CommandText = strsql
        }
        Dim ds1 As New DataTable
        Dim da1 As New OleDbDataAdapter(c1)
        da1.Fill(ds1)

        If ds1.Rows(0).Item(0).ToString = TextBox1.Text Then
            'MessageBox.Show("Пароль принят!")
            Экспедитор = ComboBox1.Text
            Me.Close()
            TextBox1.Text = ""
            CheckBox1.Checked = False
        Else
            MessageBox.Show("Пароль НЕ принят! Повторите")
        End If
    End Sub
    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked = True Then
            TextBox1.PasswordChar = ""
        Else
            TextBox1.PasswordChar = "*"
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim strsql As String = "INSERT INTO Пароли(Логин,Парол) VALUES('" & ComboBox1.Text & "','" & TextBox1.Text & "')"
        Dim c As New OleDbCommand
        c.Connection = conn
        c.CommandText = strsql
        c.ExecuteNonQuery()
        MessageBox.Show("Ваши данные зарегистрированы!")
        ComboBox1.Text = ""
        TextBox1.Text = ""
        Combo1()
    End Sub
    Private Sub Combo1()
        Dim StrSql1 As String
        StrSql1 = "SELECT Логин FROM Пароли ORDER BY Логин"
        Dim c1 As New OleDbCommand With {
            .Connection = conn,
            .CommandText = StrSql1
        }
        Dim ds1 As New DataTable
        Dim da1 As New OleDbDataAdapter(c1)
        da1.Fill(ds1)
        Me.ComboBox1.Items.Clear()
        Me.ComboBox1.AutoCompleteCustomSource.Clear()

        For Each r As DataRow In ds1.Rows
            Me.ComboBox1.AutoCompleteCustomSource.Add(r.Item(0).ToString())
            Me.ComboBox1.Items.Add(r(0).ToString)
        Next
    End Sub
    Private Sub Пароль_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Экспедитор = ""
        conn = New OleDbConnection
        conn.ConnectionString = ConString
        Try
            conn.Open()
        Catch ex As Exception
            MessageBox.Show("Не подключен диск U")
        End Try
        Combo1()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Me.Close()
        MDIParent1.Close()
    End Sub

    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            ClickOK()
        End If
    End Sub
End Class