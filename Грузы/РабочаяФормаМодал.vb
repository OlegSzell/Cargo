Option Explicit On
Imports System.Data.OleDb

Public Class РабочаяФормаМодал
    Private Sub Удаление()
        If IDГруз = Nothing Or IDПеревоза = Nothing Then
            Exit Sub
        End If
        Dim strsql As String = "DELETE * FROM ИтогГрузПеревоз WHERE IDГруз=" & IDГруз & " AND IDПеревоз=" & IDПеревоза & ""
        Updates(strsql)
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        If TextBox1.Text = "" Then
            Удаление()
            Me.Close()
            РабочаяФорма.Клик()
            РабочаяФорма.Grid2Refresh(1)
            Exit Sub
        End If

        Dim strsql As String
        If Label1.Text = "0" Then
            strsql = "INSERT INTO ИтогГрузПеревоз(IDПеревоз, IDГруз, Примечание, Дата) VALUES(" & IDПеревоза & ",
" & IDГруз & ", '" & TextBox1.Text & "','" & TextBox2.Text & "')"
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
            РабочаяФорма.Клик()
            РабочаяФорма.Grid2Refresh(1)
            TextBox1.Text = ""
            Me.Close()

        Else
            strsql = "UPDATE ИтогГрузПеревоз SET Примечание='" & TextBox1.Text & "' WHERE IDПеревоз=" & IDПеревоза & " AND IDГруз=" & IDГруз & ""
            Updates(strsql)
            MessageBox.Show("Изменено!")
            РабочаяФорма.Клик()
            РабочаяФорма.Grid2Refresh(1)
            TextBox1.Text = ""
            Me.Close()
        End If

    End Sub

    Private Sub РабочаяФормаМодал_Load(sender As Object, e As EventArgs) Handles Me.Load
        TextBox2.Text = FormatDateTime(Now.Date)
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        '        Dim strsql As String = "UPDATE ИтогГрузПеревоз SET Примечание='" & TextBox1.Text & "', Дата='" & TextBox2.Text & "' 
        'WHERE IDПеревоз= " & IDПеревоза & " AND IDГруз=" & IDГруз & ""

        '        Dim c As New OleDbCommand
        '        c.Connection = conn
        '        c.CommandText = strsql
        '        Try
        '            c.ExecuteNonQuery()
        '        Catch ex As Exception
        '            MessageBox.Show("С Перевозчиком не велись переговоры по этому грузу!" & vbCrLf & "Нажмите кнопку 'Сохранить'")
        '            Exit Sub
        '        End Try
        '        MessageBox.Show("Изменено!")
        '        Me.Close()
        '        TextBox1.Text = ""

        ПеревозВПутиДоб.ShowDialog()

    End Sub
End Class