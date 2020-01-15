Option Explicit On
Imports System.Data.OleDb
Public Class ИзменПеревоз
    Public Da As New OleDbDataAdapter 'Адаптер
    Public Ds As New DataSet 'Пустой набор записей
    Dim tbl As New DataTable
    Dim cb As OleDb.OleDbCommandBuilder
    Dim Рик As String = "ООО Рикманс"
    Dim загр, strsql As String
    Dim ind As Boolean = False

    Private Sub ИзменПеревоз_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        conn = New OleDbConnection
        conn.ConnectionString = ConString
        Try
            conn.Open()
        Catch ex As Exception
            MessageBox.Show("Не подключен диск U")
        End Try
        TextBox5.Enabled = False
        TextBox4.Enabled = False

        дан()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()
    End Sub

    Public Sub дан()
        strsql = "SELECT * FROM ПЕревозчикиБаза WHERE ID=" & idtabl & ""
        Dim c6 As New OleDbCommand
        c6.Connection = conn
        c6.CommandText = strsql
        Dim ds As New DataTable
        Dim da6 As New OleDbDataAdapter(c6)
        da6.Fill(ds)

        TextBox1.Text = ds.Rows(0).Item(1).ToString
        TextBox2.Text = ds.Rows(0).Item(2).ToString
        TextBox3.Text = ds.Rows(0).Item(3).ToString
        TextBox4.Text = ds.Rows(0).Item(6).ToString
        TextBox5.Text = ds.Rows(0).Item(5).ToString
        TextBox6.Text = ds.Rows(0).Item(4).ToString
        TextBox7.Text = ds.Rows(0).Item(12).ToString
        TextBox8.Text = ds.Rows(0).Item(11).ToString
        TextBox9.Text = ds.Rows(0).Item(10).ToString
        TextBox10.Text = ds.Rows(0).Item(9).ToString
        TextBox11.Text = ds.Rows(0).Item(8).ToString
        TextBox12.Text = ds.Rows(0).Item(7).ToString
        TextBox13.Text = ds.Rows(0).Item(14).ToString
        TextBox14.Text = ds.Rows(0).Item(13).ToString




    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked = True Then
            CheckBox2.Checked = False

            СтраныПерЛист.Show()
        Else
            СтраныПерЛист.Close()
        End If
    End Sub

    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox2.CheckedChanged
        If CheckBox2.Checked = True Then
            CheckBox1.Checked = False

            РегионыПерЛист.Show()
        Else
            РегионыПерЛист.Close()
        End If
    End Sub
    Private Sub Ok()
        If MessageBox.Show("Сохранить изменения?", Рик, MessageBoxButtons.OKCancel, MessageBoxIcon.Question) = DialogResult.Cancel Then
            Exit Sub
            Me.Close()
        End If

        strsql = "UPDATE ПеревозчикиБаза SET [Форма собственности]='" & TextBox1.Text & "', [Наименование фирмы]='" & TextBox2.Text & "',
[Контактное лицо]='" & TextBox3.Text & "',[Телефоны]='" & TextBox6.Text & "',[Города]='" & TextBox12.Text & "',
[ADR]='" & TextBox11.Text & "',[Кол-во авто]='" & TextBox10.Text & "',[Вид_авто]='" & TextBox9.Text & "',
[Тоннаж]='" & TextBox8.Text & "',[Объем]='" & TextBox7.Text & "',[Ставка]='" & TextBox14.Text & "',[Примечание]='" & TextBox13.Text & "'
WHERE ID=" & idtabl & ""
        Dim c25 As New OleDbCommand
        c25.Connection = conn
        c25.CommandText = strsql
        c25.ExecuteNonQuery()

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Ok()
        MessageBox.Show("Данные изменены!", Рик)
        If pr2 = 0 Then
            ПоискПеревозчиков.ОбнКом1()
        End If

        Me.Close()
    End Sub

    Private Sub TextBox5_DoubleClick(sender As Object, e As EventArgs) Handles TextBox5.DoubleClick
        TextBox5.Enabled = True

    End Sub
End Class