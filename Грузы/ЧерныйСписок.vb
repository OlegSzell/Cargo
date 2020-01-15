Option Explicit On
Imports System.Data.OleDb

Public Class ЧерныйСписок
    Public g As Integer = 0
    Private Sub ЧерныйСписок_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        conn = New OleDbConnection
        conn.ConnectionString = ConString
        Try
            conn.Open()
        Catch ex As Exception
            MessageBox.Show("Не подключен диск U")
        End Try

        refgrid1()


    End Sub

    Public Sub refgrid1()
        Dim strsql As String = "SELECT DISTINCT Организация FROM ЧерныйСписок"
        Dim ds As DataTable = Selects(strsql)

        Me.ComboBox1.AutoCompleteCustomSource.Clear()
        Me.ComboBox1.Items.Clear()
        For Each r As DataRow In ds.Rows
            Me.ComboBox1.AutoCompleteCustomSource.Add(r.Item(0).ToString())
            Me.ComboBox1.Items.Add(r(0).ToString)
        Next
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        Dim strsql As String = "SELECT * FROM ЧерныйСписок WHERE Организация='" & ComboBox1.Text & "'"
        Dim ds As DataTable = Selects(strsql)

        Grid1.DataSource = ds
        Grid1.Columns(0).Visible = False
        Grid1.Columns(1).Width = 200
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If TextBox1.Text = "" Then
            MessageBox.Show("Введите цифры для поиска")
            Exit Sub
        End If
        Dim strsql As String = "SELECT * FROM ЧерныйСписок WHERE Примечание Like '%" & TextBox1.Text & "%'"
        Dim ds As DataTable = Selects(strsql)

        If errds = 1 Then
            MessageBox.Show("Нет такого номреа в Черном списке!", Рик)
            ds.Clear()
        Else
            Grid1.DataSource = ds
            Grid1.Columns(0).Visible = False
            Grid1.Columns(1).Width = 200
            Grid1.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        End If

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        ДобЧернСписок.ShowDialog()
    End Sub

    Private Sub Grid1_CellMouseDoubleClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles Grid1.CellMouseDoubleClick
        Dim d As Integer = Grid1.CurrentRow.Cells(0).Value
        Dim strsql As String = "SELECT * FROM ЧерныйСписок WHERE Код=" & d & ""
        Dim ds As DataTable = Selects(strsql)
        g = 1
        ДобЧернСписок.TextBox1.Text = ""
        ДобЧернСписок.RichTextBox1.Text = ""
        ДобЧернСписок.TextBox1.Text = ds.Rows(0).Item(2).ToString
        ДобЧернСписок.RichTextBox1.Text = ds.Rows(0).Item(1).ToString
        ДобЧернСписок.ShowDialog()
    End Sub
End Class