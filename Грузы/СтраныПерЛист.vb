Option Explicit On
Imports System.Data.OleDb

Public Class СтраныПерЛист
    Dim strsql As String
    Dim Рик As String = "ООО Рикманс"

    Private Sub СтраныПерЛист_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        conn = New OleDbConnection
        conn.ConnectionString = ConString
        Try
            conn.Open()
        Catch ex As Exception
            MessageBox.Show("Не подключен диск U")
        End Try

        strsql = "SELECT Страна FROM Страна ORDER BY Страна"
        Dim c6 As New OleDbCommand
        c6.Connection = conn
        c6.CommandText = strsql
        Dim ds As New DataTable
        Dim da6 As New OleDbDataAdapter(c6)
        da6.Fill(ds)

        For Each r As DataRow In ds.Rows
            ListBox1.Items.Add(r(0).ToString)
        Next

    End Sub
    Private Sub Interbaza()
        If MessageBox.Show("Изменить данные?", Рик, MessageBoxButtons.OKCancel) = DialogResult.Cancel Then

            Exit Sub
        End If
        If ListBox1.SelectedIndex = -1 Then

            Exit Sub
        End If
        Dim country As String
        Dim i As Integer
        For i = 0 To ListBox1.SelectedItems.Count - 1
            country = ListBox1.SelectedItems(i).ToString & ", " & country
        Next

        strsql = ""
        strsql = "UPDATE ПеревозчикиБаза SET [Страны перевозок]='" & country & "' WHERE ПеревозчикиБаза.ID = " & idtabl & ""
        Dim c25 As New OleDbCommand
        c25.Connection = conn
        c25.CommandText = strsql
        c25.ExecuteNonQuery()


    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Interbaza()
        ИзменПеревоз.CheckBox1.Checked = False
        If pr2 = 0 Then
            ПоискПеревозчиков.ОбнКом1()
        End If
        ИзменПеревоз.дан()
        Me.Close()
    End Sub

    Private Sub СтраныПерЛист_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        ИзменПеревоз.CheckBox1.Checked = False
    End Sub
End Class