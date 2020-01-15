Option Explicit On
Imports System.ComponentModel
Imports System.Data.OleDb


Public Class ДобавитьПеревозчика
    Public Da As New OleDbDataAdapter 'Адаптер
    Public Ds As New DataSet 'Пустой набор записей
    Dim tbl As New DataTable
    Dim cb As OleDb.OleDbCommandBuilder
    Dim Рик As String = "ООО Рикманс"

    Private Sub ДобавитьПеревозчика_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If ДобПер = 0 Then
            Me.MdiParent = MDIParent1
            Me.WindowState = FormWindowState.Maximized
        Else
            Me.WindowState = FormWindowState.Normal
            Me.StartPosition = FormStartPosition.CenterParent

        End If

        Очист()

        conn = New OleDbConnection
        conn.ConnectionString = ConString
        Try
            conn.Open()
        Catch ex As Exception
            MessageBox.Show("Не подключен диск U")
        End Try
        Dim Год As Integer = Year(Now)

        Dim StrSql As String
        StrSql = "SELECT DISTINCT Страна FROM Страна ORDER BY Страна"
        Dim c As New OleDbCommand With {
            .Connection = conn,
            .CommandText = StrSql
        }
        Dim ds As New DataTable
        Dim da As New OleDbDataAdapter(c)
        da.Fill(ds)
        Me.ListBox1.Items.Clear()
        For Each r As DataRow In ds.Rows
            Me.ListBox1.Items.Add(r(0).ToString)
        Next

        Dim StrSql2 As String
        StrSql2 = "SELECT DISTINCT Регионы FROM РегионыРоссии ORDER BY Регионы"
        Dim c2 As New OleDbCommand With {
            .Connection = conn,
            .CommandText = StrSql2
        }
        Dim ds2 As New DataTable
        Dim da2 As New OleDbDataAdapter(c2)
        da2.Fill(ds2)
        Me.ListBox2.Items.Clear()
        For Each r As DataRow In ds2.Rows
            Me.ListBox2.Items.Add(r(0).ToString)
        Next


        Dim StrSql3 As String
        StrSql3 = "SELECT DISTINCT [Наименование фирмы] FROM ПеревозчикиБаза ORDER BY [Наименование фирмы]"
        Dim c3 As New OleDbCommand With {
            .Connection = conn,
            .CommandText = StrSql3
        }
        Dim ds3 As New DataTable
        Dim da3 As New OleDbDataAdapter(c3)
        da3.Fill(ds3)
        Me.ComboBox1.AutoCompleteCustomSource.Clear()
        Me.ComboBox1.Items.Clear()
        For Each r As DataRow In ds3.Rows
            Me.ComboBox1.AutoCompleteCustomSource.Add(r.Item(0).ToString())
            Me.ComboBox1.Items.Add(r(0).ToString)
        Next

        ComboBox1.Focus()



    End Sub
    Private Sub Очист()
        ComboBox1.Text = ""
        ComboBox1.Text = String.Empty
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox10.Text = ""
        TextBox5.Text = ""
        TextBox6.Text = ""
        TextBox4.Text = ""
        TextBox7.Text = ""
        TextBox8.Text = ""
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        insertbaza()
        Me.Close()
    End Sub
    Private Sub insertbaza()

        If ComboBox1.Text = "" Then
            MessageBox.Show("Выберите перевозчика!", Рик, MessageBoxButtons.OK)
            Exit Sub
        End If

        If TextBox3.Text = "" Then
            MessageBox.Show("Заполните номер телефона!", Рик, MessageBoxButtons.OK)
            Exit Sub
        End If

        If ListBox1.SelectedIndex = -1 Then
            MessageBox.Show("Выберите страну!", Рик, MessageBoxButtons.OK)
            Exit Sub
        End If

        Dim country As String
        Dim i As Integer
        For i = 0 To ListBox1.SelectedItems.Count - 1
            country = ListBox1.SelectedItems(i).ToString & ", " & country
        Next
        country = Strings.Left(country, country.Length - 2)
        Dim country2 As String
        Dim i2 As Integer
        If Not ListBox2.SelectedIndex = -1 Then
            For i2 = 0 To ListBox2.SelectedItems.Count - 1
                country2 = ListBox2.SelectedItems(i2).ToString & ", " & country2
            Next
            country2 = Strings.Left(country2, country2.Length - 2)
        Else
            country2 = ""
        End If

        Dim StrSql5 As String = "INSERT INTO ПеревозчикиБаза([Наименование фирмы],[Контактное лицо],Телефоны,[Страны перевозок],Регионы,ADR,[Кол-во авто],
[Вид_авто],Тоннаж,Объем,Ставка,Примечание)
VALUES('" & Me.ComboBox1.Text & "','" & Me.TextBox2.Text & "','" & Me.TextBox3.Text & "','" & country & "','" & country2 & "','" & Me.TextBox9.Text & "',
'" & Me.TextBox4.Text & "', '" & Me.TextBox10.Text & "','" & Me.TextBox5.Text & "','" & Me.TextBox6.Text & "','" & Me.TextBox7.Text & "','" & Me.TextBox8.Text & "')"
        Dim c25 As New OleDbCommand
        c25.Connection = conn
        c25.CommandText = StrSql5
        c25.ExecuteNonQuery()

        MessageBox.Show("Перевозчик добавлен!", Рик, MessageBoxButtons.OK, MessageBoxIcon.Information)

    End Sub

    Private Sub ComboBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox1.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            TextBox2.Focus()

        End If
    End Sub

    Private Sub TextBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox2.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            TextBox3.Focus()

        End If
    End Sub

    Private Sub TextBox3_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox3.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            TextBox10.Focus()

        End If
    End Sub

    Private Sub TextBox10_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox10.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            TextBox9.Focus()

        End If
    End Sub

    Private Sub TextBox9_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox9.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            TextBox5.Focus()

        End If
    End Sub

    Private Sub TextBox5_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox5.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            TextBox6.Focus()

        End If
    End Sub

    Private Sub TextBox6_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox6.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            TextBox4.Focus()

        End If
    End Sub

    Private Sub TextBox4_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox4.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            TextBox7.Focus()

        End If
    End Sub

    Private Sub TextBox7_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox7.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            TextBox8.Focus()


        End If
    End Sub

    Private Sub ListBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles ListBox1.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            ListBox2.Focus()

        End If
    End Sub

    Private Sub ListBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles ListBox2.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Button1.Focus()

        End If
    End Sub

    Private Sub TextBox8_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox8.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True

            ListBox1.Focus()

        End If
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged

    End Sub

    Private Sub ДобавитьПеревозчика_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        Очист()
    End Sub
End Class