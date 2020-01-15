Option Explicit On
Imports System.Data.OleDb
Public Class ВодитДан
    Private Sub ВодитДан_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        If Рейс.ComboBox4.Text = "" Then
            MessageBox.Show("Выберите перевозчика!", Рик)
            Me.Close()
            Exit Sub
        End If

        Dim strsql As String = "SELECT DISTINCT НомерАвтомобиля FROM РейсыПеревозчика WHERE НазвОрганизации='" & Рейс.ComboBox4.Text & "'"
        Dim strsql1 As String = "SELECT DISTINCT Водитель FROM РейсыПеревозчика WHERE НазвОрганизации='" & Рейс.ComboBox4.Text & "' ORDER BY Водитель"
        Dim ds As DataTable = Selects(strsql)
        Dim ds1 As DataTable = Selects(strsql1)
        Label2.Text = Рейс.ComboBox4.Text

        Me.ComboBox1.AutoCompleteCustomSource.Clear()
        Me.ComboBox1.Items.Clear()
        For Each r As DataRow In ds1.Rows
            Me.ComboBox1.AutoCompleteCustomSource.Add(r.Item(0).ToString())
            Me.ComboBox1.Items.Add(r(0).ToString)
        Next


        Me.ComboBox2.AutoCompleteCustomSource.Clear()
        Me.ComboBox2.Items.Clear()
        For Each r As DataRow In ds.Rows
            Me.ComboBox2.AutoCompleteCustomSource.Add(r.Item(0).ToString())
            Me.ComboBox2.Items.Add(r(0).ToString)
        Next






    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        If ComboBox1.Text <> "" Then
            Рейс.RichTextBox9.Text = ComboBox1.Text
        End If

        If ComboBox2.Text <> "" Then
            Рейс.RichTextBox8.Text = ComboBox2.Text
        End If
        ComboBox2.Text = ""
        ComboBox1.Text = ""
        Me.Close()
    End Sub
End Class