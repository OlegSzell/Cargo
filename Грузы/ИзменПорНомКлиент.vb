Option Explicit On
Imports System.Data.OleDb
Public Class ИзменПорНомКлиент
    Private Sub ИзменПорНомКлиент_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TextBox1.Text = ""
        If ОтКогоИзмен = 1 Then
            Dim strsql As String = "SELECT КоличРейсов FROM РейсыКлиента WHERE НомерРейса=" & Рейс.НомРес & ""
            Dim ds As DataTable = Selects3(strsql)
            TextBox1.Text = ds.Rows(0).Item(0).ToString
            Label3.Text = Рейс.ComboBox3.Text
        Else
            Dim strsql As String = "SELECT КоличРейсов FROM РейсыПеревозчика WHERE НомерРейса=" & Рейс.НомРес & ""
            Dim ds As DataTable = Selects3(strsql)
            TextBox1.Text = ds.Rows(0).Item(0).ToString
            Label3.Text = Рейс.ComboBox4.Text
        End If

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If TextBox2.Text = "" Then

            MessageBox.Show("Введите новый порядковый номер рейса!", Рик)
            Exit Sub
        End If
        If ОтКогоИзмен = 1 Then
            Dim strsql As String = "UPDATE РейсыКлиента SET КоличРейсов=" & CType(TextBox2.Text, Integer) & " WHERE НомерРейса=" & Рейс.НомРес & ""
            Updates3(strsql)
            If MessageBox.Show("Данные удачно изменены в базе!" & vbCrLf & "Внести изменения в файле эксель?", Рик, MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.No Then
                Me.Close()
                Exit Sub
            End If
            Me.Close()
            Рейс.ИзменВДействРейсе("Клиент")
            TextBox1.Text = ""
            TextBox2.Text = ""
        Else
            Dim strsql As String = "UPDATE РейсыПеревозчика SET КоличРейсов=" & CType(TextBox2.Text, Integer) & " WHERE НомерРейса=" & Рейс.НомРес & ""
            Updates3(strsql)
            If MessageBox.Show("Данные удачно изменены в базе!" & vbCrLf & "Внести изменения в файле эксель?", Рик, MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.No Then
                Me.Close()
                Exit Sub
            End If
            Me.Close()
            Рейс.ИзменВДействРейсе("Перевоз")
            TextBox1.Text = ""
            TextBox2.Text = ""

        End If


    End Sub
End Class