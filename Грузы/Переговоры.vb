Option Explicit On
Imports System.ComponentModel
Imports System.Data.OleDb
Public Class ПереговорыКлиент
    Dim код As Integer
    Private Sub ПереговорыКлиент_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        комб1()
    End Sub
    Friend Sub комб1()
        MaskedTextBox1.Text = Now.ToShortDateString
        Dim strsql As String = "SELECT DISTINCT Клиент FROM ПереговорыКлиент"
        Dim ds As DataTable = Selects3(strsql)

        Me.ComboBox1.AutoCompleteCustomSource.Clear()
        Me.ComboBox1.Items.Clear()
        For Each r As DataRow In ds.Rows
            Me.ComboBox1.AutoCompleteCustomSource.Add(r.Item(0).ToString())
            Me.ComboBox1.Items.Add(r(0).ToString)
        Next
    End Sub
    Private Sub чист()
        RichTextBox3.Text = ""
        MaskedTextBox1.Text = ""
        RichTextBox1.Text = ""
        MaskedTextBox2.Text = ""
        RichTextBox2.Text = ""
        ComboBox2.Text = ""
        ComboBox2.Items.Clear()
        ComboBox2.Enabled = False
        Label7.Enabled = False
    End Sub
    Public Sub com1()
        чист()
        Dim strsql As String = "SELECT Клиент FROM ПереговорыКлиент WHERE Клиент='" & ComboBox1.Text & "'"
        Dim ds As DataTable = Selects3(strsql)

        Dim strsql1 As String = "SELECT КонтДанные FROM ПереговорыКлиент WHERE Клиент='" & ComboBox1.Text & "'"
        Dim ds1 As DataTable = Selects3(strsql1)

        ComboBox3.Items.Clear()
        For Each r As DataRow In ds1.Rows
            Me.ComboBox3.Items.Add(r(0).ToString)
        Next

        If ds.Rows.Count > 1 Then
            ComboBox2.Enabled = True
            Label7.Enabled = True
            вставка()
        Else
            Первыйраз()
        End If
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        com1()
    End Sub
    Private Sub Первыйраз()
        Dim strsql As String = "SELECT * FROM ПереговорыКлиент WHERE Клиент='" & ComboBox1.Text & "' AND Экспедитор='" & Экспедитор & "'"
        Dim ds As DataTable = Selects3(strsql)
        If errds = 1 Then Exit Sub
        RichTextBox3.Text = ds.Rows(0).Item(6).ToString
        MaskedTextBox1.Text = ComboBox2.Text
        RichTextBox1.Text = ds.Rows(0).Item(3).ToString
        MaskedTextBox2.Text = ds.Rows(0).Item(4).ToString
        RichTextBox2.Text = ds.Rows(0).Item(5).ToString
    End Sub
    Private Sub вставка()
        Dim strsql As String = "SELECT ДатаПереговоров FROM ПереговорыКлиент WHERE Клиент='" & ComboBox1.Text & "' AND Экспедитор='" & Экспедитор & "'"
        Dim ds As DataTable = Selects3(strsql)

        Me.ComboBox2.Items.Clear()
        For Each r As DataRow In ds.Rows
            Me.ComboBox2.Items.Add(Strings.Left(r(0).ToString, 10))
        Next
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        Dim strsql As String = "SELECT * FROM ПереговорыКлиент WHERE Клиент='" & ComboBox1.Text & "' and ДатаПереговоров = #" & Format(CDate(ComboBox2.Text), "MM\/dd\/yyyy") & "# AND Экспедитор='" & Экспедитор & "'"
        Dim ds As DataTable = Selects3(strsql)
        код = Nothing
        код = ds.Rows(0).Item(0)
        RichTextBox3.Text = ds.Rows(0).Item(6).ToString
        MaskedTextBox1.Text = ComboBox2.Text
        RichTextBox1.Text = ds.Rows(0).Item(3).ToString
        MaskedTextBox2.Text = ds.Rows(0).Item(4).ToString
        RichTextBox2.Text = ds.Rows(0).Item(5).ToString
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If ComboBox1.Text = "" Then
            MessageBox.Show("Выберите организацию!", Рик)
            Exit Sub
        End If

        If MaskedTextBox1.Text = Nothing Then
            MessageBox.Show("Введите дату переговоров!", Рик)
            Exit Sub
        End If

        If MaskedTextBox2.Text = Nothing Then
            MessageBox.Show("Введите дату напоминания!", Рик)
            Exit Sub
        End If

        If ComboBox2.Text <> "" Then
            Обновл()
            MessageBox.Show("Данные обновлены!", Рик)
            чист()
        Else
            Новый()
            MessageBox.Show("Данные внесены!", Рик)
            чист()
        End If

    End Sub
    Private Sub Обновл()
        Dim strsql As String
        If MaskedTextBox2.MaskCompleted = False Then
            strsql = "UPDATE ПереговорыКлиент SET ДатаПереговоров='" & MaskedTextBox1.Text & "', ТекстПереговора='" & RichTextBox1.Text & "', 
КонтДанные='" & RichTextBox3.Text & "', Экспедитор='" & Экспедитор & "'
WHERE Код=" & код & ""
        Else
            strsql = "UPDATE ПереговорыКлиент SET ДатаПереговоров='" & MaskedTextBox1.Text & "', ТекстПереговора='" & RichTextBox1.Text & "',
ДатаНапоминания='" & MaskedTextBox2.Text & "',ТекстНапоминания='" & RichTextBox2.Text & "', КонтДанные='" & RichTextBox3.Text & "', Экспедитор='" & Экспедитор & "'
WHERE Код=" & код & ""
        End If


        Updates3(strsql)
    End Sub
    Private Sub Новый()
        Dim strsql As String
        If MaskedTextBox2.MaskCompleted = False Then
            strsql = "INSERT INTO ПереговорыКлиент(Клиент,ДатаПереговоров,ТекстПереговора,КонтДанные,Экспедитор) VALUES('" & ComboBox1.Text & "',
'" & MaskedTextBox1.Text & "','" & RichTextBox1.Text & "','" & RichTextBox3.Text & "','" & Экспедитор & "')"
        Else
            strsql = "INSERT INTO ПереговорыКлиент(Клиент,ДатаПереговоров,ТекстПереговора,ДатаНапоминания,ТекстНапоминания,КонтДанные,Экспедитор) VALUES('" & ComboBox1.Text & "',
'" & MaskedTextBox1.Text & "','" & RichTextBox1.Text & "','" & MaskedTextBox2.Text & "','" & RichTextBox2.Text & "','" & RichTextBox3.Text & "','" & Экспедитор & "')"
        End If
        Updates3(strsql)
    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged
        RichTextBox3.Text = ComboBox3.Text
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        ДобОргДляПерегов.ShowDialog()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim strsql As String = "SELECT "
        ДобОргДляПерегов.ShowDialog()
    End Sub

    Private Sub ПереговорыКлиент_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        чист()
    End Sub

    Private Sub GroupBox1_Enter(sender As Object, e As EventArgs) Handles GroupBox1.Enter

    End Sub
End Class