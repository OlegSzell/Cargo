Imports System.ComponentModel

Public Class ДобОргДляПерегов
    Dim id As Integer
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        If ComboBox1.Text <> "" And RichTextBox1.Text = "" Then
            Dim strsql1 As String = "UPDATE ПереговорыКлиент SET КонтДанные='" & RichTextBox2.Text & "', Клиент='" & ComboBox1.Text & "' WHERE Код=" & id & ""
            Updates3(strsql1)
            MessageBox.Show("Данные обновлены!", Рик)

        ElseIf ComboBox1.Text <> "" And RichTextBox1.Text <> "" Then

            If MessageBox.Show("Выберите да - если хотите обновить действующего клиента" & vbCrLf & "выберите нет - если хотите создать нового!", Рик, MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.OK Then
                Dim strsql12 As String = "UPDATE ПереговорыКлиент SET КонтДанные='" & RichTextBox2.Text & "', Клиент='" & ComboBox1.Text & "' WHERE Код=" & id & ""
                Updates3(strsql12)
                MessageBox.Show("Данные обновлены!", Рик)

            Else
                Dim strsql2 As String = "SELECT ДатаПереговоров FROM ПереговорыКлиент WHERE Клиент='" & RichTextBox1.Text & "'"
                Dim ds2 As DataTable = Selects3(strsql2)
                If errds = 1 Then
                    новая()
                    MessageBox.Show("Организация добавлена!", Рик)
                    RichTextBox1.Text = ""
                    RichTextBox2.Text = ""

                End If
            End If
        Else
            Dim strsql As String = "SELECT ДатаПереговоров FROM ПереговорыКлиент WHERE Клиент='" & RichTextBox1.Text & "'"
            Dim ds As DataTable = Selects3(strsql)
            If errds = 1 Then
                новая()
                MessageBox.Show("Организация добавлена!", Рик)
                RichTextBox1.Text = ""
                RichTextBox2.Text = ""

            End If
        End If

        Me.Close()
        ПереговорыКлиент.комб1()
        RichTextBox1.Text = ""
        RichTextBox2.Text = ""
        ComboBox1.Items.Clear()


    End Sub
    Private Sub новая()
        Dim strsql As String = "INSERT INTO ПереговорыКлиент(Клиент,КонтДанные) VALUES('" & RichTextBox1.Text & "','" & RichTextBox2.Text & "')"
        Updates3(strsql)
    End Sub

    Private Sub ДобОргДляПерегов_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim strsql As String = "SELECT DISTINCT Клиент,Код FROM ПереговорыКлиент ORDER BY Клиент"
        Dim ds As DataTable = Selects3(strsql)

        ComboBox1.Items.Clear()
        ComboBox1.AutoCompleteCustomSource.Clear()
        For Each r As DataRow In ds.Rows
            ComboBox1.AutoCompleteCustomSource.Add(r.Item(0).ToString())
            ComboBox1.Items.Add(r(0).ToString)
            ComboBox2.Items.Add(r(1).ToString)
        Next


    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        id = Nothing
        id = ComboBox2.Items.Item(ComboBox1.SelectedIndex)
        RichTextBox2.Text = ""
        Dim strsql As String = "SELECT * FROM ПереговорыКлиент WHERE Код=" & id & ""
        Dim ds As DataTable = Selects3(strsql)
        RichTextBox2.Text = ds.Rows(0).Item(6).ToString

    End Sub

    Private Sub ДобОргДляПерегов_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        ComboBox1.Items.Clear()
        ComboBox2.Items.Clear()
        ComboBox1.Text = ""
        RichTextBox1.Text = ""
        RichTextBox2.Text = ""
    End Sub
End Class