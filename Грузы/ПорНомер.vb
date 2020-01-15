Option Explicit On
Imports System.ComponentModel
Imports System.Data.OleDb
Public Class ПорНомер
    Dim Files() As String
    Dim gl As Integer = 0
    Private Sub ПорНомер_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Отмена = 0
        Files = IO.Directory.GetFiles("Z:\RICKMANS\", "*" & Nm & "*", IO.SearchOption.AllDirectories)
        Dim gth4 As String
        Dim FilesList() As String = IO.Directory.GetFiles("Z:\RICKMANS\", "*" & Nm & "*", IO.SearchOption.AllDirectories)
        For n As Integer = 0 To FilesList.Length - 1
            gth4 = ""
            gth4 = IO.Path.GetFileName(FilesList(n))
            FilesList(n) = gth4
            'TextBox44.Text &= gth + vbCrLf
        Next

        ListBox1.Items.Clear()

        For i = 0 To FilesList.Length - 1 ' Распечатываем весь получившийся массив
            ListBox1.Items.Add(FilesList(i)) ' На ListBox2
        Next

        Label1.Text = Nm

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If TextBox1.Enabled = True And TextBox1.Text = "" Then
            MessageBox.Show("Запишите порядковое число последнего рейса", Рик)
            Exit Sub
        ElseIf TextBox1.Enabled = True And TextBox1.Text <> "" Then
            If pro = 1 Then
                Рейс.СлПорРейсКл = CType((TextBox1.Text), Integer) + 1
            Else
                Рейс.СлПорРейсПер = CType((TextBox1.Text), Integer) + 1
            End If
        End If

        If GroupBox2.Enabled = True Then
            Dim strsql As String
            If RichTextBox1.Text = "" Or RichTextBox2.Text = "" Then
                MessageBox.Show("Заполните Должность и фамилию руководителя!", Рик)
                Exit Sub
            End If

            If pro = 1 Then
                strsql = "UPDATE Клиент SET Должность='" & RichTextBox1.Text & "', ФИОРуководителя='" & RichTextBox2.Text & "' WHERE НазваниеОрганизации='" & Рейс.ComboBox3.Text & "'"
                Updates3(strsql)
            Else
                strsql = "UPDATE Перевозчики SET Должность='" & RichTextBox1.Text & "', ФИОРуководителя='" & RichTextBox2.Text & "' WHERE Названиеорганизации='" & Рейс.ComboBox4.Text & "'"
                Updates3(strsql)

            End If
        End If

        If GroupBox1.Enabled = True Then
            If RichTextBox4.Text = "" Or MaskedTextBox1.MaskCompleted = False Then
                MessageBox.Show("Заполните номер и дату договора!", Рик)
                Exit Sub
            End If

            If pro = 1 Then
                Dim strsql As String = "UPDATE Клиент SET Договор='" & RichTextBox4.Text & "', Дата='" & MaskedTextBox1.Text & "' WHERE НазваниеОрганизации='" & Рейс.ComboBox3.Text & "'"
                Updates3(strsql)
            Else
                Dim strsql As String = "UPDATE Перевозчики SET Договор='" & RichTextBox4.Text & "', Дата='" & MaskedTextBox1.Text & "' WHERE Названиеорганизации='" & Рейс.ComboBox4.Text & "'"
                Updates3(strsql)
            End If
        End If

        ListBox1.Enabled = False
        GroupBox1.Enabled = False
        GroupBox2.Enabled = False
        TextBox1.Text = ""
        TextBox1.Enabled = False
        RichTextBox1.Text = ""
        RichTextBox2.Text = ""
        RichTextBox4.Text = ""
        MaskedTextBox1.Text = ""
        gl = 1
        Me.Close()

    End Sub

    Private Sub ListBox1_DoubleClick(sender As Object, e As EventArgs) Handles ListBox1.DoubleClick

        If ListBox1.SelectedIndex = -1 Then
            MessageBox.Show("Выберите документ для просмотра!", Рик, MessageBoxButtons.OK)
            Exit Sub
        End If

        Dim ff As ListBox.SelectedIndexCollection = ListBox1.SelectedIndices


        If Not ListBox1.SelectedIndex = -1 Then

            For Each p As Integer In ff
                Process.Start(Files(p))
            Next

        End If

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        RichTextBox1.Text = ""
        RichTextBox2.Text = ""
        RichTextBox4.Text = ""
        MaskedTextBox1.Text = ""
        TextBox1.Text = ""
        Me.Close()
    End Sub

    Private Sub RichTextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles RichTextBox1.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.RichTextBox2.Focus()
        End If
    End Sub

    Private Sub RichTextBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles RichTextBox2.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.RichTextBox4.Focus()
        End If
    End Sub

    Private Sub RichTextBox4_KeyDown(sender As Object, e As KeyEventArgs) Handles RichTextBox4.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.MaskedTextBox1.Focus()
        End If
    End Sub

    Private Sub MaskedTextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles MaskedTextBox1.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.TextBox1.Focus()
        End If
    End Sub

    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.Button1.Focus()
        End If
    End Sub

    Private Sub ПорНомер_Closed(sender As Object, e As EventArgs) Handles Me.Closed

        If gl = 0 Then
            Отмена = 1
            Me.Close()
        End If

    End Sub


End Class