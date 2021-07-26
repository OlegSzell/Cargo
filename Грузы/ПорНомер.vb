Option Explicit On
Imports System.ComponentModel
Imports System.IO
Public Class ПорНомер
    Dim Files() As String
    Dim gl As Integer = 0
    Private Naz As String = Nothing
    Public SledPorRejsPer As Integer
    Public SledPorRejsClient As Integer
    Private list1 As BindingList(Of Путь)
    Private bslist1 As BindingSource

    Sub New()

        ' Этот вызов является обязательным для конструктора.
        InitializeComponent()

        ' Добавить код инициализации после вызова InitializeComponent().

    End Sub
    Sub New(ByVal _Naz As String)
        Naz = _Naz
        ' Этот вызов является обязательным для конструктора.
        InitializeComponent()

        ' Добавить код инициализации после вызова InitializeComponent().

    End Sub
    Private Sub ПорНомер_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Отмена = 0
        list1 = New BindingList(Of Путь)
        bslist1 = New BindingSource
        bslist1.DataSource = list1
        ListBox1.DataSource = bslist1
        ListBox1.DisplayMember = "Название"

        Dim f = Directory.GetFiles("Z:\RICKMANS\", "*" & Naz & "*", IO.SearchOption.AllDirectories).ToList()
        If f IsNot Nothing Then
            For Each b In f
                Dim m As New Путь With {.Название = Path.GetFileName(b), .ПолныйПуть = b}
                list1.Add(m)
            Next

        End If

        Label1.Text = Naz

    End Sub
    Public Class Путь
        Public Property ПолныйПуть As String
        Public Property Название As String
    End Class

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If TextBox1.Enabled = True And TextBox1.Text = "" Then
            MessageBox.Show("Запишите порядковое число последнего рейса", Рик)
            Exit Sub
        ElseIf TextBox1.Enabled = True And TextBox1.Text <> "" Then
            If pro = 1 Then
                SledPorRejsClient = CType((TextBox1.Text), Integer) + 1
            Else
                SledPorRejsPer = CType((TextBox1.Text), Integer) + 1
            End If
        End If
        Dim mo As New AllUpd
        Using db As New dbAllDataContext(_cn3)
            If GroupBox2.Enabled = True Then
                Dim strsql As String = String.Empty
                If RichTextBox1.Text = "" Or RichTextBox2.Text = "" Then
                    MessageBox.Show("Заполните Должность и фамилию руководителя!", Рик)
                    Exit Sub
                End If

                If pro = 1 Then
                    Dim f = db.Клиент.Where(Function(x) x.НазваниеОрганизации = Naz).Select(Function(x) x).FirstOrDefault()
                    If f IsNot Nothing Then
                        With f
                            .Должность = RichTextBox1.Text
                            .ФИОРуководителя = RichTextBox2.Text
                        End With
                        db.SubmitChanges()
                        mo.КлиентAllAsync()
                    End If
                    'strsql = "UPDATE Клиент SET Должность='" & RichTextBox1.Text & "', ФИОРуководителя='" & RichTextBox2.Text & "' WHERE НазваниеОрганизации='" & Рейс.ComboBox3.Text & "'"
                    'Updates3(strsql)
                Else
                    Dim f1 = db.Перевозчики.Where(Function(x) x.Названиеорганизации = Naz).Select(Function(x) x).FirstOrDefault()
                    If f1 IsNot Nothing Then
                        With f1
                            .Должность = RichTextBox1.Text
                            .ФИОРуководителя = RichTextBox2.Text
                        End With
                        db.SubmitChanges()
                        mo.ПеревозчикиAllAsync()
                    End If

                    'strsql = "UPDATE Перевозчики SET Должность='" & RichTextBox1.Text & "', ФИОРуководителя='" & RichTextBox2.Text & "' WHERE Названиеорганизации='" & Рейс.ComboBox4.Text & "'"
                    'Updates3(strsql)

                End If


            End If

            If GroupBox1.Enabled = True Then
                If RichTextBox4.Text = "" Or MaskedTextBox1.MaskCompleted = False Then
                    MessageBox.Show("Заполните номер и дату договора!", Рик)
                    Exit Sub
                End If

                If pro = 1 Then
                    Dim f2 = db.Клиент.Where(Function(x) x.НазваниеОрганизации = Naz).Select(Function(x) x).FirstOrDefault()
                    If f2 IsNot Nothing Then
                        With f2
                            .Договор = RichTextBox4.Text
                            .Дата = MaskedTextBox1.Text
                        End With
                        db.SubmitChanges()
                        mo.КлиентAllAsync()
                    End If


                    'Dim strsql As String = "UPDATE Клиент SET Договор='" & RichTextBox4.Text & "', Дата='" & MaskedTextBox1.Text & "' WHERE НазваниеОрганизации='" & Рейс.ComboBox3.Text & "'"
                    'Updates3(strsql)
                Else
                    Dim f3 = db.Перевозчики.Where(Function(x) x.Названиеорганизации = Naz).Select(Function(x) x).FirstOrDefault()
                    If f3 IsNot Nothing Then
                        With f3
                            .Договор = RichTextBox4.Text
                            .Дата = MaskedTextBox1.Text
                        End With
                        db.SubmitChanges()
                        mo.ПеревозчикиAllAsync()
                    End If
                    'Dim strsql As String = "UPDATE Перевозчики SET Договор='" & RichTextBox4.Text & "', Дата='" & MaskedTextBox1.Text & "' WHERE Названиеорганизации='" & Рейс.ComboBox4.Text & "'"
                    'Updates3(strsql)
                End If
            End If
        End Using
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
        Process.Start(list1.ElementAt(ListBox1.SelectedIndex).ПолныйПуть)
        'Dim ff As ListBox.SelectedIndexCollection = ListBox1.SelectedIndices


        'If Not ListBox1.SelectedIndex = -1 Then

        '    For Each p As Integer In ff
        '        Process.Start(Files(p))
        '    Next

        'End If

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