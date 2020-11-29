Option Explicit On
Imports System.ComponentModel
Imports System.Data.OleDb
Imports System.Threading


Public Class ДобавитьПеревозчика
    Public Da As New OleDbDataAdapter 'Адаптер
    Public Ds As New DataSet 'Пустой набор записей
    Private Delegate Sub com2()
    Private Delegate Sub com3()
    Private Delegate Sub com4()
    Private Delegate Sub comb22()
    'Dim tbl As New DataTable
    'Dim cb As OleDb.OleDbCommandBuilder

    Private Sub com11()
        Dim ds As DataTable
        Dim StrSql As String
        If ComboBox1.InvokeRequired Then
            Me.Invoke(New com2(AddressOf com11))
        Else
            StrSql = "SELECT DISTINCT [Наименование фирмы] FROM ПеревозчикиБаза ORDER BY [Наименование фирмы]"
            ds = Selects3(StrSql)
            Me.ComboBox1.AutoCompleteCustomSource.Clear()
            Me.ComboBox1.Items.Clear()
            For Each r As DataRow In ds.Rows
                Me.ComboBox1.AutoCompleteCustomSource.Add(r.Item(0).ToString())
                Me.ComboBox1.Items.Add(r(0).ToString)
            Next
            ComboBox1.Focus()
        End If
    End Sub
    Private Sub comb2()
        'Dim ds As DataTable
        Dim StrSql As String
        If ComboBox2.InvokeRequired Then
            Me.Invoke(New comb22(AddressOf comb2))
        Else
            Dim ds As DataTable = Selects3(StrSql:="SELECT ПолноеНазвание,Сокращенное FROM ФормаСобств ORDER BY ПолноеНазвание")
            Me.ComboBox2.AutoCompleteCustomSource.Clear()
            Me.ComboBox2.Items.Clear()
            Me.ComboBox3.Items.Clear()
            For Each r As DataRow In ds.Rows
                Me.ComboBox2.AutoCompleteCustomSource.Add(r.Item(0).ToString())
                Me.ComboBox2.Items.Add(r(0).ToString)
                Me.ComboBox3.Items.Add(r(1).ToString)
            Next
            ComboBox2.Focus()
        End If
    End Sub

    Private Sub com12()
        Dim ds As DataTable
        Dim StrSql As String
        If ListBox1.InvokeRequired Then
            Me.Invoke(New com3(AddressOf com12))
        Else
            StrSql = "SELECT DISTINCT Страна FROM Страна ORDER BY Страна"
            ds = Selects3(StrSql)
            ListBox1.Items.Clear()
            For Each r As DataRow In ds.Rows
                ListBox1.Items.Add(r(0).ToString)
            Next
        End If
    End Sub
    Private Sub com13()
        Dim ds As DataTable
        Dim StrSql As String
        If ListBox2.InvokeRequired Then
            Me.Invoke(New com4(AddressOf com13))
        Else
            StrSql = "SELECT DISTINCT Регионы FROM РегионыРоссии ORDER BY Регионы"
            ds = Selects3(StrSql)
            Me.ListBox2.Items.Clear()
            For Each r As DataRow In ds.Rows
                Me.ListBox2.Items.Add(r(0).ToString)
            Next
        End If
    End Sub
    Private Sub ДобавитьПеревозчика_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Очист()

        Dim Год As Integer = Year(Now)

        Dim d As New Thread(AddressOf com11)
        Dim d1 As New Thread(AddressOf com12)
        Dim d2 As New Thread(AddressOf com13)
        Dim d3 As New Thread(AddressOf comb2)
        d.IsBackground = True
        d1.IsBackground = True
        d2.IsBackground = True
        d3.IsBackground = True
        d.Start()
        d1.Start()
        d2.Start()
        d3.Start()

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
    End Sub
    Private Sub insertbaza()

        If ComboBox1.Text = "" Then
            MessageBox.Show("Выберите перевозчика!", Рик, MessageBoxButtons.OK)
            Exit Sub
        End If
        If ComboBox2.Text = "" Then
            MessageBox.Show("Выберите форму собственности перевозчика!", Рик, MessageBoxButtons.OK)
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

        Using db As New dbAllDataContext()
            Dim f As New ПеревозчикиБаза
            With f
                .Наименование_фирмы = ComboBox1.Text
                .Контактное_лицо = TextBox2.Text
                .Телефоны = TextBox3.Text
                .Страны_перевозок = country
                .Регионы = country2
                .ADR = TextBox9.Text
                .Кол_во_авто = TextBox4.Text
                .Вид_авто = TextBox10.Text
                .Тоннаж = TextBox5.Text
                .Объем = TextBox6.Text
                .Ставка = TextBox7.Text
                .Примечание = TextBox8.Text
                .Форма_собственности = ComboBox3.Text

            End With
            db.ПеревозчикиБаза.InsertOnSubmit(f)
            db.SubmitChanges()
            Dim mo As New AllUpd
            mo.ПеревозчикиБазаAllAsync()
        End Using




        '        Dim StrSql5 As String = "INSERT INTO ПеревозчикиБаза([Наименование фирмы],[Контактное лицо],Телефоны,[Страны перевозок],Регионы,ADR,[Кол-во авто],
        '[Вид_авто],Тоннаж,Объем,Ставка,Примечание,[Форма собственности])
        'VALUES('" & Me.ComboBox1.Text & "','" & Me.TextBox2.Text & "','" & Me.TextBox3.Text & "','" & country & "','" & country2 & "','" & Me.TextBox9.Text & "',
        ''" & Me.TextBox4.Text & "', '" & Me.TextBox10.Text & "','" & Me.TextBox5.Text & "','" & Me.TextBox6.Text & "','" & Me.TextBox7.Text & "','" & Me.TextBox8.Text & "','" & ComboBox3.Text & "')"
        '        Updates3(StrSql5)

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
    Private Sub ДобавитьПеревозчика_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        If ОбнПер = 1 Then
            Очист()
            ПоискПеревозчиков.Запуск()
        Else
            Очист()
        End If

    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        ComboBox3.Text = ComboBox3.Items.Item(ComboBox2.SelectedIndex) 'подчиненные комбобокс

    End Sub
End Class