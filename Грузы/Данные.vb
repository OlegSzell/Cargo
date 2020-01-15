Option Explicit On
Imports System.Data.OleDb
Imports System.Threading
Public Class Данные
    Dim hg, idClient As Integer
    Public idClient2 As String
    Dim рик As String = "ООО Рикманс"
    Private Delegate Sub com1()
    Private Delegate Sub com2()
    Private Delegate Sub com3()
    Private Delegate Sub com4()
    Private Sub com11()
        Dim ds As DataTable
        Dim StrSql As String
        If ComboBox1.InvokeRequired Then
            Me.Invoke(New com1(AddressOf com11))
        Else
            StrSql = "SELECT DISTINCT Организация FROM ГрузыКлиентов ORDER BY Организация"
            'ds = Selects3(StrSql)
            ds = Selects3(StrSql)
            Me.ComboBox1.AutoCompleteCustomSource.Clear()
            Me.ComboBox1.Items.Clear()
            For Each r As DataRow In ds.Rows
                Me.ComboBox1.AutoCompleteCustomSource.Add(r.Item(0).ToString())
                Me.ComboBox1.Items.Add(r(0).ToString)
            Next
            'ComboBox1.Text = ""

        End If
    End Sub
    Private Sub com21()
        Dim StrSql1 As String
        Dim ds1 As DataTable
        If ComboBox2.InvokeRequired Then
            Me.Invoke(New com2(AddressOf com21))
        Else
            StrSql1 = "SELECT Страна FROM Страна ORDER BY Страна"
            'ds1 = Selects3(StrSql1)
            ds1 = Selects3(StrSql1)
            Me.ComboBox2.Items.Clear()
            Me.ComboBox2.AutoCompleteCustomSource.Clear()
            For Each r As DataRow In ds1.Rows
                Me.ComboBox2.AutoCompleteCustomSource.Add(r.Item(0).ToString())
                Me.ComboBox2.Items.Add(r(0).ToString)

            Next
            ComboBox2.Text = ""
        End If

    End Sub
    Private Sub com31()
        Dim StrSql3 As String
        Dim ds3 As DataTable
        If ComboBox3.InvokeRequired Then
            Me.Invoke(New com3(AddressOf com31))
        Else
            StrSql3 = "SELECT Страна FROM Страна ORDER BY Страна"
            'ds3 = Selects3(StrSql3)
            ds3 = Selects3(StrSql3)
            Me.ComboBox3.AutoCompleteCustomSource.Clear()
            Me.ComboBox3.Items.Clear()
            For Each r As DataRow In ds3.Rows
                Me.ComboBox3.AutoCompleteCustomSource.Add(r.Item(0).ToString())
                Me.ComboBox3.Items.Add(r(0).ToString)

            Next
            ComboBox3.Text = ""

        End If
    End Sub

    Private Sub com41()
        If ComboBox4.InvokeRequired Then
            Me.Invoke(New com4(AddressOf com41))
        Else
            ComboBox4.Enabled = False
        End If
    End Sub

    Private Sub Ref()


        Dim c1 As New Thread(AddressOf com11)
        c1.IsBackground = True
        c1.Start()

        Dim c2 As New Thread(AddressOf com21)
        c2.IsBackground = True
        c2.Start()

        Dim c3 As New Thread(AddressOf com31)
        c3.IsBackground = True
        c3.Start()

        Dim c4 As New Thread(AddressOf com41)
        c4.IsBackground = True
        c4.Start()

    End Sub
    Private Sub Данные_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        'conn = New OleDbConnection
        'conn.ConnectionString = ConString
        ''conn.ConnectionString = ConString
        'Try
        '    conn.Open()
        'Catch ex As Exception
        '    MessageBox.Show("Не подключен диск U")
        'End Try
        Dim Год As Integer = Year(Now)

        'If Me.Прием_Load = vbTrue Then Form1.Load = False

        TextBox8.Text = DateTime.Now.ToString("dd.MM.yyyy")

        RichTextBox1.Text = "Добрый день.
BY>B
10-12.04.19
Рабкор(Гом.обл) - Бельгия48
Светлогорск(Гом.обл) - Бельгия48
дрова на паллетах, до 23 т- тент - 1250е,
до 22т - реф - 1150е, растаможка - Вильнюс, 
оплата безнал, евро, предложите авто."
        Dim fg As New Thread(AddressOf Ref)
        fg.IsBackground = True
        fg.Start()
    End Sub
    Private Sub Обновление()
        Dim StrSql As String
        StrSql = "UPDATE ГрузыКлиентов SET Дата = '" & TextBox8.Text & "', Груз='" & TextBox2.Text & "', СтранаЗагрузки='" & ComboBox2.Text & "',
        СтранаВыгрузки = '" & ComboBox3.Text & "', ГородЗагрузки = '" & TextBox5.Text & "',
ГородВыгрузки = '" & TextBox6.Text & "', Ставка = '" & TextBox7.Text & "', СтавкаПеревозу='" & TextBox1.Text & "', ДляСкайпа = '" & RichTextBox2.Text & "',
        ОрганизКонтакт = '" & RichTextBox3.Text & "'
        WHERE Код = " & idClient & ""
        'Updates(stroka:=StrSql)
        Updates3(stroka:=StrSql)
        MessageBox.Show("Данные обновлены!")
        refreshes()
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If hg = 0 Then
            Save()
            Ref()
        Else
            Select Case MessageBox.Show("Нажмите Да (если надо изменить данные)" & vbCrLf & "Нажмите Нет (если надо добавить новый груз", рик, MessageBoxButtons.YesNoCancel, MessageBoxIcon.Information)
                Case DialogResult.Yes
                    Обновление()
                    Ref()
                Case DialogResult.No
                    Save()
                    Ref()
                Case DialogResult.Cancel
                    Очистка()
                    Exit Sub
            End Select
        End If

    End Sub
    Private Sub Save()
        '        Dim StrSql As String = "SELECT Организация, Дата, Груз, СтранаЗагрузки,
        'СтранаВыгрузки, ГородЗагрузки, ГородВыгрузки, Ставка, регионЗагрузки, Экспедитор, СтавкаПеревозу, Состояние, ДляСкайпа, ОрганизКонтакт
        'FROM ГрузыКлиентов"
        'Dim c As New OleDbCommand
        'c.Connection = conn
        'c.CommandText = StrSql
        'Dim ds As New DataSet
        'Dim da As New OleDbDataAdapter(c)
        'da.Fill(ds, "Сохранение")
        'Dim cb As New OleDbCommandBuilder(da)
        'Dim dsNewRow As DataRow
        'dsNewRow = ds.Tables("Сохранение").NewRow()
        'dsNewRow.Item("Организация") = ComboBox1.Text
        'dsNewRow.Item("Дата") = TextBox8.Text
        'dsNewRow.Item("Груз") = TextBox2.Text
        'dsNewRow.Item("СтранаЗагрузки") = ComboBox2.Text
        'dsNewRow.Item("СтранаВыгрузки") = ComboBox3.Text
        'dsNewRow.Item("ГородЗагрузки") = Me.TextBox5.Text
        'dsNewRow.Item("ГородВыгрузки") = Me.TextBox6.Text
        'dsNewRow.Item("Ставка") = Me.TextBox7.Text
        'dsNewRow.Item("регионЗагрузки") = Me.ComboBox4.Text
        'dsNewRow.Item("Экспедитор") = Экспедитор
        'dsNewRow.Item("СтавкаПеревозу") = Me.TextBox1.Text
        'dsNewRow.Item("Состояние") = "Груз в работе"
        'dsNewRow.Item("ДляСкайпа") = RichTextBox2.Text
        'dsNewRow.Item("ОрганизКонтакт") = RichTextBox3.Text
        'ds.Tables("Сохранение").Rows.Add(dsNewRow)
        'da.Update(ds, "Сохранение")


        Updates3(stroka:="INSERT INTO ГрузыКлиентов(Организация,Дата,Груз,СтранаЗагрузки,СтранаВыгрузки,ГородЗагрузки,ГородВыгрузки,Ставка,регионЗагрузки,
Экспедитор,СтавкаПеревозу,Состояние,ДляСкайпа,ОрганизКонтакт) VALUES('" & ComboBox1.Text & "','" & TextBox8.Text & "','" & TextBox2.Text & "','" & ComboBox2.Text & "',
'" & ComboBox3.Text & "','" & TextBox5.Text & "','" & TextBox6.Text & "','" & TextBox7.Text & "','" & ComboBox4.Text & "','" & Экспедитор & "',
'" & TextBox1.Text & "','Груз в работе','" & RichTextBox2.Text & "','" & RichTextBox3.Text & "')")


        MessageBox.Show("Сохранено!")


        ComboBox2.Text = ""
        ComboBox3.Text = ""
        TextBox7.Text = ""
        TextBox8.Text = ""
        TextBox5.Text = ""
        TextBox6.Text = ""
        TextBox2.Text = ""
        ComboBox4.Text = ""
        TextBox1.Text = ""
        RichTextBox2.Text = ""
        RichTextBox3.Text = ""
        refreshes()





    End Sub
    Private Sub refreshes()
        'Dim ds As DataTable = Selects3(StrSql:="SELECT ГрузыКлиентов.Дата, ГрузыКлиентов.Груз
        'From ГрузыКлиентов WHERE ГрузыКлиентов.Организация='" & ComboBox1.Text & "' AND Экспедитор='" & Экспедитор & "'ORDER BY ГрузыКлиентов.Дата DESC")
        Dim ds As DataTable = Selects3(StrSql:="SELECT ГрузыКлиентов.Дата, ГрузыКлиентов.Груз
FROM ГрузыКлиентов WHERE ГрузыКлиентов.Организация='" & ComboBox1.Text & "' AND Экспедитор='" & Экспедитор & "'ORDER BY ГрузыКлиентов.Дата DESC")

        ListBox1.Items.Clear()
        ListBox2.Items.Clear()
        For Each r As DataRow In ds.Rows
            Me.ListBox1.Items.Add(Strings.Left(r(0), 10))
            Me.ListBox2.Items.Add(r(1).ToString)
        Next

        TextBox1.Text = ""
        ComboBox2.Text = ""
        ComboBox3.Text = ""
        TextBox7.Text = ""
        TextBox8.Text = Date.Now.ToShortDateString
        TextBox5.Text = ""
        TextBox6.Text = ""
        TextBox2.Text = ""
        ComboBox4.Text = ""
        RichTextBox2.Text = ""
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        If Not Экспедитор <> "" Then
            MessageBox.Show("Выберите экспедитора!")
            Exit Sub
        End If

        Dim strsql As String = "SELECT ГрузыКлиентов.Дата, ГрузыКлиентов.Груз, ГрузыКлиентов.ОрганизКонтакт
FROM ГрузыКлиентов WHERE ГрузыКлиентов.Организация='" & ComboBox1.Text & "' AND ГрузыКлиентов.Экспедитор='" & Экспедитор & "' ORDER BY ГрузыКлиентов.Дата desc"
        'Dim ds As DataTable = Selects3(strsql)
        Dim ds As DataTable = Selects3(strsql)

        If errds = 0 Then
            ListBox1.Items.Clear()
            ListBox2.Items.Clear()
            For Each r As DataRow In ds.Rows
                Me.ListBox1.Items.Add(Strings.Left(r(0), 10))
                Me.ListBox2.Items.Add(r(1).ToString)
            Next
            RichTextBox3.Text = ""
            RichTextBox3.Text = ds.Rows(0).Item(2).ToString
        End If

        TextBox2.Text = ""
        TextBox5.Text = ""
        TextBox6.Text = ""
        TextBox7.Text = ""
        ComboBox2.Text = String.Empty
        ComboBox3.Text = String.Empty
        ComboBox4.Text = String.Empty
        hg = 0






    End Sub

    Private Sub ComboBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox1.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            TextBox2.Focus()

        End If
    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged

    End Sub

    Private Sub TextBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox2.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            ComboBox2.Focus()

        End If
    End Sub

    Private Sub ComboBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox2.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            ComboBox3.Focus()

        End If
    End Sub

    Private Sub ComboBox3_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox3.KeyDown
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
            TextBox7.Focus()

        End If
    End Sub

    Private Sub TextBox7_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox7.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Button1.Focus()

        End If
    End Sub

    Private Sub ListBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox1.SelectedIndexChanged
        ListBox2.SelectedIndex = ListBox1.SelectedIndex
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Очистка()
    End Sub
    Private Sub Очистка()
        TextBox8.Text = Date.Now.ToShortDateString
        ComboBox1.Text = String.Empty
        ComboBox3.Text = String.Empty
        ComboBox2.Text = String.Empty
        ComboBox4.Text = String.Empty
        TextBox2.Text = ""
        TextBox5.Text = ""
        TextBox6.Text = ""
        TextBox7.Text = ""
        RichTextBox2.Text = ""
        ListBox1.Items.Clear()
        ListBox2.Items.Clear()
        hg = 0
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        If ComboBox2.Text = "Россия" Then

            '            Dim ds2 As DataTable = Selects3(StrSql:="SELECT DISTINCT РегионыРоссии.Регионы
            'From Страна INNER Join РегионыРоссии On Страна.Код = РегионыРоссии.Страны
            'Where Страна.Страна = '" & ComboBox2.Text & "'")

            Dim ds2 As DataTable = Selects3(StrSql:="SELECT DISTINCT РегионыРоссии.Регионы
From Страна INNER Join РегионыРоссии On Страна.Код = РегионыРоссии.Страны
Where Страна.Страна = '" & ComboBox2.Text & "'")


            Me.ComboBox4.Items.Clear()
            Me.ComboBox4.AutoCompleteCustomSource.Clear()
            For Each r As DataRow In ds2.Rows
                Me.ComboBox4.AutoCompleteCustomSource.Add(r.Item(0).ToString())
                Me.ComboBox4.Items.Add(r(0).ToString)
            Next

            ComboBox4.Enabled = True
            Label9.Enabled = True
        Else
            ComboBox4.Enabled = False
            Label9.Enabled = False
        End If



    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

        ПоискПеревозчиков.Show()
        ПоискПеревозчиков.ComboBox1.Text = Me.ComboBox2.Text
        If Me.ComboBox2.Text = "Россия" Then
            ПоискПеревозчиков.ComboBox2.Text = Me.ComboBox4.Text
        End If
        Me.Close()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        If MessageBox.Show("Удалить выбранный рейс?", рик, MessageBoxButtons.OKCancel, MessageBoxIcon.Question) = DialogResult.Cancel Then
            Exit Sub
        End If

        'Updates(stroka:="DELETE * FROM ГрузыКлиентов WHERE Код=" & idClient & "")
        Updates3(stroka:="DELETE FROM ГрузыКлиентов WHERE Код=" & idClient & "")

        MessageBox.Show("Данные удалены!", рик)
        refreshes()

    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        КлиентДанные.ShowDialog()
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        idClient2 = ComboBox1.Text
        ВыборкаРабочегоВариантаДляДанных.ShowDialog()

    End Sub

    'Private Sub ComboBox5_SelectedIndexChanged(sender As Object, e As EventArgs)

    '    Пароль.ShowDialog()

    '    If Парол1 = "oleg110403" And ComboBox5.Text = "Олег" Then
    '        MessageBox.Show("Пароль принят!")
    '        ComboBox5.Text = "Олег"
    '        ComboBox5.Enabled = False
    '        Экспедитор = "Олег"
    '        Exit Sub
    '    ElseIf Not Парол1 = "oleg110403" And ComboBox5.Text = "Олег" Then
    '        MessageBox.Show("Пароль НЕ принят!")
    '        Exit Sub
    '    End If

    '    If Парол2 = "12345678" And ComboBox5.Text = "Катя" Then
    '        MessageBox.Show("Пароль принят!")
    '        ComboBox5.Text = "Катя"
    '        ComboBox5.Enabled = False
    '        Экспедитор = "Катя"
    '    Else
    '        MessageBox.Show("Пароль НЕ принят!")

    '    End If




    'End Sub

    Private Sub ListBox1_DoubleClick(sender As Object, e As EventArgs) Handles ListBox1.DoubleClick
        If ListBox1.SelectedIndex = -1 And ListBox2.SelectedIndex = -1 Then
            hg = 0
            Exit Sub
        End If

        'Dim df As String = Format(DateTimePicker1.Value, "MM\/dd\/yyyy")
        Dim dt, dt1, dt2, dt3 As String
        dt = Strings.Left(ListBox1.SelectedItem, 5)
        dt1 = Strings.Left(dt, 2)
        dt2 = Strings.Right(dt, 2)
        dt3 = Strings.Right(ListBox1.SelectedItem, 4)
        dt = dt2 & "/" & dt1 & "/" & dt3




        'Dim ds As DataTable = Selects3(StrSql:="SELECT * FROM ГрузыКлиентов WHERE Организация='" & ComboBox1.Text & "' AND Дата= #" & dt & "#  AND Груз='" & ListBox2.SelectedItem & "'")
        Dim ds As DataTable = Selects3(StrSql:="SELECT * FROM ГрузыКлиентов WHERE Организация='" & ComboBox1.Text & "' AND Дата= #" & dt & "#  AND Груз='" & ListBox2.SelectedItem & "'")

        idClient = Nothing
        Try
            idClient = ds.Rows(0).Item(0)
            TextBox2.Text = ds.Rows(0).Item(3).ToString
            ComboBox2.Text = ds.Rows(0).Item(4).ToString
            ComboBox4.Text = ds.Rows(0).Item(9).ToString
            ComboBox3.Text = ds.Rows(0).Item(5).ToString
            TextBox5.Text = ds.Rows(0).Item(6).ToString
            TextBox6.Text = ds.Rows(0).Item(7).ToString
            TextBox7.Text = ds.Rows(0).Item(8).ToString
            TextBox8.Text = Strings.Left(ds.Rows(0).Item(2).ToString, 10)
            TextBox1.Text = ds.Rows(0).Item(11).ToString
            RichTextBox2.Text = ds.Rows(0).Item(13).ToString
            RichTextBox3.Text = ds.Rows(0).Item(14).ToString
            hg = 1
        Catch ex As Exception
            MessageBox.Show("Выберите организацию!", рик)
        End Try





    End Sub
End Class
