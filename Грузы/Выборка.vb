Option Explicit On
Imports System.Data.OleDb
Imports System.Data.SqlClient
Imports System.Threading
Public Class Выборка

    'Public Da As New OleDbDataAdapter 'Адаптер
    'Public Ds As New DataSet 'Пустой набор записей
    'Dim tbl As New DataTable
    'Dim cb As OleDb.OleDbCommandBuilder
    'Dim Рик As String = "ООО Рикманс"
    Dim загр As String
    Private Delegate Sub com1()
    Private Delegate Sub com2()
    Public СтрПоиска As String



    Private Sub DateTimePicker1_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker1.ValueChanged
        Refreshgrid()

    End Sub
    Public Sub Refreshgrid()

        Dim Организ As String = ComboBox1.Text
        Dim StrSql As String = String.Empty

        Dim времянач As String = Format(DateTimePicker1.Value, "MM\/dd\/yyyy")
        Dim времякон As String = DateTime.Now.ToString("MM\/dd\/yyyy")
        Dim d As Date = Now
        Dim d2 As Date = Now
        d = d.AddDays(-10)
        d2 = d2.AddDays(10)

        'Dim времянач1 As String = Format(d2, "yyyy\/MM\/dd")
        'Dim времякон1 As String = Format(d, "yyyy\/MM\/dd")



        'Dim времянач1 As String = Replace(Format(d2, "yyyy\/MM\/dd"), "/", "")
        'Dim времякон1 As String = Replace(Format(d, "yyyy\/MM\/dd"), "/", "")




        If CheckBox3.Checked = True And TextBox2.Text = "" Then
            'StrSql = "SELECT * From ГрузыКлиентов WHERE Дата BETWEEN @d1 AND @d2 ORDER BY Организация"

            StrSql = "SELECT Организация, Дата, Груз, СтранаЗагрузки as [Страна загрузки], СтранаВыгрузки as [Страна выгрузки],
            ГородЗагрузки as [Город загрузки], ГородВыгрузки as [город выгрузки], Ставка, Код, СтавкаПеревозу as [Ставка перевозчику], Состояние
            From ГрузыКлиентов Where Экспедитор ='" & Экспедитор & "' AND Дата BETWEEN @d1 AND @d2 ORDER BY Организация"

            'Where Экспедитор ='" & Экспедитор & "' AND Дата BETWEEN " & времянач1 & " AND " & времякон1 & " ORDER BY Организация" 'BETWEEN #" & времянач1 & "# AND #" & времякон1 & "# ORDER BY Организация

        ElseIf CheckBox3.Checked = False And TextBox2.Text = "" And CheckBox4.Checked = False Then
            StrSql = "SELECT Организация, Дата, Груз, СтранаЗагрузки as [Страна загрузки], СтранаВыгрузки as [Страна выгрузки],
ГородЗагрузки as [Город загрузки], ГородВыгрузки as [город выгрузки], Ставка, Код, СтавкаПеревозу as [Ставка перевозчику], Состояние
From ГрузыКлиентов
Where Организация = '" & Организ & "' AND Экспедитор='" & Экспедитор & "' ORDER BY Организация"

        ElseIf CheckBox3.Checked = True And TextBox2.Text <> "" Then

            StrSql = "SELECT Организация, Дата, Груз, СтранаЗагрузки as [Страна загрузки], СтранаВыгрузки as [Страна выгрузки],
ГородЗагрузки as [Город загрузки], ГородВыгрузки as [город выгрузки], Ставка, Код, СтавкаПеревозу as [Ставка перевозчику], Состояние
From ГрузыКлиентов
Where Экспедитор='" & Экспедитор & "' AND Дата BETWEEN @d1 AND @d2 AND Груз Like '%" & TextBox2.Text & "%' ORDER BY Организация"


        ElseIf ComboBox1.Text <> "" And TextBox2.Text <> "" Then
            StrSql = "SELECT Организация, Дата, Груз, СтранаЗагрузки as [Страна загрузки], СтранаВыгрузки as [Страна выгрузки],
ГородЗагрузки as [Город загрузки], ГородВыгрузки as [город выгрузки], Ставка, Код, СтавкаПеревозу as [Ставка перевозчику], Состояние
From ГрузыКлиентов
Where Организация = '" & Организ & "' AND Экспедитор='" & Экспедитор & "' AND Груз Like '%" & TextBox2.Text & "%' ORDER BY Организация"

        ElseIf CheckBox4.Checked = True And TextBox2.Text = "" Then
            StrSql = "SELECT Организация, Дата, Груз, СтранаЗагрузки as [Страна загрузки], СтранаВыгрузки as [Страна выгрузки],
ГородЗагрузки as [Город загрузки], ГородВыгрузки as [город выгрузки], Ставка, Код, СтавкаПеревозу as [Ставка перевозчику], Состояние
From ГрузыКлиентов ORDER BY Организация"
        ElseIf CheckBox4.Checked = True And TextBox2.Text <> "" Then
            StrSql = "SELECT Организация, Дата, Груз, СтранаЗагрузки as [Страна загрузки], СтранаВыгрузки as [Страна выгрузки],
ГородЗагрузки as [Город загрузки], ГородВыгрузки as [город выгрузки], Ставка, Код, СтавкаПеревозу as [Ставка перевозчику], Состояние
From ГрузыКлиентов Where Груз Like '%" & TextBox2.Text & "%' ORDER BY Организация"
        ElseIf CheckBox4.Checked = False And CheckBox3.Checked = False And ComboBox1.Text = "" And TextBox2.Text <> "" Then
            StrSql = "SELECT Организация, Дата, Груз, СтранаЗагрузки as [Страна загрузки], СтранаВыгрузки as [Страна выгрузки],
ГородЗагрузки as [Город загрузки], ГородВыгрузки as [город выгрузки], Ставка, Код, СтавкаПеревозу as [Ставка перевозчику], Состояние
From ГрузыКлиентов Where Груз Like '%" & TextBox2.Text & "%' ORDER BY Организация"

        End If

        'Dim tbl As DataTable = Selects3(StrSql)
        'Dim tbl As DataTable = Selects3(StrSql)

        Dim list As New List(Of Date)() From {d, d2}
        Dim tbl As DataTable = Selects3(StrSql, list)

        'Dim conn3 As New SqlConnection(ConString)
        'If conn3.State = ConnectionState.Closed Then
        '    conn3.Open()
        'End If

        'Dim c As New SqlCommand(StrSql, conn3)

        'c.Parameters.Add("@d2", SqlDbType.Date).Value = d2 'DateTimePicker1.Value
        'c.Parameters.Add("@d1", SqlDbType.Date).Value = d 'DateTimePicker2.Value

        'Dim dst As New DataTable
        'Dim da As New SqlDataAdapter(c)
        'Try
        '    da.Fill(dst)
        '    Dim gf As Object = dst.Rows(0).Item(0)
        '    Dim m As String = c.ExecuteScalar().ToString()
        '    If conn3.State = ConnectionState.Open Then
        '        conn3.Close()
        '    End If

        'Catch ex As Exception
        '    errds = 1
        '    If conn3.State = ConnectionState.Open Then
        '        conn3.Close()
        '    End If

        'End Try

        'Dim tbl As DataTable = dst



        'If загр = "" Then

        '        StrSql = "SELECT Организация, Дата, Груз, СтранаЗагрузки, СтранаВыгрузки, ГородЗагрузки, ГородВыгрузки, Ставка, Код
        'From ГрузыКлиентов
        'Where Организация = '" & Организ & "' And Дата Between #" & времянач & "# And #" & времякон & "# AND Экспедитор='" & Экспедитор & "' ORDER BY Дата"
        '        Else

        '            StrSql = "SELECT Организация, Дата, Груз, СтранаЗагрузки, СтранаВыгрузки, ГородЗагрузки, ГородВыгрузки, Ставка, Код
        'From ГрузыКлиентов
        'Where Организация = '" & Организ & "' And Дата Between #" & времянач & "# And #" & времякон & "# AND СтранаЗагрузки='" & загр & "' AND Экспедитор='" & Экспедитор & "' ORDER BY Дата"
        '        End If




        Grid1.DataSource = tbl
        'Grid1.Columns(0).FillWeight = 70
        'Grid1.Columns(1).FillWeight = 50
        Grid1.Columns(2).FillWeight = 550
        Grid1.Columns(8).Visible = False
        Grid1.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        For ij As Integer = 0 To Grid1.Rows.Count - 1
            Try
                If Grid1.Rows(ij).Cells("Состояние").Value = "Закрыт" Then
                    Grid1.Rows(ij).DefaultCellStyle.BackColor = Color.YellowGreen
                End If
            Catch ex As Exception

            End Try
        Next

    End Sub
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
            ComboBox1.Text = ""

        End If
    End Sub

    Private Sub com21()
        Dim ds As DataTable
        Dim StrSql As String
        If ComboBox2.InvokeRequired Then
            Me.Invoke(New com2(AddressOf com21))
        Else
            StrSql = "SELECT DISTINCT СтранаЗагрузки FROM ГрузыКлиентов ORDER BY СтранаЗагрузки"
            'ds = Selects3(StrSql)
            ds = Selects3(StrSql)
            Me.ComboBox2.AutoCompleteCustomSource.Clear()
            Me.ComboBox2.Items.Clear()
            For Each r As DataRow In ds.Rows
                Me.ComboBox2.AutoCompleteCustomSource.Add(r.Item(0).ToString())
                Me.ComboBox2.Items.Add(r(0).ToString)
            Next
            ComboBox2.Text = ""
            ComboBox2.Enabled = False

        End If
    End Sub



    Private Sub Выборка_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        'conn = New OleDbConnection
        'conn.ConnectionString = ConString
        'Try
        '    conn.Open()
        'Catch ex As Exception
        '    MessageBox.Show("Не подключен диск U")
        'End Try
        Dim Год As Integer = Year(Now)

        'If Me.Прием_Load = vbTrue Then Form1.Load = False

        TextBox1.Text = DateTime.Now.ToString("dd.MM.yyyy")

        Dim fg As New Thread(AddressOf com11)
        fg.IsBackground = True
        fg.Start()
        Dim fg1 As New Thread(AddressOf com21)
        fg1.IsBackground = True
        fg1.Start()
        DateTimePicker2.Value = Now.Date.AddDays(-50)
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        If ComboBox1.Text <> "" Then
            ComboBox2.Enabled = True
        Else
            ComboBox2.Enabled = False
        End If
        CheckBox3.Checked = False
        Refreshgrid()

    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        загр = ComboBox2.Text
        Refreshgrid()
    End Sub

    Private Sub ComboBox2_StyleChanged(sender As Object, e As EventArgs) Handles ComboBox2.StyleChanged

    End Sub

    Private Sub ComboBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox2.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            загр = ""
            Refreshgrid()
        End If
    End Sub

    Private Sub Grid1_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles Grid1.CellDoubleClick
        If CheckBox1.Checked = False And CheckBox2.Checked = False Then
            MessageBox.Show("Выберите страну поиска (место загрузки/выгрузки!)", Рик)
            Exit Sub
        End If
        If CheckBox1.Checked = True Then
            СтрПоиска = Grid1.CurrentRow.Cells("Страна загрузки").Value
        Else
            СтрПоиска = Grid1.CurrentRow.Cells("Страна выгрузки").Value
        End If
        IDГруз = Nothing
        IDГруз = Grid1.CurrentRow.Cells("Код").Value
        РабочаяФорма.Show()
        Me.Close()

    End Sub

    Private Sub CheckBox3_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox3.CheckedChanged
        Refreshgrid()
    End Sub

    Private Sub TextBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox2.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Refreshgrid()


        End If
    End Sub

    Private Sub CheckBox4_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox4.CheckedChanged
        If CheckBox4.Checked = True Then
            CheckBox3.Checked = False
            ComboBox1.Text = ""
            TextBox2.Text = ""
            Refreshgrid()
        Else
            Refreshgrid()
        End If

    End Sub

    Private Sub Grid1_CellMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles Grid1.CellMouseClick
        If e.Button = MouseButtons.Right Then
            IDГруз = Nothing
            NameПеревоза = ""

            IDГруз = Grid1.CurrentRow.Cells("Код").Value

            РабочаяФормаСостояние.TextBox1.Text = Grid1.CurrentRow.Cells("Организация").Value
            РабочаяФормаСостояние.TextBox2.Text = Grid1.CurrentRow.Cells("Груз").Value
            'Dim ds As DataTable = Selects3(StrSql:="SELECT ОрганизКонтакт FROM ГрузыКлиентов WHERE Организация='" & Grid1.CurrentRow.Cells("Организация").Value & "'")
            Dim ds As DataTable = Selects3(StrSql:="SELECT ОрганизКонтакт FROM ГрузыКлиентов WHERE Организация='" & Grid1.CurrentRow.Cells("Организация").Value & "'")
            Dim b As String = String.Empty
            For Each d As DataRow In ds.Rows
                If Not IsDBNull(d.Item(0)) Or d.Item(0).ToString <> "" Then
                    b = d.Item(0).ToString
                    Exit For
                End If
            Next

            If b <> "" Then
                РабочаяФормаСостояние.RichTextBox1.Text = b
            Else
                РабочаяФормаСостояние.RichTextBox1.Text = "Нет конт телефона"
            End If

            РабочаяФормаСостояние.ShowDialog()

        End If
    End Sub

    Private Sub TextBox3_KeyDown(sender As Object, e As KeyEventArgs)
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Refreshgrid()


        End If
    End Sub
End Class