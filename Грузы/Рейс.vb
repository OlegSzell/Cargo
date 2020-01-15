Option Explicit On
Imports System.Data.OleDb


Public Class Рейс
    Dim strsql As String
    Dim ds As DataTable
    Public СлРейс As Integer
    Public НомРес As Integer
    Public СлПорРейсКл, СлПорРейсПер, ПерезагрЛист1 As Integer
    Public RS As Object
    Public Excel_obj As Object
    Public Excel_Doc As Object
    Public ПутьРейса As String
    Dim Пров As Integer
    Dim ПрИзмНазКл As Boolean = False, ПрИзмНазПер As Boolean = False
    Dim ДогПор, ДогПорЭксп, ПорЭксп As String

    Private Sub Рейс_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.MdiParent = MDIParent1
        conn = New OleDbConnection
        conn.ConnectionString = ConString
        Try
            conn.Open()
        Catch ex As Exception
            MessageBox.Show("Не подключен диск U")
        End Try

        strsql = "SELECT Названиеорганизации FROM ПЕревозчики ORDER BY Названиеорганизации"
        ds = Selects(strsql)

        Me.ComboBox4.AutoCompleteCustomSource.Clear()
        Me.ComboBox4.Items.Clear()
        For Each r As DataRow In ds.Rows
            Me.ComboBox4.AutoCompleteCustomSource.Add(r.Item(0).ToString())
            Me.ComboBox4.Items.Add(r(0).ToString)
        Next

        Dim strsql1 As String = "SELECT НазваниеОрганизации FROM Клиент ORDER BY НазваниеОрганизации"
        Dim ds1 As DataTable = Selects(strsql1)

        Me.ComboBox3.AutoCompleteCustomSource.Clear()
        Me.ComboBox3.Items.Clear()
        For Each r As DataRow In ds1.Rows
            Me.ComboBox3.AutoCompleteCustomSource.Add(r.Item(0).ToString())
            Me.ComboBox3.Items.Add(r(0).ToString)
        Next
        ComboBox1.Items.Clear()
        ComboBox5.Items.Clear()
        ComboBox6.Items.Clear()
        ComboBox7.Items.Clear()
        ComboBox7.Enabled = False
        ComboBox8.Items.Clear()
        ComboBox8.Enabled = False
        Dim d() As String = {Now.Year, Now.Year - 1, Now.Year - 2, Now.Year - 3, Now.Year - 4}
        ComboBox1.Items.AddRange(d)
        ComboBox1.Text = Now.Year

        Dim d1() As String = {"Рубль", "Евро", "Доллар", "Росс.рубль"}
        ComboBox5.Items.AddRange(d1)
        ComboBox6.Items.AddRange(d1)

        Dim d2() As String = {"$", "€", "RUS", "BYN"}
        ComboBox8.Items.AddRange(d2)
        ComboBox7.Items.AddRange(d2)

        Dim d3() As String = {"Поручение с экспедицией", "Договор-поручение", "Договор-поручение эксп"}
        ComboBox12.Items.AddRange(d3)
        ComboBox13.Items.AddRange(d3)


        Dim d4() As String = {"БанкПоДок", "КалПоДок", "БанкПоВыг", "КалПоВыг"}
        ComboBox10.Items.AddRange(d4)
        ComboBox9.Items.AddRange(d4)


        TextBox5.Text = Now.ToShortDateString
        MaskedTextBox3.Text = "09:00"
        TextBox6.Text = Now.ToShortDateString
        MaskedTextBox4.Text = "09:00"

        Dim strsql2 As String = "SELECT ТипАвто FROM ТипАвто ORDER BY ТипАвто"
        Dim ds2 As DataTable = Selects(strsql2)
        Me.ComboBox11.AutoCompleteCustomSource.Clear()
        Me.ComboBox11.Items.Clear()
        For Each r As DataRow In ds2.Rows
            Me.ComboBox11.AutoCompleteCustomSource.Add(r.Item(0).ToString())
            Me.ComboBox11.Items.Add(r(0).ToString)
        Next

        MaskedTextBox1.Text = Now.ToShortDateString

    End Sub
    Private Sub ИзменЭксел(ByVal strsql As String, ByVal t As Integer, ByVal n As Integer, ByVal g As Integer)



        Dim xlapp As Microsoft.Office.Interop.Excel.Application
        xlapp = New Microsoft.Office.Interop.Excel.Application

        Dim XXX = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=U:\Офис\Рикманс\ДанныеРикманс.accdb; Persist Security Info=False;"
        Dim CON As New ADODB.Connection
        Dim RS As New ADODB.Recordset

        CON.ConnectionString = XXX
        CON.Open()
        RS.Open(strsql, CON)

        If t = 1 Then
            xlapp.Workbooks.Open("Z:\RICKMANS\" & ComboBox1.Text & "\" & ПутьРейса).Worksheets("ЗАК").Range("L3").CopyFromRecordset(RS)
            xlapp.Workbooks(ПутьРейса).Close(True)
            IO.File.Copy("Z:\RICKMANS\" & ComboBox1.Text & "\" & ПутьРейса, "Z:\RICKMANS\" & ComboBox1.Text & "\" & НомРес & " " & ComboBox3.Text & " " & n & " - " & ComboBox4.Text & " " & g & ".xlsm")
            IO.File.Delete("Z:\RICKMANS\" & ComboBox1.Text & "\" & ПутьРейса)
            ПутьРейса = НомРес & " " & ComboBox3.Text & " " & n & " - " & ComboBox4.Text & " " & g & ".xlsm"
        Else
            xlapp.Workbooks.Open("Z:\RICKMANS\" & ComboBox1.Text & "\" & ПутьРейса).Worksheets("ЗАК").Range("L6").CopyFromRecordset(RS)
            xlapp.Workbooks(ПутьРейса).Close(True)
            IO.File.Copy("Z:\RICKMANS\" & ComboBox1.Text & "\" & ПутьРейса, "Z:\RICKMANS\" & ComboBox1.Text & "\" & НомРес & " " & ComboBox3.Text & " " & g & " - " & ComboBox4.Text & " " & n & ".xlsm")
            IO.File.Delete("Z:\RICKMANS\" & ComboBox1.Text & "\" & ПутьРейса)
            ПутьРейса = НомРес & " " & ComboBox3.Text & " " & g & " - " & ComboBox4.Text & " " & n & ".xlsm"
        End If

        xlapp.Quit()
        releaseobject(xlapp)

        RS = Nothing
        CON.Close()

    End Sub
    Public Sub ИзменВДействРейсе(ByVal d As String)
        Me.Cursor = Cursors.WaitCursor
        Dim strsql As String = "SELECT * FROM РейсыКлиента WHERE НомерРейса=" & НомРес & ""
        Dim ds As DataTable = Selects(strsql)

        Dim strsql1 As String = "SELECT * FROM РейсыПеревозчика WHERE НомерРейса=" & НомРес & ""
        Dim ds1 As DataTable = Selects(strsql1)

        Select Case d
            Case "Клиент"
                ИзменЭксел(strsql, 1, ds.Rows(0).Item(3), ds1.Rows(0).Item(3))
            Case "Перевоз"
                ИзменЭксел(strsql1, 0, ds1.Rows(0).Item(3), ds.Rows(0).Item(3))
        End Select
        MessageBox.Show("Данные удачно изменены в файле эксель!", Рик)
        ПерегрЛист1()

        Me.Cursor = Cursors.Default

    End Sub


    Private Sub ПерегрЛист1()
        ListBox1.Items.Clear()
        Справки(CType(ComboBox1.Text, Integer))
        ListBox1.Items.AddRange(Files3)
        ListBox1.Text = Files3.Last
    End Sub
    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        ПерегрЛист1()
    End Sub

    Private Sub ListBox1_DoubleClick(sender As Object, e As EventArgs) Handles ListBox1.DoubleClick
        If ListBox1.SelectedIndex = -1 Then
            MessageBox.Show("Выберите документ для просмотра!", Рик, MessageBoxButtons.OK)
            Exit Sub
        End If

        If Not ListBox1.SelectedIndex = -1 Then
            For i = 0 To ListBox1.SelectedItems.Count - 1
                Try
                    Process.Start("Z:\RICKMANS\" & ComboBox1.Text & "\" & ListBox1.SelectedItems(i))
                Catch ex As Exception

                End Try

            Next

        End If
    End Sub
    Private Sub Вставка()
        strsql = ""
        strsql = "SELECT * FROM РейсыКлиента WHERE НомерРейса=" & НомРес & ""
        Dim ds1 As DataTable = Selects(strsql)

        Dim strsql2 As String = "SELECT * FROM РейсыПеревозчика WHERE НомерРейса=" & НомРес & ""
        Dim ds2 As DataTable = Selects(strsql2)
        Очистка()
        Try
            ds1.Rows(0).Item(1).ToString()
        Catch ex As Exception
            Exit Sub
        End Try


        ComboBox3.Text = ds1.Rows(0).Item(1).ToString
        ComboBox4.Text = ds2.Rows(0).Item(1).ToString
        TextBox1.Text = ds1.Rows(0).Item(17).ToString
        TextBox2.Text = ds2.Rows(0).Item(17).ToString
        ComboBox5.Text = ds1.Rows(0).Item(18).ToString
        ComboBox6.Text = ds2.Rows(0).Item(18).ToString
        ComboBox8.Text = ds1.Rows(0).Item(19).ToString
        ComboBox7.Text = ds2.Rows(0).Item(19).ToString
        TextBox4.Text = ds1.Rows(0).Item(20).ToString
        TextBox3.Text = ds2.Rows(0).Item(20).ToString
        MaskedTextBox1.Text = ds1.Rows(0).Item(24).ToString
        If ds1.Rows(0).Item(33).ToString = "1" Then
            ComboBox10.Text = "БанкПоДок"
        ElseIf ds1.Rows(0).Item(33).ToString = "3" Then
            ComboBox10.Text = "БанкПоВыг"
        Else
            ComboBox10.Text = ds1.Rows(0).Item(33).ToString
        End If

        If ds2.Rows(0).Item(30).ToString = "1" Then
            ComboBox9.Text = "БанкПоДок"
        ElseIf ds2.Rows(0).Item(30).ToString = "3" Then
            ComboBox9.Text = "БанкПоВыг"
        Else
            ComboBox9.Text = ds2.Rows(0).Item(30).ToString
        End If

        If ds1.Rows(0).Item(22).ToString = "1" Then
            ComboBox12.Text = "Договор-поручение"
        ElseIf ds1.Rows(0).Item(23).ToString = "1" Then
            ComboBox12.Text = "Договор-поручение эксп"
        ElseIf ds1.Rows(0).Item(25).ToString = "1" Then
            ComboBox12.Text = "Поручение с экспедицией"
        Else
            ComboBox12.Text = ""
        End If

        If ds2.Rows(0).Item(22).ToString = "1" Then
            ComboBox13.Text = "Договор-поручение"
        ElseIf ds2.Rows(0).Item(23).ToString = "1" Then
            ComboBox13.Text = "Договор-поручение эксп"
        ElseIf ds2.Rows(0).Item(25).ToString = "1" Then
            ComboBox13.Text = "Поручение с экспедицией"
        Else
            ComboBox13.Text = ""
        End If

        RichTextBox1.Text = ds1.Rows(0).Item(21).ToString
        RichTextBox2.Text = ds2.Rows(0).Item(21).ToString

        RichTextBox10.Text = ds1.Rows(0).Item(4).ToString
        TextBox5.Text = ds1.Rows(0).Item(5).ToString
        MaskedTextBox3.Text = ds1.Rows(0).Item(6).ToString
        TextBox6.Text = ds1.Rows(0).Item(7).ToString
        MaskedTextBox4.Text = ds1.Rows(0).Item(8).ToString

        RichTextBox3.Text = ds2.Rows(0).Item(9).ToString
        RichTextBox4.Text = ds2.Rows(0).Item(10).ToString
        RichTextBox7.Text = ds2.Rows(0).Item(11).ToString
        ComboBox11.Text = ds1.Rows(0).Item(12).ToString
        RichTextBox8.Text = ds2.Rows(0).Item(13).ToString
        RichTextBox9.Text = ds2.Rows(0).Item(14).ToString
        RichTextBox6.Text = ds2.Rows(0).Item(16).ToString
        RichTextBox5.Text = ds2.Rows(0).Item(15).ToString

        ДопФорма.TextBox4.Text = ds1.Rows(0).Item(27).ToString
        ДопФорма.TextBox3.Text = ds1.Rows(0).Item(26).ToString

        ДопФорма.TextBox6.Text = ds1.Rows(0).Item(30).ToString
        ДопФорма.TextBox5.Text = ds1.Rows(0).Item(31).ToString
        ДопФорма.MaskedTextBox1.Text = ds1.Rows(0).Item(32).ToString

        ДопФорма.MaskedTextBox2.Text = ds1.Rows(0).Item(28).ToString
        ДопФорма.TextBox10.Text = ds1.Rows(0).Item(29).ToString


    End Sub
    Private Sub ListBox1_Click(sender As Object, e As EventArgs) Handles ListBox1.Click
        If Not ListBox1.SelectedIndex = -1 Then
            Try
                НомРес = CType(Strings.Left(ListBox1.SelectedItem.ToString, 3), Integer)
                ПутьРейса = ListBox1.SelectedItem.ToString
            Catch ex As Exception

            End Try

        End If
        Вставка()
    End Sub
    Private Sub Очистка()
        RichTextBox10.Text = ""
        ComboBox3.Text = ""
        ComboBox4.Text = ""
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        ComboBox5.Text = ""
        ComboBox6.Text = ""
        ComboBox8.Text = ""
        ComboBox7.Text = ""
        ComboBox9.Text = ""
        ComboBox10.Text = ""
        ComboBox11.Text = ""
        ComboBox12.Text = ""
        ComboBox13.Text = ""
        RichTextBox1.Text = ""
        RichTextBox2.Text = ""
        RichTextBox3.Text = ""
        RichTextBox4.Text = ""
        RichTextBox5.Text = ""
        RichTextBox6.Text = ""
        RichTextBox7.Text = ""
        RichTextBox8.Text = ""
        RichTextBox9.Text = ""
        TextBox5.Text = ""
        MaskedTextBox3.Text = ""
        MaskedTextBox4.Text = ""
        TextBox6.Text = ""

    End Sub

    Private Sub ComboBox5_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox5.SelectedIndexChanged
        Select Case ComboBox5.SelectedItem
            Case "евро"
                ComboBox8.Enabled = True
            Case "доллар"
                ComboBox8.Enabled = True
            Case "росс.рубль"
                ComboBox8.Enabled = True
            Case "рубль"
                ComboBox8.Enabled = False
        End Select

    End Sub

    Private Sub ComboBox6_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox6.SelectedIndexChanged
        Select Case ComboBox6.SelectedItem
            Case "евро"
                ComboBox7.Enabled = True
            Case "доллар"
                ComboBox7.Enabled = True
            Case "росс.рубль"
                ComboBox7.Enabled = True
            Case "рубль"
                ComboBox7.Enabled = False
        End Select
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked = True Then
            RichTextBox2.Text = RichTextBox1.Text
        End If
    End Sub

    Private Sub ComboBox3_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox3.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.TextBox1.Focus()
        End If
    End Sub

    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.ComboBox5.Focus()
        End If
    End Sub

    Private Sub ComboBox5_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox5.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            If ComboBox8.Enabled = False Then
                TextBox4.Focus()
            Else
                Me.ComboBox8.Focus()
            End If

        End If
    End Sub

    Private Sub ComboBox8_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox8.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.TextBox4.Focus()
        End If
    End Sub

    Private Sub TextBox4_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox4.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.ComboBox10.Focus()
        End If
    End Sub

    Private Sub ComboBox10_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox10.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.ComboBox12.Focus()
        End If
    End Sub

    Private Sub ComboBox12_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox12.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.RichTextBox1.Focus()
        End If
    End Sub

    Private Sub RichTextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles RichTextBox1.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.ComboBox4.Focus()
        End If
    End Sub

    Private Sub ComboBox4_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox4.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.TextBox2.Focus()
        End If
    End Sub

    Private Sub TextBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox2.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.ComboBox6.Focus()
        End If
    End Sub

    Private Sub ComboBox6_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox6.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True

            If ComboBox7.Enabled = False Then
                TextBox3.Focus()
            Else
                ComboBox7.Focus()
            End If


        End If
    End Sub

    Private Sub ComboBox7_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox7.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.TextBox3.Focus()
        End If
    End Sub

    Private Sub TextBox3_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox3.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.ComboBox9.Focus()
        End If
    End Sub

    Private Sub ComboBox9_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox9.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.ComboBox13.Focus()
        End If
    End Sub

    Private Sub ComboBox13_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox13.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.RichTextBox2.Focus()
        End If
    End Sub

    Private Sub RichTextBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles RichTextBox2.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.RichTextBox10.Focus()
        End If
    End Sub

    Private Sub RichTextBox10_KeyDown(sender As Object, e As KeyEventArgs) Handles RichTextBox10.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.TextBox5.Focus()
        End If
    End Sub

    Private Sub TextBox5_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox5.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.MaskedTextBox3.Focus()
        End If
    End Sub

    Private Sub MaskedTextBox3_KeyDown(sender As Object, e As KeyEventArgs) Handles MaskedTextBox3.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.TextBox6.Focus()
        End If
    End Sub

    Private Sub TextBox6_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox6.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.MaskedTextBox4.Focus()
        End If
    End Sub

    Private Sub MaskedTextBox4_KeyDown(sender As Object, e As KeyEventArgs) Handles MaskedTextBox4.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.RichTextBox3.Focus()
        End If
    End Sub

    Private Sub RichTextBox3_KeyDown(sender As Object, e As KeyEventArgs) Handles RichTextBox3.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.RichTextBox4.Focus()
        End If
    End Sub

    Private Sub RichTextBox4_KeyDown(sender As Object, e As KeyEventArgs) Handles RichTextBox4.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.RichTextBox7.Focus()
        End If
    End Sub

    Private Sub RichTextBox7_KeyDown(sender As Object, e As KeyEventArgs) Handles RichTextBox7.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.ComboBox11.Focus()
        End If
    End Sub

    Private Sub ComboBox11_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox11.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.RichTextBox8.Focus()
        End If
    End Sub

    Private Sub RichTextBox8_KeyDown(sender As Object, e As KeyEventArgs) Handles RichTextBox8.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.RichTextBox9.Focus()
        End If
    End Sub

    Private Sub RichTextBox9_KeyDown(sender As Object, e As KeyEventArgs) Handles RichTextBox9.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.RichTextBox6.Focus()
        End If
    End Sub

    Private Sub RichTextBox6_KeyDown(sender As Object, e As KeyEventArgs) Handles RichTextBox6.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.RichTextBox5.Focus()
        End If
    End Sub

    Private Sub RichTextBox5_KeyDown(sender As Object, e As KeyEventArgs) Handles RichTextBox5.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.Button1.Focus()
        End If
    End Sub

    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox2.CheckedChanged
        If CheckBox2.Checked = True Then
            Очистка()
            CheckBox2.CheckState = False
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If ComboBox3.Text = "" Or TextBox1.Text = "" Then
            MessageBox.Show("Выберите рейс!", Рик)
            Exit Sub
        End If
        ДопФорма.ShowDialog()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        ВодитДан.ShowDialog()
    End Sub
    Private Function Проверка()
        If ComboBox3.Text = "" Then
            MessageBox.Show("Выберите клиента!", Рик)
            Return 1
        End If

        If ComboBox4.Text = "" Then
            MessageBox.Show("Выберите перевозчика!", Рик)
            Return 1
        End If


        If TextBox1.Text = "" Or TextBox2.Text = "" Then
            MessageBox.Show("Выберите фрахт!", Рик)
            Return 1
        End If

        If ComboBox5.Text = "" Or ComboBox6.Text = "" Then
            MessageBox.Show("Выберите валюту!", Рик)
            Return 1
        End If

        If ComboBox11.Text = "" Then
            MessageBox.Show("Выберите тип автомомбиля!", Рик)
            Return 1
        End If

        If Not ComboBox5.Text = "Рубль" And ComboBox8.Text = "" Then
            MessageBox.Show("Выберите валюту платежа!", Рик)
            Return 1
        End If

        If Not ComboBox6.Text = "Рубль" And ComboBox7.Text = "" Then
            MessageBox.Show("Выберите валюту платежа!", Рик)
            Return 1
        End If

        If TextBox3.Text = "" Or TextBox4.Text = "" Then
            MessageBox.Show("Выберите срок оплаты!", Рик)
            Return 1
        End If

        If ComboBox10.Text = "" Or ComboBox9.Text = "" Then
            MessageBox.Show("Выберите условия оплаты!", Рик)
            Return 1
        End If

        If RichTextBox1.Text = "" Or RichTextBox2.Text = "" Then
            MessageBox.Show("Заполните дополнительные условия!", Рик)
            Return 1
        End If

        If RichTextBox10.Text = "" Then
            MessageBox.Show("Заполните маршрут!", Рик)
            Return 1
        End If

        If RichTextBox7.Text = "" Then
            MessageBox.Show("Заполните наменование груза!", Рик)
            Return 1
        End If



        If RichTextBox3.Text = "" Or RichTextBox6.Text = "" Then
            MessageBox.Show("Заполните место загрузки/разгрузки!", Рик)
            Return 1
        End If

        If TextBox5.Text = "" Or TextBox6.Text = "" Then
            MessageBox.Show("Заполните даты загрузки/разгрузки!", Рик)
            Return 1
        End If

        If MaskedTextBox3.Text = "" Or MaskedTextBox4.Text = "" Then
            MessageBox.Show("Заполните время загрузки/разгрузки!", Рик)
            Return 1
        End If

        If MaskedTextBox3.MaskCompleted = False Or MaskedTextBox4.MaskCompleted = False Then
            MessageBox.Show("Заполните время загрузки/разгрузки!", Рик)
            Return 1
        End If

        If MaskedTextBox1.MaskCompleted = False Then
            MessageBox.Show("Выберите дату поручения!", Рик)
            Return 1
        End If

        Return 0
    End Function
    Private Sub ПровСледРейсКлиент()
        Dim strsql, strsql1 As String
        Dim ds As DataTable


        Try
            ds.Clear()
        Catch ex As Exception

        End Try

        Dim strsql5 As String = "SELECT Договор,Дата,Должность,ФИОРуководителя FROM Клиент WHERE НазваниеОрганизации='" & ComboBox3.Text & "'"
        Dim ds5 As DataTable = Selects(strsql5)




        If ComboBox3.Text = "Виталюр" Then
            strsql = "SELECT MAX(КоличРейсов) FROM РейсыКлиента WHERE НазвОрганизации = 'Виталюр'  AND ДатаПодачиПодЗагрузку Like '%" & Now.Year & "%' GROUP BY НазвОрганизации "
            ds = Selects(strsql)
        Else
            strsql1 = "SELECT MAX(КоличРейсов) FROM [РейсыКлиента] WHERE НазвОрганизации = '" & ComboBox3.Text & "' GROUP BY НазвОрганизации "
            ds = Selects(strsql1)
        End If


        Try
            СлПорРейсКл = ds.Rows(0).Item(0) + 1
        Catch ex As Exception

            Nm = ""
            Nm = ComboBox3.Text

            If ds5.Rows(0).Item(0).ToString = "" Or ds5.Rows(0).Item(1).ToString = "" Then
                ПорНомер.GroupBox1.Enabled = True
            End If
            If ds5.Rows(0).Item(2).ToString = "" Or ds5.Rows(0).Item(3).ToString = "" Then
                ПорНомер.GroupBox2.Enabled = True
            End If
            ПорНомер.TextBox1.Enabled = True
            ПорНомер.ListBox1.Enabled = True
            pro = 1
            ПорНомер.ShowDialog()
            If Отмена = 1 Then Exit Sub
            'СлПорРейсКл = CType((ПорНомер.TextBox1.Text), Integer) + 1
            Пров = 1
        End Try

        If Пров = 0 Then
            Dim f, g As Integer
            If ds5.Rows(0).Item(0).ToString = "" Or ds5.Rows(0).Item(1).ToString = "" Then
                ПорНомер.GroupBox1.Enabled = True
                f = 1
            End If
            If ds5.Rows(0).Item(2).ToString = "" Or ds5.Rows(0).Item(3).ToString = "" Then
                ПорНомер.GroupBox2.Enabled = True
                g = 1
            End If
            If f = 1 Or g = 1 Then
                pro = 1
                ПорНомер.ShowDialog()
                If Отмена = 1 Then Exit Sub
            End If
        End If
    End Sub
    Private Sub ПровСледРейсПер()

        Dim strsql6 As String = "SELECT Договор,Дата,Должность,ФИОРуководителя FROM Перевозчики WHERE Названиеорганизации='" & ComboBox4.Text & "'"
        Dim ds6 As DataTable = Selects(strsql6)


        Dim strsql2 As String = "SELECT MAX(КоличРейсов) FROM РейсыПеревозчика WHERE НазвОрганизации = '" & ComboBox4.Text & "' GROUP BY НазвОрганизации"
        Dim ds2 As DataTable = Selects(strsql2)

        Try
            СлПорРейсПер = ds2.Rows(0).Item(0) + 1
        Catch ex As Exception
            Nm = ""
            Nm = ComboBox4.Text
            pro = 0

            If ds6.Rows(0).Item(0).ToString = "" Or ds6.Rows(0).Item(1).ToString = "" Then
                ПорНомер.GroupBox1.Enabled = True
            End If
            If ds6.Rows(0).Item(2).ToString = "" Or ds6.Rows(0).Item(3).ToString = "" Then
                ПорНомер.GroupBox2.Enabled = True
            End If
            ПорНомер.TextBox1.Enabled = True
            ПорНомер.ListBox1.Enabled = True
            ПорНомер.ShowDialog()
            If Отмена = 1 Then Exit Sub
            'СлПорРейсПер = CType((ПорНомер.TextBox1.Text), Integer) + 1
            Пров = 1
        End Try

        If Пров = 0 Then
            Dim f, g As Integer
            If ds6.Rows(0).Item(0).ToString = "" Or ds6.Rows(0).Item(1).ToString = "" Then
                ПорНомер.GroupBox1.Enabled = True
                f = 1
            End If
            If ds6.Rows(0).Item(2).ToString = "" Or ds6.Rows(0).Item(3).ToString = "" Then
                ПорНомер.GroupBox2.Enabled = True
                g = 1
            End If

            If f = 1 Or g = 1 Then
                pro = 0
                ПорНомер.ShowDialog()
                If Отмена = 1 Then Exit Sub
            End If
        End If
    End Sub
    Private Sub ComB13()

        ДогПор = "0"
        ДогПорЭксп = "0"
        ПорЭксп = "0"

        If ComboBox13.Text = "" Then
            ДогПор = "0"
            ДогПорЭксп = "0"
            ПорЭксп = "0"
        ElseIf ComboBox13.Text = "Договор-поручение" Then
            ДогПор = "1"
            ДогПорЭксп = "0"
            ПорЭксп = "0"
        ElseIf ComboBox13.Text = "Договор-поручение эксп" Then
            ДогПор = "0"
            ДогПорЭксп = "1"
            ПорЭксп = "0"
        ElseIf ComboBox13.Text = "Поручение с экспедицией" Then
            ДогПор = "0"
            ДогПорЭксп = "0"
            ПорЭксп = "1"
        Else
            ДогПор = "0"
            ДогПорЭксп = "0"
            ПорЭксп = "0"
        End If
    End Sub
    Private Sub ComB12()
        ДогПор = "0"
        ДогПорЭксп = "0"
        ПорЭксп = "0"


        If ComboBox12.Text = "" Then
            ДогПор = "0"
            ДогПорЭксп = "0"
            ПорЭксп = "0"
        ElseIf ComboBox12.Text = "Договор-поручение" Then
            ДогПор = "1"
            ДогПорЭксп = "0"
            ПорЭксп = "0"
        ElseIf ComboBox12.Text = "Договор-поручение эксп" Then
            ДогПор = "0"
            ДогПорЭксп = "1"
            ПорЭксп = "0"
        ElseIf ComboBox12.Text = "Поручение с экспедицией" Then
            ДогПор = "0"
            ДогПорЭксп = "0"
            ПорЭксп = "1"
        Else
            ДогПор = "0"
            ДогПорЭксп = "0"
            ПорЭксп = "0"
        End If
    End Sub
    Private Sub НовыйРейс()
        Dim strsql, strsql3 As String
        Dim ds1 As DataTable

        ПровСледРейсКлиент()
        If Отмена = 1 Then Exit Sub

        Пров = 0
        strsql = ""
        strsql = "SELECT MAX(НомерРейса) FROM РейсыКлиента"
        ds1 = Selects(strsql)

        СлРейс = ds1.Rows(0).Item(0) + 1

        ПровСледРейсПер()
        If Отмена = 1 Then Exit Sub

        ComB13()
        strsql3 = "INSERT INTO РейсыПеревозчика(НазвОрганизации,НомерРейса,КоличРейсов,Маршрут,ДатаПодачиПодЗагрузку,ВремяПодачи,ДатаПодачиПодРастаможку,
ВремяПодачиВыгРаст,ТочныйАдресЗагрузки,АдресЗатаможки,НаименованиеГруза,ТипТрСредства,НомерАвтомобиля,Водитель,
ТочнАдресРаста,ТочнАдресРазгр,СтоимостьФрахта,Валюта,ВалютаПлатежа,СрокОплаты,ДопУсловия,
ДогПор,ДогПорЭксп,ДатаПоручения,ПорЭксп,УсловияОплаты)
VALUES('" & ComboBox4.Text & "'," & СлРейс & "," & СлПорРейсПер & ",'" & Trim(RichTextBox10.Text) & "','" & TextBox5.Text & "','" & MaskedTextBox3.Text & "','" & TextBox6.Text & "',
'" & MaskedTextBox4.Text & "','" & Trim(RichTextBox3.Text) & "','" & Trim(RichTextBox4.Text) & "','" & Trim(RichTextBox7.Text) & "','" & ComboBox11.Text & "','" & Trim(RichTextBox8.Text) & "','" & Trim(RichTextBox9.Text) & "',
'" & Trim(RichTextBox5.Text) & "','" & Trim(RichTextBox6.Text) & "','" & TextBox2.Text & "','" & ComboBox6.Text & "','" & ComboBox7.Text & "','" & TextBox3.Text & "','" & Trim(RichTextBox2.Text) & "',
'" & ДогПор & "','" & ДогПорЭксп & "','" & MaskedTextBox1.Text & "','" & ПорЭксп & "','" & ComboBox9.Text & "')"
        Updates(strsql3)

        ComB12()
        Dim strsql4 As String = "INSERT INTO РейсыКлиента(НазвОрганизации,НомерРейса,КоличРейсов,Маршрут,ДатаПодачиПодЗагрузку,ВремяПодачи,ДатаПодачиПодРастаможку,
ВремяПодачиВыгРаст,ТочныйАдресЗагрузки,АдресЗатаможки,НаименованиеГруза,ТипТрСредства,НомерАвтомобиля,Водитель,
ТочнАдресРаста,ТочнАдресРазгр,СтоимостьФрахта,Валюта,ВалютаПлатежа,СрокОплаты,ДопУсловия,
ДогПор,ДогПорЭксп,ДатаПоручения,ПорЭксп,УсловияОплаты,Год)
VALUES('" & ComboBox3.Text & "'," & СлРейс & "," & СлПорРейсКл & ",'" & Trim(RichTextBox10.Text) & "','" & TextBox5.Text & "','" & MaskedTextBox3.Text & "','" & TextBox6.Text & "',
'" & MaskedTextBox4.Text & "','" & Trim(RichTextBox3.Text) & "','" & Trim(RichTextBox4.Text) & "','" & Trim(RichTextBox7.Text) & "','" & ComboBox11.Text & "','" & Trim(RichTextBox8.Text) & "','" & Trim(RichTextBox9.Text) & "',
'" & Trim(RichTextBox5.Text) & "','" & Trim(RichTextBox6.Text) & "','" & TextBox1.Text & "','" & ComboBox5.Text & "','" & ComboBox8.Text & "','" & TextBox4.Text & "','" & Trim(RichTextBox1.Text) & "',
'" & ДогПор & "','" & ДогПорЭксп & "','" & MaskedTextBox1.Text & "','" & ПорЭксп & "','" & ComboBox10.Text & "','" & Now.ToShortDateString & "')"
        Updates(strsql4)


    End Sub
    Private Sub Доки()
        Me.Cursor = Cursors.WaitCursor
        'If conn.State = ConnectionState.Open Then
        '    conn.Close()
        'End If

        Dim xlapp As Microsoft.Office.Interop.Excel.Application
        'Dim xlworkbook As Microsoft.Office.Interop.Excel.Workbook
        'Dim xlworkbooks As Microsoft.Office.Interop.Excel.Workbooks
        'Dim xlworksheet As Microsoft.Office.Interop.Excel.Worksheet
        'Dim xlworksheets As Microsoft.Office.Interop.Excel.Workbooks
        xlapp = New Microsoft.Office.Interop.Excel.Application
        'xlworkbook = xlapp.Workbooks.Add("Z:\RICKMANS\496 Ивановский_8 - Петровский_1.xlsm")
        'xlworksheet = xlworkbook.Worksheets("ЗАК")

        Dim XXX = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=U:\Офис\Рикманс\ДанныеРикманс.accdb; Persist Security Info=False;"
        Dim CON As New ADODB.Connection
        Dim RS As New ADODB.Recordset
        Dim RS1 As New ADODB.Recordset
        Dim RS2 As New ADODB.Recordset
        Dim RS3 As New ADODB.Recordset

        Dim strSQL, strSQL1, strSQL2, strSQL3 As String
        strSQL = "select * from РейсыКлиента WHERE НомерРейса=" & СлРейс & ""
        strSQL1 = "select * from РейсыПеревозчика WHERE НомерРейса=" & СлРейс & ""
        strSQL2 = "select * from Клиент WHERE НазваниеОрганизации='" & ComboBox3.Text & "'"
        strSQL3 = "select * from перевозчики WHERE Названиеорганизации='" & ComboBox4.Text & "'"
        'Excel_obj = CreateObject("Excel.Application")
        'Excel_obj.visible = True
        'Excel_Doc = Excel_obj.workbooks.add   '("Z:\RICKMANS\496 Ивановский_8 - Петровский_1.xlsm")
        'RS = CreateObject("ADODB.Recordset")
        'RS.Open(strSQL, CON)

        CON.ConnectionString = XXX
        CON.Open()
        RS.Open(strSQL, CON)

        'xlworksheet.Range("L3").CopyFromRecordset(RS)

        RS1.Open(strSQL1, CON)
        'xlworksheet.Range("L6").CopyFromRecordset(RS1)

        RS2.Open(strSQL2, CON)
        'xlworksheet.Range("L4").CopyFromRecordset(RS2)

        RS3.Open(strSQL3, CON)
        'xlworksheet.Range("L5").CopyFromRecordset(RS3)
        xlapp.Workbooks.Open("Z:\RICKMANS\496 Ивановский_8 - Петровский_1.xlsm").Worksheets("ЗАК").Range("L3").CopyFromRecordset(RS)
        xlapp.Workbooks("496 Ивановский_8 - Петровский_1.xlsm").Worksheets("ЗАК").Range("L6").CopyFromRecordset(RS1)
        xlapp.Workbooks("496 Ивановский_8 - Петровский_1.xlsm").Worksheets("ЗАК").Range("L4").CopyFromRecordset(RS2)
        xlapp.Workbooks("496 Ивановский_8 - Петровский_1.xlsm").Worksheets("ЗАК").Range("L5").CopyFromRecordset(RS3)
        xlapp.Workbooks("496 Ивановский_8 - Петровский_1.xlsm").Close(True)

        IO.File.Copy("Z:\RICKMANS\496 Ивановский_8 - Петровский_1.xlsm", "Z:\RICKMANS\" & Now.Year & "\" & СлРейс & " " & ComboBox3.Text & " " & СлПорРейсКл & " - " & ComboBox4.Text & " " & СлПорРейсПер & ".xlsm")

        'xlworksheet.SaveAs("Z:\RICKMANS\" & Now.Year & "\" & СлРейс & " " & ComboBox3.Text & " " & СлПорРейсКл & " - " & ComboBox4.Text & " " & СлПорРейсПер & ".xlsm")
        'xlworkbooks("Z:\RICKMANS\496 Ивановский_8 - Петровский_1.xlsm").Close(True)

        'xlworkbook.Close()
        'xlworksheet("Z:\RICKMANS\496 Ивановский_8 - Петровский_1.xlsm").Close(True)

        xlapp.Quit()
        releaseobject(xlapp)

        'releaseobject(xlworkbook)
        'releaseobject(xlworksheet)


        'Excel_Doc("496 Ивановский_8 - Петровский_1.xlsm").Close

        'Excel_Doc.Range("a1").CopyFromRecordset(RS)   'копируем рекордсет на лист Excel без заморочек
        RS = Nothing
        RS1 = Nothing
        RS2 = Nothing
        RS3 = Nothing
        CON.Close()

        Me.Cursor = Cursors.Default
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        ОтКогоИзмен = 0
        If ComboBox3.Text = "" Then
            MessageBox.Show("Выберите рейс для внесения изменений!", Рик)
            Exit Sub
        End If
        ОтКогоИзмен = 1
        ИзменПорНомКлиент.ShowDialog()
    End Sub

    Private Sub РедакцияСтарогоРейса()
        Dim strsql, strsql1, strsql2, strsql3 As String
        strsql = "SELECT НазвОрганизации FROM РейсыКлиента WHERE НомерРейса=" & НомРес & ""
        Dim ds As DataTable = Selects(strsql)
        If Not ds.Rows(0).Item(0).ToString = ComboBox3.Text Then
            ПровСледРейсКлиент()
            If Отмена = 1 Then Exit Sub
            ComB12()
            strsql1 = "UPDATE РейсыКлиента SET НазвОрганизации='" & ComboBox3.Text & "', КоличРейсов=" & СлПорРейсКл & ", Маршрут='" & Trim(RichTextBox10.Text) & "',
ДатаПодачиПодЗагрузку='" & TextBox5.Text & "', ВремяПодачи='" & MaskedTextBox3.Text & "', ДатаПодачиПодРастаможку='" & TextBox6.Text & "', ВремяПодачиВыгРаст='" & MaskedTextBox4.Text & "',
ТочныйАдресЗагрузки='" & Trim(RichTextBox3.Text) & "', АдресЗатаможки='" & Trim(RichTextBox4.Text) & "', НаименованиеГруза='" & Trim(RichTextBox7.Text) & "', ТипТрСредства='" & ComboBox11.Text & "',
НомерАвтомобиля='" & Trim(RichTextBox8.Text) & "',Водитель ='" & Trim(RichTextBox9.Text) & "', ТочнАдресРаста='" & Trim(RichTextBox5.Text) & "', ТочнАдресРазгр='" & Trim(RichTextBox6.Text) & "',
СтоимостьФрахта='" & TextBox1.Text & "', Валюта='" & ComboBox5.Text & "', ВалютаПлатежа='" & ComboBox8.Text & "', СрокОплаты='" & TextBox4.Text & "', УсловияОплаты='" & ComboBox10.Text & "',
ДогПор='" & ДогПор & "',ДогПорЭксп='" & ДогПорЭксп & "',ПорЭксп= '" & ПорЭксп & "', ДопУсловия='" & Trim(RichTextBox1.Text) & "', ДатаПоручения='" & MaskedTextBox1.Text & "'
WHERE НомерРейса=" & НомРес & ""
            Updates(strsql1)
            ПрИзмНазКл = True
        Else
            ComB12()
            strsql1 = "UPDATE РейсыКлиента SET Маршрут='" & Trim(RichTextBox10.Text) & "',
ДатаПодачиПодЗагрузку='" & TextBox5.Text & "', ВремяПодачи='" & MaskedTextBox3.Text & "', ДатаПодачиПодРастаможку='" & TextBox6.Text & "', ВремяПодачиВыгРаст='" & MaskedTextBox4.Text & "',
ТочныйАдресЗагрузки='" & Trim(RichTextBox3.Text) & "', АдресЗатаможки='" & Trim(RichTextBox4.Text) & "', НаименованиеГруза='" & Trim(RichTextBox7.Text) & "', ТипТрСредства='" & ComboBox11.Text & "',
НомерАвтомобиля='" & Trim(RichTextBox8.Text) & "',Водитель ='" & Trim(RichTextBox9.Text) & "', ТочнАдресРаста='" & Trim(RichTextBox5.Text) & "', ТочнАдресРазгр='" & Trim(RichTextBox6.Text) & "',
СтоимостьФрахта='" & TextBox1.Text & "', Валюта='" & ComboBox5.Text & "', ВалютаПлатежа='" & ComboBox8.Text & "', СрокОплаты='" & TextBox4.Text & "', УсловияОплаты='" & ComboBox10.Text & "',
ДогПор='" & ДогПор & "', ДогПорЭксп='" & ДогПорЭксп & "',ПорЭксп= '" & ПорЭксп & "', ДопУсловия='" & Trim(RichTextBox1.Text) & "', ДатаПоручения='" & MaskedTextBox1.Text & "'
WHERE НомерРейса=" & НомРес & ""
            Updates(strsql1)
        End If

        strsql2 = "SELECT НазвОрганизации FROM РейсыПеревозчика WHERE НомерРейса=" & НомРес & ""
        Dim ds1 As DataTable = Selects(strsql2)
        If Not ds1.Rows(0).Item(0).ToString = ComboBox4.Text Then
            ПровСледРейсПер()
            If Отмена = 1 Then Exit Sub
            ComB13()
            strsql3 = "UPDATE РейсыПеревозчика SET НазвОрганизации='" & ComboBox4.Text & "', КоличРейсов=" & СлПорРейсПер & ", Маршрут='" & Trim(RichTextBox10.Text) & "',
ДатаПодачиПодЗагрузку='" & TextBox5.Text & "', ВремяПодачи='" & MaskedTextBox3.Text & "', ДатаПодачиПодРастаможку='" & TextBox6.Text & "', ВремяПодачиВыгРаст='" & MaskedTextBox4.Text & "',
ТочныйАдресЗагрузки='" & Trim(RichTextBox3.Text) & "', АдресЗатаможки='" & Trim(RichTextBox4.Text) & "', НаименованиеГруза='" & Trim(RichTextBox7.Text) & "', ТипТрСредства='" & ComboBox11.Text & "',
НомерАвтомобиля='" & Trim(RichTextBox8.Text) & "',Водитель ='" & Trim(RichTextBox9.Text) & "', ТочнАдресРаста='" & Trim(RichTextBox5.Text) & "', ТочнАдресРазгр='" & Trim(RichTextBox6.Text) & "',
СтоимостьФрахта='" & TextBox2.Text & "', Валюта='" & ComboBox6.Text & "', ВалютаПлатежа='" & ComboBox7.Text & "', СрокОплаты='" & TextBox3.Text & "', УсловияОплаты='" & ComboBox9.Text & "',
ДогПор='" & ДогПор & "', ДогПорЭксп='" & ДогПорЭксп & "',ПорЭксп= '" & ПорЭксп & "', ДопУсловия='" & Trim(RichTextBox2.Text) & "', ДатаПоручения='" & MaskedTextBox1.Text & "'
WHERE НомерРейса=" & НомРес & ""
            Updates(strsql3)
            ПрИзмНазПер = True
        Else
            ComB13()
            strsql3 = "UPDATE РейсыПеревозчика SET Маршрут='" & Trim(RichTextBox10.Text) & "',
ДатаПодачиПодЗагрузку='" & TextBox5.Text & "', ВремяПодачи='" & MaskedTextBox3.Text & "', ДатаПодачиПодРастаможку='" & TextBox6.Text & "', ВремяПодачиВыгРаст='" & MaskedTextBox4.Text & "',
ТочныйАдресЗагрузки='" & Trim(RichTextBox3.Text) & "', АдресЗатаможки='" & Trim(RichTextBox4.Text) & "', НаименованиеГруза='" & Trim(RichTextBox7.Text) & "', ТипТрСредства='" & ComboBox11.Text & "',
НомерАвтомобиля='" & Trim(RichTextBox8.Text) & "',Водитель ='" & Trim(RichTextBox9.Text) & "', ТочнАдресРаста='" & Trim(RichTextBox5.Text) & "', ТочнАдресРазгр='" & Trim(RichTextBox6.Text) & "',
СтоимостьФрахта='" & TextBox2.Text & "', Валюта='" & ComboBox6.Text & "', ВалютаПлатежа='" & ComboBox7.Text & "', СрокОплаты='" & TextBox3.Text & "', УсловияОплаты='" & ComboBox9.Text & "',
ДогПор='" & ДогПор & "', ДогПорЭксп='" & ДогПорЭксп & "',ПорЭксп= '" & ПорЭксп & "', ДопУсловия='" & Trim(RichTextBox2.Text) & "', ДатаПоручения='" & MaskedTextBox1.Text & "'
WHERE НомерРейса=" & НомРес & ""
            Updates(strsql3)
        End If
        СлРейс = НомРес
        ДокиОбновление()

    End Sub
    Public Sub ДокиОбновление()
        Me.Cursor = Cursors.WaitCursor
        Dim xlapp As Microsoft.Office.Interop.Excel.Application
        xlapp = New Microsoft.Office.Interop.Excel.Application

        Dim XXX = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=U:\Офис\Рикманс\ДанныеРикманс.accdb; Persist Security Info=False;"
        Dim CON As New ADODB.Connection
        Dim RS As New ADODB.Recordset
        Dim RS1 As New ADODB.Recordset
        Dim RS2 As New ADODB.Recordset
        Dim RS3 As New ADODB.Recordset

        Dim strSQL, strSQL1, strSQL2, strSQL3 As String
        strSQL = "select * from РейсыКлиента WHERE НомерРейса=" & СлРейс & ""
        strSQL1 = "select * from РейсыПеревозчика WHERE НомерРейса=" & СлРейс & ""
        strSQL2 = "select * from Клиент WHERE НазваниеОрганизации='" & ComboBox3.Text & "'"
        strSQL3 = "select * from Перевозчики WHERE Названиеорганизации='" & ComboBox4.Text & "'"

        CON.ConnectionString = XXX
        CON.Open()
        RS.Open(strSQL, CON)
        RS1.Open(strSQL1, CON)
        RS2.Open(strSQL2, CON)
        RS3.Open(strSQL3, CON)

        xlapp.Workbooks.Open("Z:\RICKMANS\" & ComboBox1.Text & "\" & ПутьРейса).Worksheets("ЗАК").Range("L3").CopyFromRecordset(RS)
        xlapp.Workbooks(ПутьРейса).Worksheets("ЗАК").Range("L6").CopyFromRecordset(RS1)
        xlapp.Workbooks(ПутьРейса).Worksheets("ЗАК").Range("L4").CopyFromRecordset(RS2)
        xlapp.Workbooks(ПутьРейса).Worksheets("ЗАК").Range("L5").CopyFromRecordset(RS3)
        xlapp.Workbooks(ПутьРейса).Close(True)

        Dim ds1 As DataTable = Selects(strSQL)
        Dim ds2 As DataTable = Selects(strSQL1)

        If ПрИзмНазКл = True Then
            Dim G As String = ПутьРейса
            Try
                IO.File.Copy("Z:\RICKMANS\" & ComboBox1.Text & "\" & ПутьРейса, "Z:\RICKMANS\" & ComboBox1.Text & "\" & НомРес & " " & ComboBox3.Text & " " & ds1.Rows(0).Item(3) & " - " & ComboBox4.Text & " " & ds2.Rows(0).Item(3) & ".xlsm")
            Catch ex As Exception
                IO.File.Delete("Z:\RICKMANS\" & ComboBox1.Text & "\" & НомРес & " " & ComboBox3.Text & " " & ds1.Rows(0).Item(3) & " - " & ComboBox4.Text & " " & ds2.Rows(0).Item(3) & ".xlsm")
                IO.File.Copy("Z:\RICKMANS\" & ComboBox1.Text & "\" & ПутьРейса, "Z:\RICKMANS\" & ComboBox1.Text & "\" & НомРес & " " & ComboBox3.Text & " " & ds1.Rows(0).Item(3) & " - " & ComboBox4.Text & " " & ds2.Rows(0).Item(3) & ".xlsm")
            End Try

            IO.File.Delete("Z:\RICKMANS\" & ComboBox1.Text & "\" & G)
            ПрИзмНазКл = False
            ПутьРейса = НомРес & " " & ComboBox3.Text & " " & ds1.Rows(0).Item(3) & " - " & ComboBox4.Text & " " & ds2.Rows(0).Item(3) & ".xlsm"
        End If

        If ПрИзмНазПер = True Then
            Dim G As String = ПутьРейса
            Try
                IO.File.Copy("Z:\RICKMANS\" & ComboBox1.Text & "\" & ПутьРейса, "Z:\RICKMANS\" & ComboBox1.Text & "\" & НомРес & " " & ComboBox3.Text & " " & ds1.Rows(0).Item(3) & " - " & ComboBox4.Text & " " & ds2.Rows(0).Item(3) & ".xlsm")
            Catch ex As Exception
                IO.File.Delete("Z:\RICKMANS\" & ComboBox1.Text & "\" & НомРес & " " & ComboBox3.Text & " " & ds1.Rows(0).Item(3) & " - " & ComboBox4.Text & " " & ds2.Rows(0).Item(3) & ".xlsm")
                IO.File.Copy("Z:\RICKMANS\" & ComboBox1.Text & "\" & ПутьРейса, "Z:\RICKMANS\" & ComboBox1.Text & "\" & НомРес & " " & ComboBox3.Text & " " & ds1.Rows(0).Item(3) & " - " & ComboBox4.Text & " " & ds2.Rows(0).Item(3) & ".xlsm")
            End Try

            IO.File.Delete("Z:\RICKMANS\" & ComboBox1.Text & "\" & G)
            ПрИзмНазПер = False
            ПутьРейса = НомРес & " " & ComboBox3.Text & " " & ds1.Rows(0).Item(3) & " - " & ComboBox4.Text & " " & ds2.Rows(0).Item(3) & ".xlsm"
        End If

        xlapp.Quit()
        releaseobject(xlapp)

        RS = Nothing
        RS1 = Nothing
        RS2 = Nothing
        RS3 = Nothing
        CON.Close()
        Me.Cursor = Cursors.Default

    End Sub
    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        If ComboBox4.Text = "" Then
            MessageBox.Show("Выберите рейс для внесения изменений!", Рик)
            Exit Sub
        End If
        ОтКогоИзмен = 0
        ИзменПорНомКлиент.ShowDialog()
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click

        If MessageBox.Show("Удалить " & НомРес & " Рейс?", Рик, MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.No Then
            Exit Sub
        End If

        Dim strsql As String = "DELETE * FROM РейсыПеревозчика WHERE НомерРейса=" & НомРес & ""
        Updates(strsql)

        Dim strsql1 As String = "DELETE * FROM РейсыКлиента WHERE НомерРейса=" & НомРес & ""
        Updates(strsql1)
        IO.File.Copy("Z:\RICKMANS\" & ComboBox1.Text & "\" & ПутьРейса, "Z:\RICKMANS\" & ComboBox1.Text & "\СОРВАННЫЕ ЗАГРУЗКИ\" & ПутьРейса)
        IO.File.Delete("Z:\RICKMANS\" & ComboBox1.Text & "\" & ПутьРейса)

        MessageBox.Show("Рейс полностью удалён!", Рик)
        Очистка()
        ПерегрЛист1()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Отмена = 0
        ПерезагрЛист1 = 0
        If Проверка() = 1 Then Exit Sub
        Me.Cursor = Cursors.WaitCursor
        If НомРес > 0 Then
            Dim Res As DialogResult = MessageBox.Show("Выберите 'Да'- если, хотите заменить данные этого рейса" & vbCrLf & " Выберите 'Нет'- если, хотите создать новый рейс" & vbCrLf & "Выберите 'Отмена'- если, хотите выйти", Рик, MessageBoxButtons.YesNoCancel)
            If Res = DialogResult.No Then
                НовыйРейсГлавная()
            End If
            If Res = DialogResult.Yes Then
                РедакцияСтарогоРейса()
                If Отмена = 1 Then
                    Me.Cursor = Cursors.Default
                    Exit Sub
                End If
                MessageBox.Show("Рейс изменен!", Рик)
                ПерегрЛист1()
            End If
            If Res = DialogResult.Cancel Then
                Me.Cursor = Cursors.Default
                Exit Sub
            End If
        End If
        If НомРес = Nothing Then
            ПерезагрЛист1 = 1
            НовыйРейсГлавная()
            Очистка()
            ПерезагрЛист1 = 0
        End If

        Me.Cursor = Cursors.Default
    End Sub
    Private Sub НовыйРейсГлавная()
        НовыйРейс()
        If Отмена = 1 Then Exit Sub
        Доки()
        MessageBox.Show("Рейс оформлен!", Рик)
        If ПерезагрЛист1 = 0 Then
            ПерегрЛист1()
        End If

    End Sub
End Class