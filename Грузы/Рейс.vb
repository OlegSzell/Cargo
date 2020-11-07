Option Explicit On
Imports System.Data.OleDb
Imports System.Threading
Imports System.IO

Public Class Рейс
    Dim file26(), file36() As String
    Dim ПутьПолный As String
    Dim b As New Thread(AddressOf ЗапускБыстро)
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
    Private Delegate Sub comb38()
    Private Delegate Sub comb3()
    Private Delegate Sub comb11()
    Private Delegate Sub comb1d()
    Private Delegate Sub comb2d()
    Private Delegate Sub comb3d()
    Private Delegate Sub comb4d()
    Private Delegate Sub comb5d()
    Private Delegate Sub comb5dd()
    Public ШтрафКлиент As Boolean = False
    Public ШтрафПер As Boolean = False
    Public ЧастичнаяОплатаПеревозчик, ЧастичнаяОплатаКлиент As String
    Public comdjx1 As String
    Private ОплатаПоКурсу As String = "False"
    Private com4all As List(Of IDNaz)
    Private bscom4 As BindingSource
    Private com3all As List(Of IDNaz)
    Private bscom3 As BindingSource
    Private lst1all As List(Of ПутиДоков)
    Private bslst11 As BindingSource


    Public Sub COM4()
        Dim f As New СпискиВсе
        Dim m As List(Of IDNaz) = f.Перевозчики
        If com4all IsNot Nothing Then
            com4all.Clear()
        End If
        com4all = New List(Of IDNaz)
        bscom4 = New BindingSource
        bscom4.DataSource = com4all
        ComboBox4.DataSource = bscom4
        ComboBox4.DisplayMember = "Naz"
        com4all.AddRange(m)
        bscom4.ResetBindings(False)
        ComboBox4.Text = String.Empty
    End Sub
    'Private Async Sub Awai()
    '    Await Task.Delay(50000)
    'End Sub

    Private Sub COM3()

        Dim f As New СпискиВсе
        Dim m As List(Of IDNaz) = f.Клиенты
        If com3all IsNot Nothing Then
            com3all.Clear()
        End If
        com3all = New List(Of IDNaz)
        bscom3 = New BindingSource
        bscom3.DataSource = com3all
        ComboBox3.DataSource = bscom3
        ComboBox3.DisplayMember = "Naz"
        com3all.AddRange(m)
        bscom3.ResetBindings(False)
        ComboBox3.Text = String.Empty







        ''Dim strsql1 As String
        ''Dim ds1 As DataTable
        'If ComboBox3.InvokeRequired Then
        '    Me.Invoke(New comb3(AddressOf COM3))
        'Else
        '    'strsql1 = "SELECT НазваниеОрганизации FROM Клиент ORDER BY НазваниеОрганизации"
        '    'ds1 = Selects3(strsql1)
        '    Using db As New dbAllDataContext()
        '        Dim var = (From x In db.Клиент
        '                   Order By x.НазваниеОрганизации
        '                   Select x.НазваниеОрганизации).ToList()


        '        If var.Count > 0 Then
        '            Me.ComboBox3.AutoCompleteCustomSource.Clear()
        '            Me.ComboBox3.Items.Clear()
        '            For Each r In var
        '                Me.ComboBox3.AutoCompleteCustomSource.Add(r)
        '                Me.ComboBox3.Items.Add(r)
        '            Next
        '        End If
        '    End Using
        'End If

    End Sub

    Private Sub COM11()

        If ComboBox11.InvokeRequired Then
            Me.Invoke(New comb11(AddressOf COM11))
        Else
            Using db As New dbAllDataContext()
                Dim var = db.ТипАвто.Select(Function(x) x).ToList()

                If var.Count > 0 Then
                    Me.ComboBox11.AutoCompleteCustomSource.Clear()
                    Me.ComboBox11.Items.Clear()
                    For Each r In var
                        Me.ComboBox11.AutoCompleteCustomSource.Add(r.ТипАвто)
                        Me.ComboBox11.Items.Add(r.ТипАвто)
                    Next
                End If
            End Using

        End If

    End Sub
    Private Sub COM12()

        Dim d() As String
        If ComboBox1.InvokeRequired Then
            Me.Invoke(New comb1d(AddressOf COM12))
        Else
            ComboBox1.Items.Clear()
            d = {Now.Year, Now.Year - 1, Now.Year - 2, Now.Year - 3, Now.Year - 4}
            ComboBox1.Items.AddRange(d)
            ComboBox1.Text = Now.Year
        End If

    End Sub
    Private Sub COM13()

        Dim d1() As String
        If ComboBox5.InvokeRequired Or ComboBox6.InvokeRequired Then
            Me.Invoke(New comb2d(AddressOf COM13))
        Else
            ComboBox5.Items.Clear()
            ComboBox6.Items.Clear()
            d1 = {"Рубль", "Евро", "Доллар", "Росс.рубль"}
            ComboBox5.Items.AddRange(d1)
            ComboBox6.Items.AddRange(d1)

        End If

    End Sub
    Private Sub COM14()

        Dim d1() As String
        If ComboBox7.InvokeRequired Or ComboBox8.InvokeRequired Then
            Me.Invoke(New comb3d(AddressOf COM14))
        Else
            ComboBox7.Items.Clear()
            ComboBox7.Enabled = False
            ComboBox8.Items.Clear()
            ComboBox8.Enabled = False
            d1 = {"$", "€", "RUS", "BYN"}
            ComboBox8.Items.AddRange(d1)
            ComboBox7.Items.AddRange(d1)

        End If

    End Sub
    Private Sub COM15()

        Dim d1() As String
        If ComboBox12.InvokeRequired Or ComboBox13.InvokeRequired Then
            Me.Invoke(New comb4d(AddressOf COM15))
        Else
            d1 = {"Поручение с экспедицией", "Договор-поручение", "Договор-поручение эксп"}
            ComboBox12.Items.AddRange(d1)
            ComboBox13.Items.AddRange(d1)

        End If

    End Sub
    Private Sub COM16()

        Dim d1() As String
        If ComboBox9.InvokeRequired Or ComboBox10.InvokeRequired Then
            Me.Invoke(New comb5d(AddressOf COM16))
        Else
            d1 = {"БанкПоДок", "КалПоДок", "БанкПоВыг", "КалПоВыг"}
            ComboBox10.Items.AddRange(d1)
            ComboBox9.Items.AddRange(d1)
            ПерегрЛист1()
        End If

    End Sub



    Private Sub ЗапускБыстро()
        'COM4()
        Dim y4 As New Thread(AddressOf COM4)
        y4.IsBackground = True
        y4.Start()

        Dim x3 As New Thread(AddressOf COM3)
        x3.IsBackground = True
        x3.Start()

        Dim x1 As New Thread(AddressOf COM11)
        x1.IsBackground = True
        x1.Start()

        Dim x4 As New Thread(AddressOf COM12)
        x4.IsBackground = True
        x4.Start()

        Dim x5 As New Thread(AddressOf COM13)
        x5.IsBackground = True
        x5.Start()

        Dim x6 As New Thread(AddressOf COM14)
        x6.IsBackground = True
        x6.Start()

        Dim x7 As New Thread(AddressOf COM15)
        x7.IsBackground = True
        x7.Start()

        Dim x8 As New Thread(AddressOf COM16)
        x8.IsBackground = True
        x8.Start()

        'Dim x9 As New Thread(AddressOf dtVyborka)
        'x9.IsBackground = True
        'x9.Start()

    End Sub
    Private Async Sub ПредзагрузкаAsync()
        Await Task.Run(Sub() Предзагрузка())
    End Sub
    Private Sub Предзагрузка()
        Dim mo As New AllUpd
        Do While AllClass.РейсыКлиента Is Nothing
            mo.РейсыКлиентаAll()
        Loop

        Do While AllClass.РейсыПеревозчика Is Nothing
            mo.РейсыПеревозчикаAll()
        Loop

    End Sub
    Private Sub Рейс_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ПредзагрузкаAsync()
        Me.Cursor = Cursors.WaitCursor
        lst1all = New List(Of ПутиДоков)
        bslst11 = New BindingSource
        bslst11.DataSource = lst1all
        ListBox1.DataSource = bslst11
        ListBox1.DisplayMember = "Путь"
        ListBox1.Text = String.Empty


        COM4()
        COM3()
        COM11()
        COM12()
        COM13()
        COM14()
        COM15()
        COM16()


        'b.IsBackground = True
        'b.Start()

        TextBox5.Text = Now.ToShortDateString
        MaskedTextBox3.Text = "09:00"
        TextBox6.Text = Now.ToShortDateString
        MaskedTextBox4.Text = "09:00"
        MaskedTextBox1.Text = Now.ToShortDateString
        comdjx1 = ComboBox1.Text




        Me.Cursor = Cursors.Default


    End Sub


    'Public Sub dtVyborka()

    '    'If ComboBox9.InvokeRequired Or ComboBox10.InvokeRequired Then
    '    '    Me.Invoke(New comb5dd(AddressOf dtVyborka))
    '    'Else
    '    dtZak = New DataTable()
    '    dtZak = Selects3(StrSql:="SELECT * FROM РейсыКлиента")
    '    dtPer = New DataTable()
    '    dtPer = Selects3(StrSql:="SELECT * FROM РейсыПеревозчика")
    '    'End If
    'End Sub

    Private Sub ИзменЭксел(ByVal strsql As String, ByVal t As Integer, ByVal n As Integer, ByVal g As Integer)



        Dim xlapp As Microsoft.Office.Interop.Excel.Application
        xlapp = New Microsoft.Office.Interop.Excel.Application

        'Dim XXX = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=U:\Офис\Рикманс\ДанныеРикманс.accdb; Persist Security Info=False;"
        'Dim XXX = "Provider='SQLOLEDB';Data Source=45.14.50.13\723\SQLEXPRESS,1433;Network Library=DBMSSOCN;Initial Catalog=Rickmans;User ID=userOleg;Password=Zf6VpP37Ol"
        Dim XXX = "Provider='SQLOLEDB';Data Source=45.14.50.142\2749\SQLEXPRESS,1433;Network Library=DBMSSOCN;Initial Catalog=Rickmans;User ID=userOleg1;Password=Zf6VpP37Ol"

        Dim CON As New ADODB.Connection
        Dim RS As New ADODB.Recordset

        CON.ConnectionString = XXX
        CON.Open()
        RS.Open(strsql, CON)

        If t = 1 Then
            xlapp.Workbooks.Open("Z:\RICKMANS\" & ComboBox1.Text & "\" & ПутьРейса).Worksheets("ЗАК").Range("L3").CopyFromRecordset(RS)
            xlapp.Workbooks(ПутьРейса).Close(True)
            IO.File.Copy("Z:\RICKMANS\" & ComboBox1.Text & "\" & ПутьРейса, "Z:\RICKMANS\" & ComboBox1.Text & "\" & НомРес & " " & ComboBox3.Text & " " & n & " - " & ComboBox4.Text & " " & g & ".xlsm", True)
            IO.File.Delete("Z:\RICKMANS\" & ComboBox1.Text & "\" & ПутьРейса)
            ПутьРейса = НомРес & " " & ComboBox3.Text & " " & n & " - " & ComboBox4.Text & " " & g & ".xlsm"
        Else
            xlapp.Workbooks.Open("Z:\RICKMANS\" & ComboBox1.Text & "\" & ПутьРейса).Worksheets("ЗАК").Range("L6").CopyFromRecordset(RS)
            xlapp.Workbooks(ПутьРейса).Close(True)
            IO.File.Copy("Z:\RICKMANS\" & ComboBox1.Text & "\" & ПутьРейса, "Z:\RICKMANS\" & ComboBox1.Text & "\" & НомРес & " " & ComboBox3.Text & " " & g & " - " & ComboBox4.Text & " " & n & ".xlsm", True)
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
        Dim ds As DataTable = Selects3(strsql)

        Dim strsql1 As String = "SELECT * FROM РейсыПеревозчика WHERE НомерРейса=" & НомРес & ""
        Dim ds1 As DataTable = Selects3(strsql1)

        Select Case d
            Case "Клиент"
                ИзменЭксел(strsql, 1, ds.Rows(0).Item(3), ds1.Rows(0).Item(3))
            Case "Перевоз"
                ИзменЭксел(strsql1, 0, ds1.Rows(0).Item(3), ds.Rows(0).Item(3))
        End Select
        MessageBox.Show("Данные удачно изменены в файле эксель!", Рик)
        ПерегрЛист1()
        ПерегрДанныхИзБазы()

        Me.Cursor = Cursors.Default

    End Sub
    Private Sub ПерегрДанныхИзБазы()
        dtVyborkaS()
        dtVyborkaS1()
        Перевозчики2()
        Клиенты2()
    End Sub

    Private Sub ПерегрЛист1()
        'ListBox1.Items.Clear()
        Справки(CType(ComboBox1.Text, Integer))
        If Files3.Count > 0 Then
            If lst1all IsNot Nothing Then
                lst1all.Clear()
            End If
            lst1all.AddRange(Files3)
            bslst11.ResetBindings(False)
            'ListBox1.Items.AddRange(Files3)
            ListBox1.Text = lst1all.Select(Function(x) x.Путь).LastOrDefault()
        End If

    End Sub
    Private Sub ListBox1_DoubleClick(sender As Object, e As EventArgs) Handles ListBox1.DoubleClick
        If ListBox1.SelectedIndex = -1 Then
            MessageBox.Show("Выберите документ для просмотра!", Рик, MessageBoxButtons.OK)
            Exit Sub
        End If
        Dim f As ПутиДоков = ListBox1.SelectedItem
        'Process.Start("Z:\RICKMANS\" & ComboBox1.Text & "\" & f.Путь)
        Process.Start(f.ПолныйПуть)
        'If Not ListBox1.SelectedIndex = -1 Then
        '    For i = 0 To ListBox1.SelectedItems.Count - 1
        '        Try
        '            Process.Start("Z:\RICKMANS\" & ComboBox1.Text & "\" & ListBox1.SelectedItems(i))
        '        Catch ex As Exception

        '        End Try

        '    Next

        'End If
    End Sub
    Private Async Function КонтЛицоТел(ByVal ds As DataTable) As Task(Of String)
        Dim strsql As String = "SELECT [Контактное лицо], Телефон FROM Перевозчики WHERE Названиеорганизации='" & ds.Rows(0).Item(1).ToString & "'"
        Dim ds1 As DataTable = Selects3(strsql)
        Await Task.Delay(0)
        Return Trim(ds1.Rows(0).Item(0).ToString & " " & ds1.Rows(0).Item(1).ToString)
    End Function

    Private Sub Вставка()
        Cursor = Cursors.WaitCursor

        Очистка()

        Dim mo As New AllUpd
        Do While AllClass.РейсыКлиента Is Nothing
            mo.РейсыКлиентаAll()
        Loop

        Do While AllClass.РейсыПеревозчика Is Nothing
            mo.РейсыПеревозчикаAll()
        Loop


        Dim rowzak As РейсыКлиента
        Dim rowper As РейсыПеревозчика

        rowzak = (From x In AllClass.РейсыКлиента
                  Where x.НомерРейса = НомРес
                  Select x).FirstOrDefault()


        rowper = (From x In AllClass.РейсыПеревозчика
                  Where x.НомерРейса = НомРес
                  Select x).FirstOrDefault()


        'Using db As New dbAllDataContext()
        '    rowzak = db.РейсыКлиента.Where(Function(x) x.НомерРейса = НомРес).Select(Function(x) x).FirstOrDefault()
        '    rowper = db.РейсыПеревозчика.Where(Function(x) x.НомерРейса = НомРес).Select(Function(x) x).FirstOrDefault()
        'End Using

        If rowzak Is Nothing Then
            Cursor = Cursors.Default
            Exit Sub
        End If
        If rowper Is Nothing Then
            Cursor = Cursors.Default
            Exit Sub
        End If



        ComboBox3.Text = rowzak.НазвОрганизации
        ComboBox4.Text = rowper.НазвОрганизации
        TextBox1.Text = rowzak.СтоимостьФрахта
        TextBox2.Text = rowper.СтоимостьФрахта
        ComboBox5.Text = rowzak.Валюта
        ComboBox6.Text = rowper.Валюта
        ComboBox8.Text = rowzak.ВалютаПлатежа
        ComboBox7.Text = rowper.ВалютаПлатежа
        TextBox4.Text = rowzak.СрокОплаты

        If rowzak.Предоплата <> "" Then 'клиент
            Button7.BackColor = Color.Red
            ЧастичнаяОплатаКлиент = rowzak.Предоплата
        End If
        If rowper.Предоплата <> "" Then 'перевозчик
            Button2.BackColor = Color.Red
            ЧастичнаяОплатаПеревозчик = rowper.Предоплата
        End If

        TextBox3.Text = rowper.СрокОплаты
        MaskedTextBox1.Text = rowzak.ДатаПоручения

        If rowzak.УсловияОплаты = "1" Then
            ComboBox10.Text = "БанкПоДок"
        ElseIf rowzak.УсловияОплаты = "3" Then
            ComboBox10.Text = "БанкПоВыг"
        Else
            ComboBox10.Text = rowzak.УсловияОплаты
        End If

        If rowper.УсловияОплаты = "1" Then
            ComboBox9.Text = "БанкПоДок"
        ElseIf rowper.УсловияОплаты = "3" Then
            ComboBox9.Text = "БанкПоВыг"
        Else
            ComboBox9.Text = rowper.УсловияОплаты
        End If

        If rowzak.ДогПор = "1" Then
            ComboBox12.Text = "Договор-поручение"
        ElseIf rowzak.ДогПорЭксп = "1" Then
            ComboBox12.Text = "Договор-поручение эксп"
        ElseIf rowzak.ПорЭксп = "1" Then
            ComboBox12.Text = "Поручение с экспедицией"
        Else
            ComboBox12.Text = ""
        End If

        If rowper.ДогПор = "1" Then
            ComboBox13.Text = "Договор-поручение"
        ElseIf rowper.ДогПорЭксп = "1" Then
            ComboBox13.Text = "Договор-поручение эксп"
        ElseIf rowper.ПорЭксп = "1" Then
            ComboBox13.Text = "Поручение с экспедицией"
        Else
            ComboBox13.Text = ""
        End If

        RichTextBox1.Text = rowzak.ДопУсловия
        RichTextBox2.Text = rowper.ДопУсловия

        RichTextBox10.Text = rowzak.Маршрут
        TextBox5.Text = rowzak.ДатаПодачиПодЗагрузку
        MaskedTextBox3.Text = rowzak.ВремяПодачи
        TextBox6.Text = rowzak.ДатаПодачиПодРастаможку
        MaskedTextBox4.Text = rowzak.ВремяПодачиВыгРаст

        RichTextBox3.Text = rowper.ТочныйАдресЗагрузки
        RichTextBox4.Text = rowper.АдресЗатаможки
        RichTextBox7.Text = rowper.НаименованиеГруза
        ComboBox11.Text = rowzak.ТипТрСредства
        RichTextBox8.Text = rowper.НомерАвтомобиля
        RichTextBox9.Text = rowper.Водитель
        RichTextBox6.Text = rowper.ТочнАдресРазгр
        RichTextBox5.Text = rowper.ТочнАдресРаста



        Cursor = Cursors.Default


    End Sub
    Private Sub ClkLst()
        If Not ListBox1.SelectedIndex = -1 Then
            Dim f As ПутиДоков = ListBox1.SelectedItem
            Try
                НомРес = CType(Strings.Left(f.Путь, 3), Integer)
                ПутьРейса = f.Путь
                ПутьПолный = f.ПолныйПуть



            Catch ex As Exception

            End Try

        End If
        Вставка()
    End Sub
    Private Sub ListBox1_Click(sender As Object, e As EventArgs) Handles ListBox1.Click
        ClkLst()
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
        Select Case ComboBox5.Text
            Case "Евро"
                ComboBox8.Enabled = True
            Case "Доллар"
                ComboBox8.Enabled = True
            Case "Росс.рубль"
                ComboBox8.Enabled = True
            Case "Рубль"
                ComboBox8.Enabled = False
        End Select

    End Sub

    Private Sub ComboBox6_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox6.SelectedIndexChanged
        Select Case ComboBox6.Text
            Case "Евро"
                ComboBox7.Enabled = True
            Case "Доллар"
                ComboBox7.Enabled = True
            Case "Росс.рубль"
                ComboBox7.Enabled = True
            Case "Рубль"
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

    Private Sub Button2_Click(sender As Object, e As EventArgs)
        'If ComboBox3.Text = "" Or TextBox1.Text = "" Then
        '    MessageBox.Show("Выберите рейс!", Рик)
        '    Exit Sub
        'End If
        'ДопФорма.ShowDialog()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs)
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
        Dim mo As New AllUpd
        Do While AllClass.РейсыКлиента Is Nothing
            mo.РейсыКлиентаAll()
        Loop
        Do While AllClass.РейсыКлиента Is Nothing
            mo.РейсыКлиентаAll()
        Loop

        'Dim strsql, strsql1 As String
        ''Dim ds As New DataTable


        Dim f As IDNaz = ComboBox3.SelectedItem
        Dim ds5 As Клиент
        Dim com3 As String
        com3 = f.Naz
        Using db As New dbAllDataContext()
            ds5 = db.Клиент.Where(Function(x) x.НазваниеОрганизации = com3).Select(Function(x) x).FirstOrDefault()
        End Using


        Dim f2 As Integer
        If com3 = "Виталюр" Then
            f2 = AllClass.РейсыКлиента.OrderBy(Function(x) x.НазвОрганизации).Where(Function(x) x.НазвОрганизации = "Виталюр" And Format(x.Год, "yyyy") = Now.Year).Select(Function(x) x.КоличРейсов).LastOrDefault()

            'strsql = "SELECT MAX(КоличРейсов) FROM РейсыКлиента WHERE НазвОрганизации = 'Виталюр'  AND ДатаПодачиПодЗагрузку Like '%" & Now.Year & "%' GROUP BY НазвОрганизации "
            'ds = Selects3(strsql)
        Else
            f2 = AllClass.РейсыКлиента.OrderBy(Function(x) x.НазвОрганизации).Where(Function(x) x.НазвОрганизации = com3).Select(Function(x) x.КоличРейсов).LastOrDefault()
            'strsql1 = "SELECT MAX(КоличРейсов) FROM [РейсыКлиента] WHERE НазвОрганизации = '" & ComboBox3.Text & "' GROUP BY НазвОрганизации "
            'ds = Selects3(strsql1)
        End If


        Try
            СлПорРейсКл = f2 + 1
        Catch ex As Exception

            Nm = ""
            Nm = com3
            Dim f4 As New ПорНомер
            If ds5.Договор = "" Or ds5.Дата = "" Then
                f4.GroupBox1.Enabled = True
            End If
            If ds5.Должность = "" Or ds5.ФИОРуководителя = "" Then
                f4.GroupBox2.Enabled = True
            End If
            f4.TextBox1.Enabled = True
            f4.ListBox1.Enabled = True
            pro = 1
            f4.ShowDialog()
            If Отмена = 1 Then Exit Sub
            'СлПорРейсКл = CType((ПорНомер.TextBox1.Text), Integer) + 1
            Пров = 1
        End Try

        If Пров = 0 Then
            Dim f1, g As Integer
            Dim f5 As New ПорНомер
            If ds5.Договор = "" Or ds5.Дата = "" Then
                f5.GroupBox1.Enabled = True
                f1 = 1
            End If
            If ds5.Должность = "" Or ds5.ФИОРуководителя = "" Then
                f5.GroupBox2.Enabled = True
                g = 1
            End If
            If f1 = 1 Or g = 1 Then
                pro = 1
                f5.ShowDialog()
                If Отмена = 1 Then Exit Sub
            End If
        End If
    End Sub
    Private Sub ПровСледРейсПер()

        Dim mo As New AllUpd
        Do While AllClass.Перевозчики Is Nothing
            mo.ПеревозчикиAll()
        Loop

        Do While AllClass.РейсыПеревозчика Is Nothing
            mo.РейсыПеревозчикаAll()
        Loop
        Dim f As IDNaz = ComboBox4.SelectedItem
        Dim ds6 = AllClass.Перевозчики.Where(Function(x) x.Названиеорганизации = f.Naz).Select(Function(x) x).FirstOrDefault()
        Dim ds2 = AllClass.РейсыПеревозчика.OrderBy(Function(x) x.КоличРейсов).Where(Function(x) x.НазвОрганизации = f.Naz).Select(Function(x) x.КоличРейсов).LastOrDefault()

        'Dim strsql6 As String = "SELECT Договор,Дата,Должность,ФИОРуководителя FROM Перевозчики WHERE Названиеорганизации='" & ComboBox4.Text & "'"
        'Dim ds6 As DataTable = Selects3(strsql6)


        'Dim strsql2 As String = "SELECT MAX(КоличРейсов) FROM РейсыПеревозчика WHERE НазвОрганизации = '" & ComboBox4.Text & "' GROUP BY НазвОрганизации"
        'Dim ds2 As DataTable = Selects3(strsql2)
        If ds2 Is Nothing Then
            ds2 = 0
        End If
        Try
            СлПорРейсПер = ds2 + 1
        Catch ex As Exception
            Nm = ""
            Nm = ComboBox4.Text
            pro = 0

            If ds6.Договор = "" Or ds6.Дата = "" Then
                ПорНомер.GroupBox1.Enabled = True
            End If
            If ds6.Должность = "" Or ds6.ФИОРуководителя = "" Then
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
            Dim f2, g As Integer
            Dim f3 As New ПорНомер
            If ds6.Договор = "" Or ds6.Дата = "" Then
                f3.GroupBox1.Enabled = True
                f2 = 1
            End If
            If ds6.Должность = "" Or ds6.ФИОРуководителя = "" Then
                f3.GroupBox2.Enabled = True
                g = 1
            End If

            If f2 = 1 Or g = 1 Then
                pro = 0
                f3.ShowDialog()
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
        'Dim ds1 As DataTable


        ПровСледРейсКлиент()
        If Отмена = 1 Then Exit Sub

        Пров = 0
        Dim mo As New AllUpd
        Do While AllClass.РейсыКлиента Is Nothing
            mo.РейсыКлиентаAll()
        Loop
        Dim ds1 As Integer = AllClass.РейсыКлиента.OrderBy(Function(x) x.НомерРейса).Select(Function(x) x.НомерРейса).LastOrDefault()
        'strsql = "SELECT MAX(НомерРейса) FROM РейсыКлиента"
        '    ds1 = Selects3(strsql)

        СлРейс = ds1 + 1
        Dim cm4 As IDNaz = ComboBox4.SelectedItem
        ПровСледРейсПер()
        If Отмена = 1 Then Return

        ComB13()
        Dim f1 As New РейсыПеревозчика
        With f1
            .НазвОрганизации = cm4.Naz
            .НомерРейса = СлРейс
            .КоличРейсов = СлПорРейсПер
            .Маршрут = Trim(RichTextBox10.Text)
            .ДатаПодачиПодЗагрузку = TextBox5.Text
            .ВремяПодачи = MaskedTextBox3.Text
            .ДатаПодачиПодРастаможку = TextBox6.Text
            .ВремяПодачиВыгРаст = MaskedTextBox4.Text
            .ТочныйАдресЗагрузки = Trim(RichTextBox3.Text)
            .АдресЗатаможки = Trim(RichTextBox4.Text)
            .НаименованиеГруза = Trim(RichTextBox7.Text)
            .ТипТрСредства = ComboBox11.Text
            .НомерАвтомобиля = Trim(RichTextBox8.Text)
            .Водитель = Trim(RichTextBox9.Text)
            .ТочнАдресРаста = Trim(RichTextBox5.Text)
            .ТочнАдресРазгр = Trim(RichTextBox6.Text)
            .СтоимостьФрахта = TextBox2.Text
            .Валюта = ComboBox6.Text
            .ВалютаПлатежа = ComboBox7.Text
            .СрокОплаты = TextBox3.Text
            .ДопУсловия = Trim(RichTextBox2.Text)
            .ДогПор = ДогПор
            .ДогПорЭксп = ДогПорЭксп
            .ДатаПоручения = MaskedTextBox1.Text
            .ПорЭксп = ПорЭксп
            .УсловияОплаты = ComboBox9.Text
            .РазмерШтрафаЗаСрыв = ШтрафКлиент
            .Предоплата = ЧастичнаяОплатаПеревозчик
            .СрывЗагр20Проц = Procenty20
            .ДатаСоздания = Now
            .Экспедитор = Экспедитор
        End With
        Using db As New dbAllDataContext()
            db.РейсыПеревозчика.InsertOnSubmit(f1)
            db.SubmitChanges()
        End Using
        '        strsql3 = "INSERT INTO РейсыПеревозчика(НазвОрганизации,НомерРейса,КоличРейсов,Маршрут,ДатаПодачиПодЗагрузку,ВремяПодачи,ДатаПодачиПодРастаможку,
        'ВремяПодачиВыгРаст,ТочныйАдресЗагрузки,АдресЗатаможки,НаименованиеГруза,ТипТрСредства,НомерАвтомобиля,Водитель,
        'ТочнАдресРаста,ТочнАдресРазгр,СтоимостьФрахта,Валюта,ВалютаПлатежа,СрокОплаты,ДопУсловия,
        'ДогПор,ДогПорЭксп,ДатаПоручения,ПорЭксп,УсловияОплаты,РазмерШтрафаЗаСрыв,Предоплата,СрывЗагр20Проц)
        'VALUES('" & ComboBox4.Text & "'," & СлРейс & "," & СлПорРейсПер & ",'" & Trim(RichTextBox10.Text) & "','" & TextBox5.Text & "','" & MaskedTextBox3.Text & "','" & TextBox6.Text & "',
        ''" & MaskedTextBox4.Text & "','" & Trim(RichTextBox3.Text) & "','" & Trim(RichTextBox4.Text) & "','" & Trim(RichTextBox7.Text) & "','" & ComboBox11.Text & "','" & Trim(RichTextBox8.Text) & "','" & Trim(RichTextBox9.Text) & "',
        ''" & Trim(RichTextBox5.Text) & "','" & Trim(RichTextBox6.Text) & "','" & TextBox2.Text & "','" & ComboBox6.Text & "','" & ComboBox7.Text & "','" & TextBox3.Text & "','" & Trim(RichTextBox2.Text) & "',
        ''" & ДогПор & "','" & ДогПорЭксп & "','" & MaskedTextBox1.Text & "','" & ПорЭксп & "','" & ComboBox9.Text & "','" & ШтрафКлиент & "','" & ЧастичнаяОплатаПеревозчик & "', '" & Procenty20 & "')"
        '        Updates3(strsql3)

        ComB12()
        ComB5()

        Using db As New dbAllDataContext()
            Dim var As New РейсыКлиента
            var.НазвОрганизации = ComboBox3.Text
            var.НомерРейса = СлРейс
            var.КоличРейсов = СлПорРейсКл
            var.Маршрут = Trim(RichTextBox10.Text)
            var.ДатаПодачиПодЗагрузку = TextBox5.Text
            var.ВремяПодачи = MaskedTextBox3.Text
            var.ДатаПодачиПодРастаможку = TextBox6.Text
            var.ВремяПодачиВыгРаст = MaskedTextBox4.Text
            var.ТочныйАдресЗагрузки = Trim(RichTextBox3.Text)
            var.АдресЗатаможки = Trim(RichTextBox4.Text)
            var.НаименованиеГруза = Trim(RichTextBox7.Text)
            var.ТипТрСредства = ComboBox11.Text
            var.НомерАвтомобиля = Trim(RichTextBox8.Text)
            var.Водитель = Trim(RichTextBox9.Text)
            var.ТочнАдресРаста = Trim(RichTextBox5.Text)
            var.ТочнАдресРазгр = Trim(RichTextBox6.Text)
            var.СтоимостьФрахта = TextBox1.Text
            var.Валюта = ComboBox5.Text
            var.ВалютаПлатежа = ComboBox8.Text
            var.СрокОплаты = TextBox4.Text
            var.ДопУсловия = Trim(RichTextBox1.Text)
            var.ДогПор = ДогПор
            var.ДогПорЭксп = ДогПорЭксп
            var.ДатаПоручения = MaskedTextBox1.Text
            var.ПорЭксп = ПорЭксп
            var.УсловияОплаты = ComboBox10.Text
            var.Год = Now.ToShortDateString
            var.РазмерШтрафаЗаСрыв = ШтрафПер
            var.Предоплата = ЧастичнаяОплатаКлиент
            var.ОплатаПоКурсу = ОплатаПоКурсу
            var.Экспедитор = Экспедитор
            var.ДатаСоздания = Now

            db.РейсыКлиента.InsertOnSubmit(var)
            db.SubmitChanges()

        End Using




    End Sub
    Private Sub ComB5()
        If Not ComboBox5.Text = "Рубль" And ComboBox8.Text = "BYN" Then
            ОплатаПоКурсу = "True"
        Else
            ОплатаПоКурсу = "False"
        End If
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

        'Dim XXX = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=U:\Офис\Рикманс\ДанныеРикманс.accdb; Persist Security Info=False;"
        'Dim XXX = "Provider='SQLOLEDB';Data Source=45.14.50.13\723\SQLEXPRESS,1433;Network Library=DBMSSOCN;Initial Catalog=Rickmans;User ID=userOleg;Password=Zf6VpP37Ol"
        ' Dim XXX = "Provider='SQLOLEDB';Data Source=45.14.50.142\2749\SQLEXPRESS,1433;Network Library=DBMSSOCN;Initial Catalog=Rickmans;User ID=userOleg1;Password=Zf6VpP37Ol"
        Dim XXX = "Provider='SQLOLEDB';Data Source=178.124.211.175,52891;Network Library=DBMSSOCN;Initial Catalog=Rickmans;Persist Security Info=True;User ID=Rickmans;Password=Zf6VpP37Ol"

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
        Dim f As New ИзменПорНомерКлиента2(НомРес, ComboBox3.Text, ComboBox4.Text)
        f.ShowDialog()
        If f.Rez IsNot Nothing Then
            ИзменВДействРейсе(f.Rez)
        End If

    End Sub

    Private Sub РедакцияСтарогоРейса()
        Dim strsql, strsql1, strsql2, strsql3 As String



        'Dim ds As DataRow() = РейсыКлиента(НомРес)
        Dim ds As String
        Using db As New dbAllDataContext()
            ds = db.РейсыКлиента.Where(Function(x) x.НомерРейса = НомРес).Select(Function(x) x.НазвОрганизации).FirstOrDefault()
            If ds Is Nothing Then Exit Sub
        End Using




        If Not ds = ComboBox3.Text Then
            ПровСледРейсКлиент()
            If Отмена = 1 Then Exit Sub
            ComB5()
            ComB12()
            strsql1 = "UPDATE РейсыКлиента SET НазвОрганизации='" & ComboBox3.Text & "', КоличРейсов=" & СлПорРейсКл & ", Маршрут='" & Trim(RichTextBox10.Text) & "',
ДатаПодачиПодЗагрузку='" & TextBox5.Text & "', ВремяПодачи='" & MaskedTextBox3.Text & "', ДатаПодачиПодРастаможку='" & TextBox6.Text & "', ВремяПодачиВыгРаст='" & MaskedTextBox4.Text & "',
ТочныйАдресЗагрузки='" & Trim(RichTextBox3.Text) & "', АдресЗатаможки='" & Trim(RichTextBox4.Text) & "', НаименованиеГруза='" & Trim(RichTextBox7.Text) & "', ТипТрСредства='" & ComboBox11.Text & "',
НомерАвтомобиля='" & Trim(RichTextBox8.Text) & "',Водитель ='" & Trim(RichTextBox9.Text) & "', ТочнАдресРаста='" & Trim(RichTextBox5.Text) & "', ТочнАдресРазгр='" & Trim(RichTextBox6.Text) & "',
СтоимостьФрахта='" & TextBox1.Text & "', Валюта='" & ComboBox5.Text & "', ВалютаПлатежа='" & ComboBox8.Text & "', СрокОплаты='" & TextBox4.Text & "', УсловияОплаты='" & ComboBox10.Text & "',
ДогПор='" & ДогПор & "',ДогПорЭксп='" & ДогПорЭксп & "',ПорЭксп= '" & ПорЭксп & "', ДопУсловия='" & Trim(RichTextBox1.Text) & "', ДатаПоручения='" & MaskedTextBox1.Text & "',
РазмерШтрафаЗаСрыв='" & ШтрафКлиент & "',Предоплата='" & ЧастичнаяОплатаКлиент & "', ОплатаПоКурсу='" & ОплатаПоКурсу & "'
WHERE НомерРейса=" & НомРес & ""
            Updates3(strsql1)
            ПрИзмНазКл = True
        Else
            ComB5()
            ComB12()
            strsql1 = "UPDATE РейсыКлиента SET Маршрут='" & Trim(RichTextBox10.Text) & "',
ДатаПодачиПодЗагрузку='" & TextBox5.Text & "', ВремяПодачи='" & MaskedTextBox3.Text & "', ДатаПодачиПодРастаможку='" & TextBox6.Text & "', ВремяПодачиВыгРаст='" & MaskedTextBox4.Text & "',
ТочныйАдресЗагрузки='" & Trim(RichTextBox3.Text) & "', АдресЗатаможки='" & Trim(RichTextBox4.Text) & "', НаименованиеГруза='" & Trim(RichTextBox7.Text) & "', ТипТрСредства='" & ComboBox11.Text & "',
НомерАвтомобиля='" & Trim(RichTextBox8.Text) & "',Водитель ='" & Trim(RichTextBox9.Text) & "', ТочнАдресРаста='" & Trim(RichTextBox5.Text) & "', ТочнАдресРазгр='" & Trim(RichTextBox6.Text) & "',
СтоимостьФрахта='" & TextBox1.Text & "', Валюта='" & ComboBox5.Text & "', ВалютаПлатежа='" & ComboBox8.Text & "', СрокОплаты='" & TextBox4.Text & "', УсловияОплаты='" & ComboBox10.Text & "',
ДогПор='" & ДогПор & "', ДогПорЭксп='" & ДогПорЭксп & "',ПорЭксп= '" & ПорЭксп & "', ДопУсловия='" & Trim(RichTextBox1.Text) & "', ДатаПоручения='" & MaskedTextBox1.Text & "',
РазмерШтрафаЗаСрыв='" & ШтрафКлиент & "',Предоплата='" & ЧастичнаяОплатаКлиент & "', ОплатаПоКурсу='" & ОплатаПоКурсу & "'
WHERE НомерРейса=" & НомРес & ""
            Updates3(strsql1)
        End If

        'strsql2 = "SELECT НазвОрганизации FROM РейсыПеревозчика WHERE НомерРейса=" & НомРес & ""
        'Dim ds1 As DataTable = Selects3(strsql2)

        'Dim ds1 As DataRow() = РейсыПеревозчик(НомРес)
        Dim ds1 As РейсыПеревозчика
        Using db As New dbAllDataContext()
            ds1 = (From x In db.РейсыПеревозчика
                   Where x.НомерРейса = НомРес
                   Select x).FirstOrDefault()

            If ds1 IsNot Nothing Then
                If ds1.СрывЗагр20Проц = "True" Then
                    Procenty20 = "True"
                End If
            Else
                Exit Sub
            End If
        End Using


        If Not ds1.НазвОрганизации = ComboBox4.Text Then
            ПровСледРейсПер()
            If Отмена = 1 Then Exit Sub
            ComB13()
            strsql3 = "UPDATE РейсыПеревозчика SET НазвОрганизации='" & ComboBox4.Text & "', КоличРейсов=" & СлПорРейсПер & ", Маршрут='" & Trim(RichTextBox10.Text) & "',
ДатаПодачиПодЗагрузку='" & TextBox5.Text & "', ВремяПодачи='" & MaskedTextBox3.Text & "', ДатаПодачиПодРастаможку='" & TextBox6.Text & "', ВремяПодачиВыгРаст='" & MaskedTextBox4.Text & "',
ТочныйАдресЗагрузки='" & Trim(RichTextBox3.Text) & "', АдресЗатаможки='" & Trim(RichTextBox4.Text) & "', НаименованиеГруза='" & Trim(RichTextBox7.Text) & "', ТипТрСредства='" & ComboBox11.Text & "',
НомерАвтомобиля='" & Trim(RichTextBox8.Text) & "',Водитель ='" & Trim(RichTextBox9.Text) & "', ТочнАдресРаста='" & Trim(RichTextBox5.Text) & "', ТочнАдресРазгр='" & Trim(RichTextBox6.Text) & "',
СтоимостьФрахта='" & TextBox2.Text & "', Валюта='" & ComboBox6.Text & "', ВалютаПлатежа='" & ComboBox7.Text & "', СрокОплаты='" & TextBox3.Text & "', УсловияОплаты='" & ComboBox9.Text & "',
ДогПор='" & ДогПор & "', ДогПорЭксп='" & ДогПорЭксп & "',ПорЭксп= '" & ПорЭксп & "', ДопУсловия='" & Trim(RichTextBox2.Text) & "',
ДатаПоручения='" & MaskedTextBox1.Text & "',РазмерШтрафаЗаСрыв='" & ШтрафПер & "',Предоплата='" & ЧастичнаяОплатаПеревозчик & "',СрывЗагр20Проц='" & Procenty20 & "'
WHERE НомерРейса=" & НомРес & ""
            Updates3(strsql3)
            ПрИзмНазПер = True
        Else
            ComB13()
            strsql3 = "UPDATE РейсыПеревозчика SET Маршрут='" & Trim(RichTextBox10.Text) & "',
ДатаПодачиПодЗагрузку='" & TextBox5.Text & "', ВремяПодачи='" & MaskedTextBox3.Text & "', ДатаПодачиПодРастаможку='" & TextBox6.Text & "', ВремяПодачиВыгРаст='" & MaskedTextBox4.Text & "',
ТочныйАдресЗагрузки='" & Trim(RichTextBox3.Text) & "', АдресЗатаможки='" & Trim(RichTextBox4.Text) & "', НаименованиеГруза='" & Trim(RichTextBox7.Text) & "', ТипТрСредства='" & ComboBox11.Text & "',
НомерАвтомобиля='" & Trim(RichTextBox8.Text) & "',Водитель ='" & Trim(RichTextBox9.Text) & "', ТочнАдресРаста='" & Trim(RichTextBox5.Text) & "', ТочнАдресРазгр='" & Trim(RichTextBox6.Text) & "',
СтоимостьФрахта='" & TextBox2.Text & "', Валюта='" & ComboBox6.Text & "', ВалютаПлатежа='" & ComboBox7.Text & "', СрокОплаты='" & TextBox3.Text & "', УсловияОплаты='" & ComboBox9.Text & "',
ДогПор='" & ДогПор & "', ДогПорЭксп='" & ДогПорЭксп & "',ПорЭксп= '" & ПорЭксп & "', ДопУсловия='" & Trim(RichTextBox2.Text) & "',
ДатаПоручения='" & MaskedTextBox1.Text & "',РазмерШтрафаЗаСрыв='" & ШтрафПер & "',Предоплата='" & ЧастичнаяОплатаПеревозчик & "',СрывЗагр20Проц='" & Procenty20 & "'
WHERE НомерРейса=" & НомРес & ""
            Updates3(strsql3)
        End If
        СлРейс = НомРес
        ДокиОбновление()

    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs)
        'If НомРес = Nothing Then
        '    MessageBox.Show("Выберите рейс!", Рик)
        '    Exit Sub
        'End If
        'bl = False
        'Отчет.ShowDialog()



    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs)
        'Сводная_по_рейсам.Show()

    End Sub

    Private Sub ListBox1_MouseDown(sender As Object, e As MouseEventArgs) Handles ListBox1.MouseDown
        If e.Button = MouseButtons.Right Then
            ContextMenuStrip1.Show(MousePosition, ToolStripDropDownDirection.Right)
        End If
    End Sub
    Private Sub Запускэксель(ByVal d As String)
        Dim xlapp1 As Microsoft.Office.Interop.Excel.Application
        Dim xlworkbook1 As Microsoft.Office.Interop.Excel.Workbook
        Dim xlworksheet1 As Microsoft.Office.Interop.Excel.Worksheet
        'Dim misvalue As Object = Reflection.Missing.Value
        xlapp1 = New Microsoft.Office.Interop.Excel.Application With {
            .Visible = False
        }
        'xlworkbook = xlapp.Workbooks.Add(misvalue)
        xlworkbook1 = xlapp1.Workbooks.Open(ПутьПолный,, True)

        xlworksheet1 = xlworkbook1.Sheets(d)
        xlworksheet1.PrintOutEx(,, 1)
        xlworkbook1.Close(False)
        xlapp1.Quit()

        releaseobject(xlapp1)
        releaseobject(xlworkbook1)
        releaseobject(xlworksheet1)
    End Sub

    Private Sub ПоручениеПеревозчикпечататьToolStripMenuItem_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub ОбаПечататьToolStripMenuItem_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub ЛистокпечататьToolStripMenuItem_Click(sender As Object, e As EventArgs)


    End Sub

    Private Sub ОбаЛистокпечататьToolStripMenuItem_Click(sender As Object, e As EventArgs)

        Запускэксель("ЗАК")
        Запускэксель("ПЕР")
        Запускэксель("РЕЙС")
    End Sub

    Private Sub ПоручениеКлиентпечататьToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles ПоручениеКлиентпечататьToolStripMenuItem1.Click
        Dim d As New Thread(Sub() Запускэксель("ЗАК"))
        d.IsBackground = True
        d.Start()
    End Sub

    Private Sub ПоручениеПеревозчикпечататьToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles ПоручениеПеревозчикпечататьToolStripMenuItem1.Click
        Dim d As New Thread(Sub() Запускэксель("ПЕР"))
        d.IsBackground = True
        d.Start()

    End Sub

    Private Sub ОбапечататьToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles ОбапечататьToolStripMenuItem1.Click
        Dim d As New Thread(Sub() Запускэксель("ПЕР"))
        d.IsBackground = True
        d.Start()
        Dim d1 As New Thread(Sub() Запускэксель("ЗАК"))
        d1.IsBackground = True
        d1.Start()
        'Запускэксель("ЗАК")
        'Запускэксель("ПЕР")
    End Sub

    Private Sub ЛистокпечататьToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles ЛистокпечататьToolStripMenuItem1.Click
        Dim d As New Thread(Sub() Запускэксель("РЕЙС"))
        d.IsBackground = True
        d.Start()
    End Sub

    Private Sub ОбаЛистокпечататьToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles ОбаЛистокпечататьToolStripMenuItem1.Click
        Dim d As New Thread(Sub() Запускэксель("ПЕР"))
        d.IsBackground = True
        d.Start()
        Dim d1 As New Thread(Sub() Запускэксель("ЗАК"))
        d1.IsBackground = True
        d1.Start()
        Dim d2 As New Thread(Sub() Запускэксель("РЕЙС"))
        d2.IsBackground = True
        d2.Start()
    End Sub

    Private Sub ОткрытьToolStripMenuItem_Click(sender As Object, e As EventArgs)
        numberRe(1)
        'For i As Integer = 0 To file36.Count - 1
        '    Process.Start(file36(i))
        'Next

    End Sub
    Private Sub numberRe(ByVal tu As Integer)
        Dim f() As String = {"РейсыПеревозчика", "РейсыКлиента"}
        Dim dr As New Thread(AddressOf Сбор)
        dr.IsBackground = True
        Dim dict As New Dictionary(Of String, String)



        Try
            Dim ds As DataTable = Selects3(StrSql:="SELECT НазвОрганизации FROM " & f(tu) & " WHERE НомерРейса=" & НомРес & "")
            dict.Add(f(tu), ds.Rows(0).Item(0).ToString)
            dr.Start(dict)
            'Return ds.Rows(0).Item(0).ToString
        Catch ex As Exception
            'Return 0
        End Try

    End Sub

    Private Sub СводнаяToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles СводнаяToolStripMenuItem.Click
        Dim f As New Сводная_по_рейсам
        f.Show()

    End Sub

    Private Sub ОплатаToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ОплатаToolStripMenuItem.Click
        If НомРес = Nothing Then
            MessageBox.Show("Выберите рейс!", Рик)
            Exit Sub
        End If
        bl = False
        Dim f As New Отчет(НомРес)
        f.ShowDialog()
    End Sub

    Private Sub ШтрафыToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ШтрафыToolStripMenuItem.Click
        Штрафы.ShowDialog()
    End Sub

    Private Sub АктСчетИРазбивкаToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles АктСчетИРазбивкаToolStripMenuItem.Click
        If ComboBox3.Text = "" Or TextBox1.Text = "" Then
            MessageBox.Show("Выберите рейс!", Рик)
            Exit Sub
        End If
        Dim f As New ДопФорма(НомРес, TextBox1.Text)
        f.ShowDialog()
        If f.ОбнвлExcel = True Then
            ДокиОбновление(f.Num)
        End If
    End Sub

    Private Sub ПеревозчикToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles ПеревозчикToolStripMenuItem1.Click
        НовыйКлиент.ShowDialog()
    End Sub

    Private Sub ПеревозчикToolStripMenuItem2_Click(sender As Object, e As EventArgs) Handles ПеревозчикToolStripMenuItem2.Click
        НовыйПеревоз.ShowDialog()
    End Sub

    Private Sub ОткрытьToolStripMenuItem1_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub КлиентToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles КлиентToolStripMenuItem.Click
        numberRe(1)
    End Sub

    Private Sub ПеревозчикToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ПеревозчикToolStripMenuItem.Click
        numberRe(0)
    End Sub

    Private Sub Button7_Click_1(sender As Object, e As EventArgs) Handles Button7.Click
        ПредоплатаКлиент.ShowDialog()
    End Sub

    Private Sub ToolStripMenuItem2_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem2.Click
        Dim d As New Thread(Sub() Запускэксель("СФ"))
        d.IsBackground = True
        d.Start()
    End Sub

    Private Sub ToolStripMenuItem3_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem3.Click

        Dim d As New Thread(Sub() Запускэксель("СФ"))
        d.IsBackground = True
        d.Start()

        Dim d1 As New Thread(Sub() Запускэксель("СФ"))
        d1.IsBackground = True
        d1.Start()


    End Sub

    Private Sub ToolStripMenuItem4_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem4.Click
        Dim d As New Thread(Sub() Запускэксель("АКТ"))
        d.IsBackground = True
        d.Start()
    End Sub

    Private Sub ToolStripMenuItem5_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem5.Click
        Dim d As New Thread(Sub() Запускэксель("АКТ"))
        d.IsBackground = True
        d.Start()

        Dim d1 As New Thread(Sub() Запускэксель("АКТ"))
        d1.IsBackground = True
        d1.Start()
    End Sub

    Private Sub Button8_Click_1(sender As Object, e As EventArgs) Handles Button8.Click
        comdjx1 = ComboBox1.Text
        Dim f As New ПоискВРейсах(_год:=ComboBox1.Text)
        f.ShowDialog()
        If f.NumbCor IsNot Nothing Then
            If f.NumbCor.Length > 0 Then
                Dim ind = ListBox1.FindString(f.NumbCor)
                ListBox1.SetSelected(ind, True)
                ClkLst()
            End If
        End If

    End Sub

    Private Sub СчетактToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles СчетактToolStripMenuItem.Click
        If ComboBox3.Text = "" Or TextBox1.Text = "" Then
            MessageBox.Show("Выберите рейс!", Рик)
            Exit Sub
        End If
        Dim f As New ДопФорма(НомРес, TextBox1.Text)
        f.ShowDialog()
        If f.ОбнвлExcel = True Then
            ДокиОбновление(f.Num)
        End If
    End Sub

    Private Sub Button2_Click_1(sender As Object, e As EventArgs) Handles Button2.Click
        Dim f As New ПредоплатаПеревозчик
        f.ShowDialog()
    End Sub

    Private Sub ВодительToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ВодительToolStripMenuItem.Click
        Dim f As New ВодитДан
        f.ShowDialog()
    End Sub

    Private Sub Условия20ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles Условия20ToolStripMenuItem.Click
        ДопПроц.ShowDialog()
    End Sub

    Private Sub ComboBox1_SelectedValueChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedValueChanged
        ПерегрЛист1()
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs)
        Штрафы.ShowDialog()
    End Sub

    Private Sub Сбор(ByVal Ном As Dictionary(Of String, String))
        'Dim mas() = {file3, file2}



        If Not Ном Is Nothing Then
            If Ном.ContainsKey("РейсыПеревозчика") Then
                file26 = IO.Directory.GetFiles("Z:\RICKMANS\", "*" & "ДОГОВОР" & "*" & Ном("РейсыПеревозчика") & "*", IO.SearchOption.AllDirectories)
                For i As Integer = 0 To file26.Count - 1
                    Process.Start(file26(i))
                Next
            Else
                file36 = IO.Directory.GetFiles("Z:\RICKMANS\", "*" & "ДОГОВОР" & "*" & Ном("РейсыКлиента") & "*", IO.SearchOption.AllDirectories)
                For i As Integer = 0 To file36.Count - 1
                    Process.Start(file36(i))
                Next
            End If


        End If


    End Sub
    Public Sub ДокиОбновление(Optional ByVal Num As Integer = 0)
        If Num > 0 Then
            СлРейс = Num
        End If
        Me.Cursor = Cursors.WaitCursor
        Dim xlapp As Microsoft.Office.Interop.Excel.Application
        xlapp = New Microsoft.Office.Interop.Excel.Application

        'Dim XXX = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=U:\Офис\Рикманс\ДанныеРикманс.accdb; Persist Security Info=False;"
        'Dim XXX = "Provider='SQLOLEDB';Data Source=45.14.50.13\723\SQLEXPRESS,1433;Network Library=DBMSSOCN;Initial Catalog=Rickmans;User ID=userOleg;Password=Zf6VpP37Ol"
        ' Dim XXX = "Provider='SQLOLEDB';Data Source=45.14.50.142\2749\SQLEXPRESS,1433;Network Library=DBMSSOCN;Initial Catalog=Rickmans;User ID=userOleg1;Password=Zf6VpP37Ol"
        Dim XXX = "Provider='SQLOLEDB';Data Source=178.124.211.175,52891;Network Library=DBMSSOCN;Initial Catalog=Rickmans;User ID=Rickmans;Password=Zf6VpP37Ol"

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

        Dim ds1 As DataTable = Selects3(strSQL)
        Dim ds2 As DataTable = Selects3(strSQL1)

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
        Dim f As New ИзменПорНомерКлиента2(НомРес, ComboBox3.Text, ComboBox4.Text)
        f.ShowDialog()
        If f.Rez IsNot Nothing Then
            ИзменВДействРейсе(f.Rez)
        End If

    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click

        If MessageBox.Show("Удалить " & НомРес & " Рейс?", Рик, MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.No Then
            Exit Sub
        End If

        Using db As New dbAllDataContext()
            Dim f = db.РейсыПеревозчика.Where(Function(x) x.НомерРейса = НомРес).Select(Function(x) x).FirstOrDefault()
            If f IsNot Nothing Then
                db.РейсыПеревозчика.DeleteOnSubmit(f)
                db.SubmitChanges()
            End If
            Dim f1 = db.РейсыКлиента.Where(Function(x) x.НомерРейса = НомРес).Select(Function(x) x).FirstOrDefault()
            If f1 IsNot Nothing Then
                db.РейсыКлиента.DeleteOnSubmit(f1)
                db.SubmitChanges()
            End If
        End Using
        Dim mo As New AllUpd
        mo.РейсыКлиентаAllAsync()
        mo.РейсыПеревозчикаAllAsync()


        'Dim strsql As String = "DELETE FROM РейсыПеревозчика WHERE НомерРейса=" & НомРес & ""
        'Updates3(strsql)

        'Dim strsql1 As String = "DELETE FROM РейсыКлиента WHERE НомерРейса=" & НомРес & ""
        'Updates3(strsql1)

        If Not IO.Directory.Exists("Z:\RICKMANS\" & ComboBox1.Text & "\СОРВАННЫЕ ЗАГРУЗКИ\") Then
            IO.Directory.CreateDirectory("Z:\RICKMANS\" & ComboBox1.Text & "\СОРВАННЫЕ ЗАГРУЗКИ\")
        End If

        IO.File.Copy("Z:\RICKMANS\" & ComboBox1.Text & "\" & ПутьРейса, "Z:\RICKMANS\" & ComboBox1.Text & "\СОРВАННЫЕ ЗАГРУЗКИ\" & ПутьРейса, True)
        IO.File.Delete("Z:\RICKMANS\" & ComboBox1.Text & "\" & ПутьРейса)

        MessageBox.Show("Рейс полностью удалён!", Рик)
        ПерегрДанныхИзБазы()
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
                ПерегрДанныхИзБазы()
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
        Button2.BackColor = Color.LightBlue
        Button7.BackColor = Color.LightBlue
        Me.Cursor = Cursors.Default
    End Sub
    Private Sub НовыйРейсГлавная()
        НовыйРейс()
        If Отмена = 1 Then Exit Sub
        Доки()
        ПерегрДанныхИзБазы()
        MessageBox.Show("Рейс оформлен!", Рик)
        If ПерезагрЛист1 = 0 Then
            ПерегрЛист1()
        End If

    End Sub
End Class
Public Class IDNaz
    Public Property ID As Integer
    Public Property Naz As String
End Class
Public Class СпискиВсе
    Public Function Клиенты()
        Dim mo As New AllUpd
        Do While AllClass.Клиент Is Nothing
            mo.КлиентAll()
        Loop
        Dim var = AllClass.Клиент.OrderBy(Function(x) x.НазваниеОрганизации).Select(Function(x) New IDNaz With {.Naz = x.НазваниеОрганизации}).ToList()
        If var IsNot Nothing Then
            If var.Count > 0 Then
                Return var
            Else
                Return Nothing
            End If
        Else
            Return Nothing
        End If

    End Function
    Public Function Перевозчики()
        Dim mo As New AllUpd
        Do While AllClass.Перевозчики Is Nothing
            mo.ПеревозчикиAll()
        Loop
        Dim var = AllClass.Перевозчики.OrderBy(Function(x) x.Названиеорганизации).Select(Function(x) New IDNaz With {.Naz = x.Названиеорганизации}).ToList()
        If var IsNot Nothing Then
            If var.Count > 0 Then
                Return var
            Else
                Return Nothing
            End If
        Else
            Return Nothing
        End If
    End Function

End Class
Public Class ПутиДоков
    Public Property Путь As String
    Public Property ПолныйПуть As String
End Class