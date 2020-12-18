Option Explicit On
Imports System.ComponentModel
Imports System.Data.OleDb
Imports System.Reflection
Imports System.Threading
Imports ClosedXML.Excel
Imports System.Data.SqlClient
'Imports TableDependency.SqlClient

'Imports Microsoft.Office.Interop.Excel

Public Class Рейс
    Dim file26(), file36() As String
    Dim ПутьПолный As String
    'Dim b As New Thread(AddressOf ЗапускБыстро)
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
    Private lst1all As BindingList(Of ПутиДоков)
    Private bslst11 As BindingSource
    Private com11all As List(Of IDNaz)
    Private bscom11 As BindingSource

    Private xlapp As Microsoft.Office.Interop.Excel.Application
    Private xlworkbook As Microsoft.Office.Interop.Excel.Workbook
    Private xlworksheet As Microsoft.Office.Interop.Excel.Worksheet
    Private misvalue As Object = Missing.Value

    Private PutPolnStroka As String
    Private PutCorStroka As String

    Dim arrtbox As New Dictionary(Of String, String)
    Dim arrtcom As New Dictionary(Of String, String)
    Dim arrtmask As New Dictionary(Of String, String)
    Dim arrtRichbox As New Dictionary(Of String, String)

    Private ClStrokaForDoc As New РейсыКлиента
    Private PerStrokaForDoc As New РейсыПеревозчика
    Private NewPutForListInNewRejs As ПутиДоков
    Private pls As Boolean = False
    Private plsForPrint As Boolean = False

    Private changeCount As Integer = 0

    Private list1selПуть As ПутиДоков

    Private NewListDynamFlag As Boolean = False
    Private listbxDyn As ListBox
    Private ClicnewLstDyn As Boolean = False

    'для SqlDependency
    'Dim conn7c7 As New SqlConnection(ConString)
    'Dim c7c As New SqlCommand("Select * from РейсыКлиента", conn7c7)
    'Private WithEvents Dep7с7 As SqlDependency = New SqlDependency(c7c)
    'Private Event OnChange As OnChangeEventHandler
    'Private Sub Событие(sender As Object, e As EventArgs) Handles Dep7с7.OnChange
    '    MessageBox.Show("Egc")

    'End Sub


    Public Sub COM4()

        If com4all IsNot Nothing Then
            com4all.Clear()
        End If
        com4all = New List(Of IDNaz)
        bscom4 = New BindingSource
        bscom4.DataSource = com4all
        ComboBox4.DataSource = bscom4
        ComboBox4.DisplayMember = "Naz"
        Com4Async()


    End Sub
    Private Async Sub Com4Async()
        Await Task.Run(Sub() Com4A())
        bscom4.ResetBindings(False)
        ComboBox4.Text = String.Empty
    End Sub
    Private Sub Com4A()
        Dim f As New СпискиВсе
        Dim m As List(Of IDNaz) = f.Перевозчики
        If m IsNot Nothing Then
            com4all.AddRange(m)
        End If


    End Sub

    Private Sub COM3()
        If com3all IsNot Nothing Then
            com3all.Clear()
        End If
        com3all = New List(Of IDNaz)
        bscom3 = New BindingSource
        bscom3.DataSource = com3all
        ComboBox3.DataSource = bscom3
        ComboBox3.DisplayMember = "Naz"
        Com3Async()



    End Sub
    Private Async Sub Com3Async()
        Await Task.Run(Sub() Com3A())
        bscom3.ResetBindings(False)
        ComboBox3.Text = String.Empty
    End Sub
    Private Sub Com3A()
        Dim f As New СпискиВсе
        Dim m As List(Of IDNaz) = f.Клиенты
        com3all.AddRange(m)
    End Sub

    Private Sub COM11()

        com11all = New List(Of IDNaz)
        bscom11 = New BindingSource
        bscom11.DataSource = com11all
        ComboBox11.DataSource = bscom11
        ComboBox11.DisplayMember = "Naz"


        Com11Async()

    End Sub

    Private Async Sub Com11Async()
        Await Task.Run(Sub() Com11A())
        bscom11.ResetBindings(False)
        ComboBox11.Text = String.Empty
    End Sub
    Private Sub Com11A()
        Dim mo As New AllUpd
        Do While AllClass.ТипАвто Is Nothing
            mo.ТипАвтоAll()
        Loop
        Try
            'Using db As New dbAllDataContext
            Dim mk = AllClass.ТипАвто.OrderBy(Function(x) x.ТипАвто).Select(Function(x) New IDNaz With {.Naz = x.ТипАвто}).ToList()
            If mk IsNot Nothing Then
                com11all.AddRange(mk)
            End If
            'End Using


        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try



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

        Do While AllClass.ТипАвто Is Nothing
            mo.ТипАвтоAll()
        Loop

    End Sub
    'Dim f = AllClass.РейсыКлиента.Where(Function(x) x.НомерРейса?.ToString = 591).Select(Function(x) x).FirstOrDefault()  ''перебор свойств класса количество
    'Dim type As Type = f.[GetType]()

    'Dim properties As PropertyInfo() = type.GetProperties()
    'Dim fo As New List(Of String)
    'For Each prope As PropertyInfo In properties
    'Try
    '            fo.Add(prope.GetValue(f, Nothing))
    '        Catch ex As Exception
    'Continue For
    'End Try

    'Next
    Private Function GetProp(Of T)(ByVal f As T) As PropertyInfo()
        Dim type As Type = f.[GetType]()
        Dim properties As PropertyInfo() = type.GetProperties()
        If properties IsNot Nothing Then
            Return properties
        Else
            Return Nothing
        End If

    End Function
    Private Async Sub properAsync()
        Try
            Await Task.Run(Sub() proper())
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
        ClearDiction()
    End Sub
    Private Sub proper()


        plsForPrint = True
        Dim mo As New AllUpd


        Do While AllClass.Клиент Is Nothing
            mo.КлиентAll()
        Loop

        Do While AllClass.Перевозчики Is Nothing
            mo.ПеревозчикиAll()
        Loop

        'xlapp = New Microsoft.Office.Interop.Excel.Application With {
        '    .Visible = False
        '}
        ''xlworkbook = New Microsoft.Office.Interop.Excel.Workbook
        'xlworkbook = xlapp.Workbooks.Add("Z:\RICKMANS\496 Ивановский_8 - Петровский_1.xlsm")
        ''xlworksheet = New Microsoft.Office.Interop.Excel.Worksheet
        'xlworksheet = xlworkbook.Sheets("ЗАК")




        Do While pls = False
            'Await Task.Delay(1000)
        Loop

        Dim f = ClStrokaForDoc

        ''перебор свойств класса количество
        Dim gtp = GetProp(f)

        Dim i As Integer = 0
        Dim fo As New List(Of String)
        For Each prope As PropertyInfo In gtp
            Try
                fo.Add(prope.GetValue(f, Nothing))
                If prope.Name = "ВремяПодачи" Or prope.Name = "ВремяПодачиВыгРаст" Then
                    Dim msd = CType(Replace(fo(i), ":", "ч. "), String)
                    xlworksheet.Cells(3, 12 + i) = Strings.Left(msd, 5) & Strings.Right(msd, 2) & "мин."

                Else
                    xlworksheet.Cells(3, 12 + i) = fo(i)
                End If


                i += 1
            Catch ex As Exception
                i += 1
                Continue For
            End Try

        Next
        '/
        Dim f1 = PerStrokaForDoc
        ''перебор свойств класса количество

        Dim gtp1 = GetProp(f1)

        Dim i1 As Integer = 0
        Dim fo1 As New List(Of String)
        For Each prope As PropertyInfo In gtp1
            Try
                fo1.Add(prope.GetValue(f1, Nothing))
                If prope.Name = "ВремяПодачи" Or prope.Name = "ВремяПодачиВыгРаст" Then
                    Dim msd = CType(Replace(fo1(i1), ":", "ч. "), String)

                    xlworksheet.Cells(6, 12 + i1) = Strings.Left(msd, 5) & Strings.Right(msd, 2) & "мин."
                Else
                    xlworksheet.Cells(6, 12 + i1) = fo1(i1)
                End If
                i1 += 1
            Catch ex As Exception
                i1 += 1
                Continue For
            End Try

        Next

        '/
        Dim clias As String = arrtcom("ComboBox3")
        Dim f2 = (From x In AllClass.Клиент
                  Where x.НазваниеОрганизации = clias
                  Select x).FirstOrDefault()

        ''перебор свойств класса количество
        Dim gtp2 = GetProp(f2)

        Dim i2 As Integer = 0
        Dim fo2 As New List(Of String)
        For Each prope As PropertyInfo In gtp2
            Try
                fo2.Add(prope.GetValue(f2, Nothing))
                xlworksheet.Cells(4, 12 + i2) = fo2(i2)
                i2 += 1
            Catch ex As Exception
                i2 += 1
                Continue For
            End Try

        Next


        '/
        Dim f3 = (From x In AllClass.Перевозчики
                  Where x.Названиеорганизации = arrtcom("ComboBox4")
                  Select x).FirstOrDefault()

        ''перебор свойств класса количество
        Dim gtp3 = GetProp(f3)

        Dim i3 As Integer = 0
        Dim fo3 As New List(Of String)
        For Each prope As PropertyInfo In gtp3
            Try
                fo3.Add(prope.GetValue(f3, Nothing))
                xlworksheet.Cells(5, 12 + i3) = fo3(i3)
                i3 += 1
            Catch ex As Exception
                i3 += 1
                Continue For
            End Try

        Next


        Dim ya = Now.Year.ToString

        xlworkbook.SaveAs(Filename:=("Z:\RICKMANS\" & ya & "\" & СлРейс & " " & arrtcom("ComboBox3") & " " & СлПорРейсКл & " - " & arrtcom("ComboBox4") & " " & СлПорРейсПер & ".xlsm"), FileFormat:=52, CreateBackup:=False)
        'xlworkbook.SaveCopyAs("Z:\RICKMANS\Test2.xlsx")

        xlworkbook.Close(True)
        xlapp.Quit()

        'Dim f9 As New List(Of Object)
        'f9.AddRange({xlapp, xlworkbook, xlworksheet})
        'releaseobjectAsync(f9)
        releaseobject(xlapp)
        releaseobject(xlworkbook)
        releaseobject(xlworksheet)

        'ПредExcel()
        plsForPrint = False
    End Sub
    Private Async Sub ПредExcelAsync()
        pls = False
        Await Task.Run(Sub() ПредExcel())

    End Sub
    Private Sub ПредExcel()


        xlapp = New Microsoft.Office.Interop.Excel.Application With {
            .Visible = False
        }
        'xlworkbook = New Microsoft.Office.Interop.Excel.Workbook
        xlworkbook = xlapp.Workbooks.Add("Z:\RICKMANS\496 Ивановский_8 - Петровский_1.xlsm")
        'xlworksheet = New Microsoft.Office.Interop.Excel.Worksheet
        xlworksheet = xlworkbook.Sheets("ЗАК")
        pls = True
    End Sub
    'Public Event OnChanged As TableDependency.SqlClient.EventArgs.SqlRecordChangedEventArgs(sender As Object, e As EventArgs)


    'Private Sub nx()
    '    Dim mapper = New TableDependency.SqlClient.Base.ModelToTableMapper(Of РейсыКлиента)()
    '    mapper.AddMapping(Function(c) c.НомерРейса, "НомерРейса")
    '    mapper.AddMapping(Function(c) c.НазвОрганизации, "НазвОрганизации")
    '    Using bd As New TableDependency.SqlClient.SqlTableDependency(Of РейсыКлиента)(ConString, "РейсыКлиента", mapper:=mapper)
    '        'bd.OnChanged += Changed2()
    '        AddHandler bd.OnChanged, AddressOf Changed2 = AddHandler bd.OnChanged + AddHandler bd.OnChanged, AddressOf Changed2

    '        'bd.OnChanged()
    '        bd.Start()

    '        'Console.WriteLine("Press a key to exit")
    '        'Console.ReadKey()

    '        bd.Stop()
    '    End Using


    'End Sub

    'Public Shared Sub Changed2(ByVal sender As Object, ByVal e As TableDependency.SqlClient.Base.EventArgs.RecordChangedEventArgs(Of РейсыКлиента))
    '    Dim changedEntity = e.Entity
    '    MessageBox.Show("DML operation: " & e.ChangeType)
    '    'Console.WriteLine("DML operation: " & e.ChangeType)
    '    'Console.WriteLine("ID: " & changedEntity.НомерРейса)
    '    'Console.WriteLine("Name: " & changedEntity.Name)
    '    'Console.WriteLine("Surname: " & changedEntity.Surname)
    'End Sub
    ''Private Sub ExecuteWatchingQuery()
    '    Using connection As SqlConnection = New SqlConnection(ConString)
    '        connection.Open()
    '        Using command = New SqlCommand("SELECT НомерРейса FROM РейсыКлиента", connection)
    '            Dim sqlDependency = New SqlDependency(command)
    '            AddHandler sqlDependency.OnChange, AddressOf OnDatabaseChange
    '            command.ExecuteReader()
    '        End Using
    '    End Using
    'End Sub
    'Private Sub OnDatabaseChange(ByVal sender As Object, ByVal args As SqlNotificationEventArgs)
    '    Try
    '        Dim info As SqlNotificationInfo = args.Info
    '        If SqlNotificationInfo.Insert.Equals(info) Or SqlNotificationInfo.Update.Equals(info) Or SqlNotificationInfo.Delete.Equals(info) Then
    '            changeCount += 1
    '            Invoke(Sub()
    '                       Label1.Text = changeCount & " changes have occurred."
    '                       With ListBox1.Items
    '                           .Clear()
    '                           .Add("Info:   " & args.Info.ToString())
    '                           .Add("Source: " & args.Source.ToString())
    '                           .Add("Type:   " & args.Type.ToString())
    '                       End With
    '                   End Sub)
    '        End If
    '        ExecuteWatchingQuery()
    '    Catch ex As Exception
    '        MsgBox(ex.StackTrace)
    '    End Try
    'End Sub


    Private Sub Рейс_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        UpdListForTime(Now.Year)
        'proper()
        'ПредExcelAsync()
        'SqlDependency.Stop(ConString)
        'SqlDependency.Start(ConString)
        'ExecuteWatchingQuery()




        ПредзагрузкаAsync()
        Me.Cursor = Cursors.WaitCursor
        lst1all = New BindingList(Of ПутиДоков)
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
        'Dim XXX = "Provider='SQLOLEDB';Data Source=45.14.50.142\2749\SQLEXPRESS,1433;Network Library=DBMSSOCN;Initial Catalog=Rickmans;User ID=userOleg1;Password=Zf6VpP37Ol"
        Dim XXX = "Provider='SQLOLEDB';Data Source=178.124.211.175,52891;Network Library=DBMSSOCN;Initial Catalog=Rickmans;User ID=Rickmans;Password=Zf6VpP37Ol"

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
        'dtVyborkaS()
        'dtVyborkaS1()
        Перевозчики2()
        Клиенты2()
    End Sub


    Private Sub ПерегрЛист1(Optional f As ПутиДоков = Nothing, Optional com1 As String = Nothing)
        If f Is Nothing Then
            Dim cm1 As String

            If arrtcom.Count = 0 And com1 Is Nothing Then
                cm1 = ComboBox1.Text
            ElseIf arrtcom.Count > 0 And com1 Is Nothing Then
                cm1 = arrtcom("ComboBox1")
            Else
                cm1 = com1
            End If
            ListBox1.BeginUpdate()
            Справки(CType(cm1, Integer))


            If Files3.Count > 0 Then
                If lst1all IsNot Nothing Then
                    lst1all.Clear()
                End If
                For Each b2 In Files3
                    lst1all.Add(b2)
                Next

                'bslst11.ResetBindings(False)


                ListBox1.Text = lst1all.Select(Function(x) x.Путь).LastOrDefault()


            End If
        Else
            ListBox1.BeginUpdate()
            lst1all.Add(f)
            'bslst11.ResetBindings(False)
            ListBox1.Text = lst1all.Select(Function(x) x.Путь).LastOrDefault()
        End If

        ListBox1.EndUpdate()
    End Sub
    Private Async Sub ListBox1_DoubleClick(sender As Object, e As EventArgs) Handles ListBox1.DoubleClick
        If ListBox1.SelectedIndex = -1 Then
            MessageBox.Show("Выберите документ для просмотра!", Рик, MessageBoxButtons.OK)
            Exit Sub
        End If
        Cursor = Cursors.WaitCursor

        Do While plsForPrint = True

        Loop

        'timer2.Change(Timeout.Infinite, Timeout.Infinite)  'остановка таймера


        Dim f As ПутиДоков = ListBox1.SelectedItem
        'Process.Start("Z:\RICKMANS\" & ComboBox1.Text & "\" & f.Путь)
        If IO.File.Exists(f.ПолныйПуть) Then
            Process.Start(f.ПолныйПуть)
            Cursor = Cursors.Default
        Else
            Cursor = Cursors.WaitCursor
            Do While IO.File.Exists(f.ПолныйПуть) = False
                Await Task.Delay(1000)
            Loop
            Cursor = Cursors.Default
            Process.Start(f.ПолныйПуть)
        End If


        'UpdListForTime(ComboBox1.Text)  'запуск таймера


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

        Dim f = AllClass.РейсыКлиента.Count()
        Dim f1 As Integer

        Do While AllClass.РейсыПеревозчика Is Nothing
            mo.РейсыПеревозчикаAll()
        Loop

        Dim f2 = AllClass.РейсыПеревозчика.Count
        Dim f3 As Integer
        Using db As New dbAllDataContext()
            f1 = db.РейсыКлиента.Count()
            f3 = db.РейсыПеревозчика.Count()
            If Not f1 = f Then
                mo.РейсыКлиентаAll()
            End If
            If Not f2 = f3 Then
                mo.РейсыПеревозчикаAll()
            End If
        End Using


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
                list1selПуть = f


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
        Do While AllClass.Клиент Is Nothing
            mo.КлиентAll()
        Loop


        Dim f As IDNaz = ComboBox3.SelectedItem
        Dim ds5 As Клиент
        Dim com3 As String = f.Naz

        ds5 = AllClass.Клиент.Where(Function(x) x.НазваниеОрганизации = com3).Select(Function(x) x).FirstOrDefault()

        Dim f2 As Integer
        If com3 = "Виталюр" Then
            f2 = (From x In AllClass.РейсыКлиента
                  Order By x.КоличРейсов
                  Where x.НазвОрганизации = "Виталюр" And Format(x.Год, "yyyy") = Now.Year
                  Select x.КоличРейсов).LastOrDefault()

        Else
            f2 = AllClass.РейсыКлиента.OrderBy(Function(x) x.КоличРейсов).Where(Function(x) x.НазвОрганизации = com3).Select(Function(x) IIf(x.КоличРейсов Is Nothing, 0, x.КоличРейсов)).LastOrDefault()

        End If


        Try
            СлПорРейсКл = f2 + 1
        Catch ex As Exception

            Nm = ""
            Nm = com3
            Dim f4 As New ПорНомер(com3)
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
            СлПорРейсКл = f4.SledPorRejsClient
            Пров = 1
        End Try

        If Пров = 0 Then
            Dim f1, g As Integer
            Dim f5 As New ПорНомер(com3)
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


        If ds2 Is Nothing Then
            ds2 = 0
        End If
        Try
            СлПорРейсПер = ds2 + 1
        Catch ex As Exception
            Nm = ""
            Nm = ComboBox4.Text
            pro = 0
            Dim f4 As New ПорНомер(f.Naz)
            If ds6.Договор = "" Or ds6.Дата = "" Then
                ПорНомер.GroupBox1.Enabled = True
            End If
            If ds6.Должность = "" Or ds6.ФИОРуководителя = "" Then
                ПорНомер.GroupBox2.Enabled = True
            End If
            f4.TextBox1.Enabled = True
            f4.ListBox1.Enabled = True
            f4.ShowDialog()
            If Отмена = 1 Then Exit Sub
            СлПорРейсПер = f4.SledPorRejsPer
            Пров = 1
        End Try

        If Пров = 0 Then
            Dim f2, g As Integer
            Dim f3 As New ПорНомер(f.Naz)
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
        Dim cm3 As String = arrtcom("ComboBox13")

        If cm3.Length = 0 Then
            ДогПор = "0"
            ДогПорЭксп = "0"
            ПорЭксп = "0"
        ElseIf cm3 = "Договор-поручение" Then
            ДогПор = "1"
            ДогПорЭксп = "0"
            ПорЭксп = "0"
        ElseIf cm3 = "Договор-поручение эксп" Then
            ДогПор = "0"
            ДогПорЭксп = "1"
            ПорЭксп = "0"
        ElseIf cm3 = "Поручение с экспедицией" Then
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

        Dim cm12 As String = arrtcom("ComboBox12")
        If cm12.Length = 0 Then
            ДогПор = "0"
            ДогПорЭксп = "0"
            ПорЭксп = "0"
        ElseIf cm12 = "Договор-поручение" Then
            ДогПор = "1"
            ДогПорЭксп = "0"
            ПорЭксп = "0"
        ElseIf cm12 = "Договор-поручение эксп" Then
            ДогПор = "0"
            ДогПорЭксп = "1"
            ПорЭксп = "0"
        ElseIf cm12 = "Поручение с экспедицией" Then
            ДогПор = "0"
            ДогПорЭксп = "0"
            ПорЭксп = "1"
        Else
            ДогПор = "0"
            ДогПорЭксп = "0"
            ПорЭксп = "0"
        End If



    End Sub
    Private Sub UpddateНовыйКлиент()
        Dim mo As New AllUpd
        mo.РейсыПеревозчикаAll()
        mo.РейсыКлиентаAll()
    End Sub
    Private Async Sub НовыйРейсСохранениеБазаAsync(ByVal cm4 As IDNaz)
        Await Task.Run(Sub() НовыйРейсСохранениеБаза(cm4))
        Await Task.Run(Sub() UpddateНовыйКлиент())
    End Sub
    Public Class ДогПодрCXlass
        Public Property ДогПор As String
        Public Property ДогПорЭксп As String
        Public Property ПорЭксп As String
    End Class
    Private Function ComB13A() As ДогПодрCXlass

        Dim cm3 As String = arrtcom("ComboBox13")

        If cm3.Length = 0 Then
            Dim m As New ДогПодрCXlass With {.ДогПор = "0", .ДогПорЭксп = "0", .ПорЭксп = "0"}
            Return m
        ElseIf cm3 = "Договор-поручение" Then
            Dim m As New ДогПодрCXlass With {.ДогПор = "1", .ДогПорЭксп = "0", .ПорЭксп = "0"}
            Return m

        ElseIf cm3 = "Договор-поручение эксп" Then
            Dim m As New ДогПодрCXlass With {.ДогПор = "0", .ДогПорЭксп = "1", .ПорЭксп = "0"}
            Return m

        ElseIf cm3 = "Поручение с экспедицией" Then
            Dim m As New ДогПодрCXlass With {.ДогПор = "0", .ДогПорЭксп = "0", .ПорЭксп = "1"}
            Return m

        Else
            Dim m As New ДогПодрCXlass With {.ДогПор = "0", .ДогПорЭксп = "0", .ПорЭксп = "0"}
            Return m
        End If
    End Function

    Private Function ComB12A() As ДогПодрCXlass


        Dim cm12 As String = arrtcom("ComboBox12")

        If cm12.Length = 0 Then
            Dim m As New ДогПодрCXlass With {.ДогПор = "0", .ДогПорЭксп = "0", .ПорЭксп = "0"}
            Return m

        ElseIf cm12 = "Договор-поручение" Then
            Dim m As New ДогПодрCXlass With {.ДогПор = "1", .ДогПорЭксп = "0", .ПорЭксп = "0"}
            Return m

        ElseIf cm12 = "Договор-поручение эксп" Then
            Dim m As New ДогПодрCXlass With {.ДогПор = "0", .ДогПорЭксп = "1", .ПорЭксп = "0"}
            Return m

        ElseIf cm12 = "Поручение с экспедицией" Then
            Dim m As New ДогПодрCXlass With {.ДогПор = "0", .ДогПорЭксп = "0", .ПорЭксп = "1"}
            Return m

        Else
            Dim m As New ДогПодрCXlass With {.ДогПор = "0", .ДогПорЭксп = "0", .ПорЭксп = "0"}
            Return m
        End If



    End Function
    Private Sub НовыйРейсСохранениеБаза(ByVal cm4 As IDNaz)

        Dim k = ComB13A()

        Using db As New dbAllDataContext()
            Dim f1 As New РейсыПеревозчика
            With f1
                .НазвОрганизации = cm4.Naz
                .НомерРейса = СлРейс
                .КоличРейсов = СлПорРейсПер
                .Маршрут = Trim(arrtRichbox("RichTextBox10"))
                .ДатаПодачиПодЗагрузку = arrtbox("TextBox5")
                .ВремяПодачи = arrtmask("MaskedTextBox3")
                .ДатаПодачиПодРастаможку = arrtbox("TextBox6")
                .ВремяПодачиВыгРаст = arrtmask("MaskedTextBox4")
                .ТочныйАдресЗагрузки = Trim(arrtRichbox("RichTextBox3"))
                .АдресЗатаможки = Trim(arrtRichbox("RichTextBox4"))
                .НаименованиеГруза = Trim(arrtRichbox("RichTextBox7"))
                .ТипТрСредства = arrtcom("ComboBox11")
                .НомерАвтомобиля = Trim(arrtRichbox("RichTextBox8"))
                .Водитель = Trim(arrtRichbox("RichTextBox9"))
                .ТочнАдресРаста = Trim(arrtRichbox("RichTextBox5"))
                .ТочнАдресРазгр = Trim(arrtRichbox("RichTextBox6"))
                .СтоимостьФрахта = arrtbox("TextBox2")
                .Валюта = arrtcom("ComboBox6")
                .ВалютаПлатежа = arrtcom("ComboBox7")
                .СрокОплаты = arrtbox("TextBox3")
                .ДопУсловия = Trim(arrtRichbox("RichTextBox2"))
                .ДогПор = k.ДогПор
                .ДогПорЭксп = k.ДогПорЭксп
                .ДатаПоручения = arrtmask("MaskedTextBox1")
                .ПорЭксп = k.ПорЭксп
                .УсловияОплаты = arrtcom("ComboBox9")
                .РазмерШтрафаЗаСрыв = ШтрафКлиент
                .Предоплата = ЧастичнаяОплатаПеревозчик
                .СрывЗагр20Проц = Procenty20
                .ДатаСоздания = Now
                .Экспедитор = Экспедитор
            End With
            db.РейсыПеревозчика.InsertOnSubmit(f1)
            db.SubmitChanges()
        End Using


        Dim k1 = ComB12A()
        ComB5()

        Using db As New dbAllDataContext()
            Dim var As New РейсыКлиента
            var.НазвОрганизации = arrtcom("ComboBox3")
            var.НомерРейса = СлРейс
            var.КоличРейсов = СлПорРейсКл
            var.Маршрут = Trim(arrtRichbox("RichTextBox10"))
            var.ДатаПодачиПодЗагрузку = arrtbox("TextBox5")
            var.ВремяПодачи = arrtmask("MaskedTextBox3")
            var.ДатаПодачиПодРастаможку = arrtbox("TextBox6")
            var.ВремяПодачиВыгРаст = arrtmask("MaskedTextBox4")
            var.ТочныйАдресЗагрузки = Trim(arrtRichbox("RichTextBox3"))
            var.АдресЗатаможки = Trim(arrtRichbox("RichTextBox4"))
            var.НаименованиеГруза = Trim(arrtRichbox("RichTextBox7"))
            var.ТипТрСредства = arrtcom("ComboBox11")
            var.НомерАвтомобиля = Trim(arrtRichbox("RichTextBox8"))
            var.Водитель = Trim(arrtRichbox("RichTextBox9"))
            var.ТочнАдресРаста = Trim(arrtRichbox("RichTextBox5"))
            var.ТочнАдресРазгр = Trim(arrtRichbox("RichTextBox6"))
            var.СтоимостьФрахта = arrtbox("TextBox1")
            var.Валюта = arrtcom("ComboBox5")
            var.ВалютаПлатежа = arrtcom("ComboBox8")
            var.СрокОплаты = arrtbox("TextBox4")
            var.ДопУсловия = Trim(arrtRichbox("RichTextBox1"))
            var.ДогПор = k1.ДогПор
            var.ДогПорЭксп = k1.ДогПорЭксп
            var.ДатаПоручения = arrtmask("MaskedTextBox1")
            var.ПорЭксп = k1.ПорЭксп
            var.УсловияОплаты = arrtcom("ComboBox10")
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



    Private Sub НовыйРейс()

        ПровСледРейсКлиент()
        If Отмена = 1 Then Exit Sub

        Пров = 0

        Dim mo As New AllUpd
        Do While AllClass.РейсыКлиента Is Nothing
            mo.РейсыКлиентаAll()
        Loop

        Dim ds1 As Integer = AllClass.РейсыКлиента.OrderBy(Function(x) x.НомерРейса).Select(Function(x) x.НомерРейса).LastOrDefault()      '

        СлРейс = ds1 + 1

        Dim cm4 As IDNaz = ComboBox4.SelectedItem
        ПровСледРейсПер()
        If Отмена = 1 Then Return

        Dim ya As String = Now.Year.ToString
        PutPolnStroka = "Z:\RICKMANS\" & ya & "\" & СлРейс & " " & ComboBox3.Text & " " & СлПорРейсКл & " - " & ComboBox4.Text & " " & СлПорРейсПер & ".xlsm"
        PutCorStroka = СлРейс & " " & ComboBox3.Text & " " & СлПорРейсКл & " - " & ComboBox4.Text & " " & СлПорРейсПер & ".xlsm"

        NewPutForListInNewRejs = New ПутиДоков With {.Путь = PutCorStroka, .ПолныйПуть = PutPolnStroka} 'переменная хранить пути для листбокс добавления

        НовыйРейсСохранениеБазаAsync(cm4) 'сохранение в базу

        'создаем данные для документов

        ClStrokaForDoc = Nothing
        PerStrokaForDoc = Nothing

        ComB13()
        Dim f1 As New РейсыПеревозчика
        With f1
            .НазвОрганизации = cm4.Naz
            .НомерРейса = СлРейс
            .КоличРейсов = СлПорРейсПер
            .Маршрут = Trim(arrtRichbox("RichTextBox10"))
            .ДатаПодачиПодЗагрузку = arrtbox("TextBox5")
            .ВремяПодачи = arrtmask("MaskedTextBox3")
            .ДатаПодачиПодРастаможку = arrtbox("TextBox6")
            .ВремяПодачиВыгРаст = arrtmask("MaskedTextBox4")
            .ТочныйАдресЗагрузки = Trim(arrtRichbox("RichTextBox3"))
            .АдресЗатаможки = Trim(arrtRichbox("RichTextBox4"))
            .НаименованиеГруза = Trim(arrtRichbox("RichTextBox7"))
            .ТипТрСредства = arrtcom("ComboBox11")
            .НомерАвтомобиля = Trim(arrtRichbox("RichTextBox8"))
            .Водитель = Trim(arrtRichbox("RichTextBox9"))
            .ТочнАдресРаста = Trim(arrtRichbox("RichTextBox5"))
            .ТочнАдресРазгр = Trim(arrtRichbox("RichTextBox6"))
            .СтоимостьФрахта = arrtbox("TextBox2")
            .Валюта = arrtcom("ComboBox6")
            .ВалютаПлатежа = arrtcom("ComboBox7")
            .СрокОплаты = arrtbox("TextBox3")
            .ДопУсловия = Trim(arrtRichbox("RichTextBox2"))
            .ДогПор = ДогПор
            .ДогПорЭксп = ДогПорЭксп
            .ДатаПоручения = arrtmask("MaskedTextBox1")
            .ПорЭксп = ПорЭксп
            .УсловияОплаты = arrtcom("ComboBox9")
            .РазмерШтрафаЗаСрыв = ШтрафКлиент
            .Предоплата = ЧастичнаяОплатаПеревозчик
            .СрывЗагр20Проц = Procenty20
            .ДатаСоздания = Now
            .Экспедитор = Экспедитор
        End With

        PerStrokaForDoc = f1



        ComB12()
        ComB5()
        Dim var As New РейсыКлиента
        With var
            .НазвОрганизации = arrtcom("ComboBox3")
            .НомерРейса = СлРейс
            .КоличРейсов = СлПорРейсКл
            .Маршрут = Trim(arrtRichbox("RichTextBox10"))
            .ДатаПодачиПодЗагрузку = arrtbox("TextBox5")
            .ВремяПодачи = arrtmask("MaskedTextBox3")
            .ДатаПодачиПодРастаможку = arrtbox("TextBox6")
            .ВремяПодачиВыгРаст = arrtmask("MaskedTextBox4")
            .ТочныйАдресЗагрузки = Trim(arrtRichbox("RichTextBox3"))
            .АдресЗатаможки = Trim(arrtRichbox("RichTextBox4"))
            .НаименованиеГруза = Trim(arrtRichbox("RichTextBox7"))
            .ТипТрСредства = arrtcom("ComboBox11")
            .НомерАвтомобиля = Trim(arrtRichbox("RichTextBox8"))
            .Водитель = Trim(arrtRichbox("RichTextBox9"))
            .ТочнАдресРаста = Trim(arrtRichbox("RichTextBox5"))
            .ТочнАдресРазгр = Trim(arrtRichbox("RichTextBox6"))
            .СтоимостьФрахта = arrtbox("TextBox1")
            .Валюта = arrtcom("ComboBox5")
            .ВалютаПлатежа = arrtcom("ComboBox8")
            .СрокОплаты = arrtbox("TextBox4")
            .ДопУсловия = Trim(arrtRichbox("RichTextBox1"))
            .ДогПор = ДогПор
            .ДогПорЭксп = ДогПорЭксп
            .ДатаПоручения = arrtmask("MaskedTextBox1")
            .ПорЭксп = ПорЭксп
            .УсловияОплаты = arrtcom("ComboBox10")
            .Год = Now.ToShortDateString
            .РазмерШтрафаЗаСрыв = ШтрафПер
            .Предоплата = ЧастичнаяОплатаКлиент
            .ОплатаПоКурсу = ОплатаПоКурсу
            .Экспедитор = Экспедитор
            .ДатаСоздания = Now
        End With

        ClStrokaForDoc = var 'данные для документов



    End Sub
    Private Sub ComB5()
        Dim cm5 As String = arrtcom("ComboBox5")
        Dim cm8 As String = arrtcom("ComboBox8")
        If Not cm5 = "Рубль" And cm8 = "BYN" Then
            ОплатаПоКурсу = "True"
        Else
            ОплатаПоКурсу = "False"
        End If
    End Sub
    Private Sub ДокиNew()
        Dim str As String = "Z:\RICKMANS\496 Ивановский_8 - Петровский_1.xlsm"
        Dim xml = New XLWorkbook(str)
        Dim wok = xml.Worksheet("ЗАК")
        Dim mo As New AllUpd

        Do While AllClass.РейсыКлиента Is Nothing
            mo.РейсыКлиентаAll()
        Loop

        Do While AllClass.РейсыПеревозчика Is Nothing
            mo.РейсыПеревозчикаAll()
        Loop


        Do While AllClass.Клиент Is Nothing
            mo.КлиентAll()
        Loop
        Do While AllClass.Перевозчики Is Nothing
            mo.ПеревозчикиAll()
        Loop

        Dim КлиентРейс = (From x In AllClass.РейсыКлиента
                          Where x.НомерРейса = СлРейс
                          Select New РейсыКлиента With {.АдресЗатаможки = IIf(x.АдресЗатаможки Is Nothing, String.Empty, x.АдресЗатаможки),
.Валюта = IIf(x.Валюта Is Nothing, String.Empty, x.Валюта), .ВалютаПлатежа = IIf(x.ВалютаПлатежа Is Nothing, String.Empty, x.ВалютаПлатежа)}).FirstOrDefault()

        Dim КлиентРейс2 = (From x In AllClass.РейсыКлиента
                           Where x.НомерРейса = СлРейс
                           Select New With {x.АдресЗатаможки, x.Валюта, x.ВалютаПлатежа, x.Водитель, x.ВремяПодачи, x.ВремяПодачиВыгРаст, x.Год,
                               x.ДатаАкта, x.ДатаЗаявки, x.ДатаОплаты, x.ДатаОтправкиДоков, x.ДатаПодачиПодЗагрузку})?.ToList()


        Dim df As New List(Of String)
        df.Add(КлиентРейс.АдресЗатаможки)
        df.Add(КлиентРейс.Валюта)
        df.Add(КлиентРейс.ВалютаПлатежа)

        If КлиентРейс IsNot Nothing Then
            'wok.Cell("L3").SetValue(df)
            Dim var = wok.Cell("L3").InsertData(КлиентРейс2.AsEnumerable)
        End If
        xml.SaveAs("B:\124.xlsm")

    End Sub
    Private Sub Доки()

        Me.Cursor = Cursors.WaitCursor


        'Dim xlapp As Application
        'xlapp = New Application

        'Dim XXX = "Provider='SQLOLEDB';Data Source=178.124.211.175,52891;Network Library=DBMSSOCN;Initial Catalog=Rickmans;Persist Security Info=True;User ID=Rickmans;Password=Zf6VpP37Ol"

        'Dim CON As New ADODB.Connection
        'Dim RS As New ADODB.Recordset
        'Dim RS1 As New ADODB.Recordset
        'Dim RS2 As New ADODB.Recordset
        'Dim RS3 As New ADODB.Recordset

        'Dim strSQL, strSQL1, strSQL2, strSQL3 As String
        'strSQL = "select * from РейсыКлиента WHERE НомерРейса=" & СлРейс & ""
        'strSQL1 = "select * from РейсыПеревозчика WHERE НомерРейса=" & СлРейс & ""
        'strSQL2 = "select * from Клиент WHERE НазваниеОрганизации='" & ComboBox3.Text & "'"
        'strSQL3 = "select * from перевозчики WHERE Названиеорганизации='" & ComboBox4.Text & "'"


        'CON.ConnectionString = XXX
        'CON.Open()

        'RS.Open(strSQL, CON)
        'RS1.Open(strSQL1, CON)
        'RS2.Open(strSQL2, CON)
        'RS3.Open(strSQL3, CON)


        '/////
        'proper()
        properAsync()
        '/////

        'xlapp.Workbooks.Open("Z:\RICKMANS\496 Ивановский_8 - Петровский_1.xlsm").Worksheets("ЗАК").Range("L3").CopyFromRecordset(RS)
        ''xlapp.Workbooks.Open("Z:\RICKMANS\496 Ивановский_8 - Петровский_1.xlsm").Worksheets("ЗАК").Range("L3").CopyFromRecordset(dt)
        'xlapp.Workbooks("496 Ивановский_8 - Петровский_1.xlsm").Worksheets("ЗАК").Range("L6").CopyFromRecordset(RS1)
        'xlapp.Workbooks("496 Ивановский_8 - Петровский_1.xlsm").Worksheets("ЗАК").Range("L4").CopyFromRecordset(RS2)
        'xlapp.Workbooks("496 Ивановский_8 - Петровский_1.xlsm").Worksheets("ЗАК").Range("L5").CopyFromRecordset(RS3)
        'xlapp.Workbooks("496 Ивановский_8 - Петровский_1.xlsm").Close(True)

        'Dim ya As Integer = Now.Year
        'If ya = 0 Then
        '    ya = Now.Year
        'End If
        'IO.File.Copy("Z:\RICKMANS\496 Ивановский_8 - Петровский_1.xlsm", "Z:\RICKMANS\" & ya & "\" & СлРейс & " " & ComboBox3.Text & " " & СлПорРейсКл & " - " & ComboBox4.Text & " " & СлПорРейсПер & ".xlsm")



        'xlapp.Quit()
        'releaseobject(xlapp)


        ''Excel_Doc.Range("a1").CopyFromRecordset(RS)   'копируем рекордсет на лист Excel без заморочек
        'RS = Nothing
        'RS1 = Nothing
        'RS2 = Nothing
        'RS3 = Nothing
        'CON.Close()

        Me.Cursor = Cursors.Default
    End Sub
    Private Async Sub releaseobjectAsync(ByVal obj As List(Of Object))
        Dim f As New List(Of Object)({obj})
        f.Add(obj)
        Await Task.Run(Sub() releaseobject(obj(0), f))


    End Sub
    Private Sub releaseobject(ByVal obj As Object, Optional d As List(Of Object) = Nothing)
        If d Is Nothing Then
            Try
                Runtime.InteropServices.Marshal.ReleaseComObject(obj)
                obj = Nothing
            Catch ex As Exception
                MessageBox.Show(ex.Message)
                obj = Nothing
            Finally
                GC.Collect()
            End Try
        Else
            For Each b In d
                Try
                    Runtime.InteropServices.Marshal.ReleaseComObject(b)
                    b = Nothing
                Catch ex As Exception
                    MessageBox.Show(ex.Message)
                    b = Nothing
                Finally
                    GC.Collect()
                End Try
            Next

        End If

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

        Dim mo As New AllUpd

        'Dim ds As DataRow() = РейсыКлиента(НомРес)
        Dim ds As String
        Using db As New dbAllDataContext()
            ds = db.РейсыКлиента.Where(Function(x) x.НомерРейса = НомРес).Select(Function(x) x.НазвОрганизации).FirstOrDefault()
            If ds Is Nothing Then Exit Sub
        End Using




        If Not ds = ComboBox3.Text Then 'если новый клиент
            ПровСледРейсКлиент()
            If Отмена = 1 Then Exit Sub
            ComB5()
            ComB12()

            Using db As New dbAllDataContext()
                Dim f = db.РейсыКлиента.Where(Function(x) x.НомерРейса = НомРес).Select(Function(x) x).FirstOrDefault()
                If f IsNot Nothing Then
                    With f
                        .НазвОрганизации = ComboBox3.Text
                        .КоличРейсов = СлПорРейсКл
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
                        .СтоимостьФрахта = TextBox1.Text
                        .Валюта = ComboBox5.Text
                        .ВалютаПлатежа = ComboBox8.Text
                        .СрокОплаты = TextBox4.Text
                        .УсловияОплаты = ComboBox10.Text
                        .ДогПор = ДогПор
                        .ДогПорЭксп = ДогПорЭксп
                        .ПорЭксп = ПорЭксп
                        .ДопУсловия = Trim(RichTextBox1.Text)
                        .ДатаПоручения = MaskedTextBox1.Text
                        .РазмерШтрафаЗаСрыв = ШтрафКлиент
                        .Предоплата = ЧастичнаяОплатаКлиент
                        .ОплатаПоКурсу = ОплатаПоКурсу
                    End With
                    db.SubmitChanges()
                    mo.РейсыКлиентаAllAsync()
                End If
            End Using

            ПрИзмНазКл = True
        Else 'если действующий клиент
            ComB5()
            ComB12()
            Using db As New dbAllDataContext()
                Dim f = db.РейсыКлиента.Where(Function(x) x.НомерРейса = НомРес).Select(Function(x) x).FirstOrDefault()
                If f IsNot Nothing Then
                    With f
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
                        .СтоимостьФрахта = TextBox1.Text
                        .Валюта = ComboBox5.Text
                        .ВалютаПлатежа = ComboBox8.Text
                        .СрокОплаты = TextBox4.Text
                        .УсловияОплаты = ComboBox10.Text
                        .ДогПор = ДогПор
                        .ДогПорЭксп = ДогПорЭксп
                        .ПорЭксп = ПорЭксп
                        .ДопУсловия = Trim(RichTextBox1.Text)
                        .ДатаПоручения = MaskedTextBox1.Text
                        .РазмерШтрафаЗаСрыв = ШтрафКлиент
                        .Предоплата = ЧастичнаяОплатаКлиент
                        .ОплатаПоКурсу = ОплатаПоКурсу
                    End With
                    db.SubmitChanges()
                    mo.РейсыКлиентаAllAsync()
                End If
            End Using



        End If


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


        If Not ds1.НазвОрганизации = ComboBox4.Text Then 'если новый перевозчик
            ПровСледРейсПер()
            If Отмена = 1 Then Exit Sub
            ComB13()
            Using db As New dbAllDataContext()
                Dim f = db.РейсыПеревозчика.Where(Function(x) x.НомерРейса = НомРес).Select(Function(x) x).FirstOrDefault()
                If f IsNot Nothing Then
                    With f
                        .НазвОрганизации = ComboBox4.Text
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
                        .УсловияОплаты = ComboBox9.Text
                        .ДогПор = ДогПор
                        .ДогПорЭксп = ДогПорЭксп
                        .ПорЭксп = ПорЭксп
                        .ДопУсловия = Trim(RichTextBox2.Text)
                        .ДатаПоручения = MaskedTextBox1.Text
                        .РазмерШтрафаЗаСрыв = ШтрафПер
                        .Предоплата = ЧастичнаяОплатаПеревозчик
                        .СрывЗагр20Проц = Procenty20
                    End With
                    db.SubmitChanges()
                    mo.РейсыПеревозчикаAllAsync()
                End If
            End Using

            ПрИзмНазПер = True
        Else  'если действущий перевозчик
            ComB13()

            Using db As New dbAllDataContext()
                Dim f = db.РейсыПеревозчика.Where(Function(x) x.НомерРейса = НомРес).Select(Function(x) x).FirstOrDefault()
                With f
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
                    .УсловияОплаты = ComboBox9.Text
                    .ДогПор = ДогПор
                    .ДогПорЭксп = ДогПорЭксп
                    .ПорЭксп = ПорЭксп
                    .ДопУсловия = Trim(RichTextBox2.Text)
                    .ДатаПоручения = MaskedTextBox1.Text
                    .РазмерШтрафаЗаСрыв = ШтрафПер
                    .Предоплата = ЧастичнаяОплатаПеревозчик
                    .СрывЗагр20Проц = Procenty20
                End With
                db.SubmitChanges()
                mo.РейсыПеревозчикаAllAsync()
            End Using


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
    Private Async Sub ЗапускэксельAsync(ByVal d As List(Of String))

        Try
            Await Task.Run(Sub() Запускэксель(d))
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try


    End Sub
    Private Sub Запускэксель(ByVal d As List(Of String))
        'Dim xlapp1 As Microsoft.Office.Interop.Excel.Application
        'Dim xlworkbook1 As Microsoft.Office.Interop.Excel.Workbook
        'Dim xlworksheet1 As Microsoft.Office.Interop.Excel.Worksheet
        ''Dim misvalue As Object = Reflection.Missing.Value
        'xlapp1 = New Microsoft.Office.Interop.Excel.Application With {
        '    .Visible = False
        '}
        ''xlworkbook = xlapp.Workbooks.Add(misvalue)
        'xlworkbook1 = xlapp1.Workbooks.Open(ПутьПолный,, True)

        'xlworksheet1 = xlworkbook1.Sheets(d)
        'xlworksheet1.PrintOutEx(,, 1)
        'xlworkbook1.Close(False)
        'xlapp1.Quit()

        Do While IO.File.Exists(ПутьПолный) = False

        Loop

        If plsForPrint = False Then 'это флаг что бы понять при добавлении нового рейса можно ли использовать эксель класс или новый создавать
            xlapp = New Microsoft.Office.Interop.Excel.Application

            xlworkbook = xlapp.Workbooks.Open(ПутьПолный,, True)
            For Each b In d
                xlworksheet = xlworkbook.Sheets(b)
                xlworksheet.PrintOutEx(,, 1)
            Next

            xlworkbook.Close(False)
            xlapp.Quit()
            releaseobject(xlapp)
            releaseobject(xlworkbook)
            releaseobject(xlworksheet)
        Else
            Dim xlapp1 As Microsoft.Office.Interop.Excel.Application
            Dim xlworkbook1 As Microsoft.Office.Interop.Excel.Workbook
            Dim xlworksheet1 As Microsoft.Office.Interop.Excel.Worksheet = Nothing
            'Dim misvalue As Object = Reflection.Missing.Value
            xlapp1 = New Microsoft.Office.Interop.Excel.Application With {
                .Visible = False
            }

            xlworkbook1 = xlapp1.Workbooks.Open(ПутьПолный,, True)
            For Each b In d
                xlworksheet1 = xlworkbook1.Sheets(b)
                xlworksheet1.PrintOutEx(,, 1)
            Next

            xlworkbook1.Close(False)
            xlapp1.Quit()
            releaseobject(xlapp1)
            releaseobject(xlworkbook1)
            releaseobject(xlworksheet1)
        End If


        'xlworkbook1 As Microsoft.Office.Interop.Excel.Workbook
        'Dim xlworksheet1 As Microsoft.Office.Interop.Excel.Worksheet
        ''Dim misvalue As Object = Reflection.Missing.Value
        'xlapp1 = New Microsoft.Office.Interop.Excel.Application With {
        '    .Visible = False
        '}
        'xlworkbook = xlapp.Workbooks.Add(misvalue)

        'releaseobject(xlapp1)
        'releaseobject(xlworkbook1)
        'releaseobject(xlworksheet1)
    End Sub
    Public timer2 As Timer
    Private interVal As Long = 60000
    Dim bolp As Boolean = False
    Private Sub UpdListForTime(ByVal Год As Integer)

        timer2 = New Timer(New TimerCallback(Sub() UpdListForTimeMetod(Год)), Nothing, 0, interVal)


    End Sub
    Private Sub Пер1(ByVal d As List(Of String))
        'Dim d1 = lst1all.Select(Function(x) x.ПолныйПуть).Intersect(d.Select(Function(y) y)).ToList()
        'Dim d2 = lst1all.Select(Function(x) x.ПолныйПуть).Union(d.Select(Function(y) y)).ToList()

        If d.Count > lst1all.Count Then
            Dim d3 = d.Select(Function(x) x).Except(lst1all.Select(Function(y) y.ПолныйПуть)).ToList()
            Dim m As String = Nothing
            For Each b1 In d3
                m = m & vbCrLf & IO.Path.GetFileName(b1)
            Next

            If lst1all IsNot Nothing Then
                lst1all.Clear()
            End If
            ListBox1.BeginUpdate()
            For Each b In d
                Dim mi As New ПутиДоков With {.ПолныйПуть = b, .Путь = IO.Path.GetFileName(b)}
                lst1all.Add(mi)
            Next

            ListBox1.SelectedItem = lst1all.Last
            ListBox1.EndUpdate()



            MessageBox.Show("Добавлены рейсы: " & vbCrLf & m, Рик)
        Else
            Dim d3 = lst1all.Select(Function(x) x.ПолныйПуть).Except(d.Select(Function(y) y)).ToList()
            Dim m As String = Nothing
            For Each b1 In d3
                m = m & vbCrLf & IO.Path.GetFileName(b1)
            Next

            If lst1all IsNot Nothing Then
                lst1all.Clear()
            End If
            ListBox1.BeginUpdate()
            For Each b In d
                Dim mi As New ПутиДоков With {.ПолныйПуть = b, .Путь = IO.Path.GetFileName(b)}
                lst1all.Add(mi)
            Next

            ListBox1.SelectedItem = lst1all.Last
            ListBox1.EndUpdate()



            MessageBox.Show("Удалены рейсы: " & vbCrLf & m, Рик)
        End If



    End Sub
    Private Sub UpdListForTimeMetod(ByVal Год As Integer)
        If lst1all Is Nothing Then Return
        Dim ColNow As Integer = lst1all.Count
        Dim NewCol As Integer = IO.Directory.GetFiles("Z:\RICKMANS\" & Год, "*.xls*", IO.SearchOption.TopDirectoryOnly).Length
        Dim NewCol1 = IO.Directory.GetFiles("Z:\RICKMANS\" & Год, "*.xls*", IO.SearchOption.TopDirectoryOnly).ToList()
        If Not ColNow = NewCol Then
            Try
                ListBox1.BeginInvoke(New MethodInvoker(Sub() Пер1(NewCol1)))
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try



        End If

    End Sub


    Private Sub ОбаЛистокпечататьToolStripMenuItem_Click(sender As Object, e As EventArgs)
        Dim f As New List(Of String)({"ЗАК", "ПЕР", "РЕЙС"})
        ЗапускэксельAsync(f)
        'ЗапускэксельAsync("ЗАК")
        'ЗапускэксельAsync("ПЕР")
        'ЗапускэксельAsync("РЕЙС")

    End Sub

    Private Sub ПоручениеКлиентпечататьToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles ПоручениеКлиентпечататьToolStripMenuItem1.Click
        Dim f As New List(Of String)({"ЗАК"})
        ЗапускэксельAsync(f)
        'ЗапускэксельAsync("ЗАК")
        'Dim d As New Thread(Sub() Запускэксель("ЗАК"))
        'd.IsBackground = True
        'd.Start()
    End Sub

    Private Sub ПоручениеПеревозчикпечататьToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles ПоручениеПеревозчикпечататьToolStripMenuItem1.Click
        Dim f As New List(Of String)({"ПЕР"})
        ЗапускэксельAsync(f)

        'ЗапускэксельAsync("ПЕР")

        'Dim d As New Thread(Sub() Запускэксель("ПЕР"))
        'd.IsBackground = True
        'd.Start()

    End Sub

    Private Sub ОбапечататьToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles ОбапечататьToolStripMenuItem1.Click
        'Dim d As New Thread(Sub() Запускэксель("ПЕР"))
        'd.IsBackground = True
        'd.Start()
        'Dim d1 As New Thread(Sub() Запускэксель("ЗАК"))
        'd1.IsBackground = True
        'd1.Start()

        Dim f As New List(Of String)({"ЗАК", "ПЕР"})
        ЗапускэксельAsync(f)
        'ЗапускэксельAsync("ЗАК")
        'ЗапускэксельAsync("ПЕР")
    End Sub

    Private Sub ЛистокпечататьToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles ЛистокпечататьToolStripMenuItem1.Click
        Dim f As New List(Of String)({"РЕЙС"})
        ЗапускэксельAsync(f)


        'Dim d As New Thread(Sub() Запускэксель("РЕЙС"))
        'd.IsBackground = True
        'd.Start()
    End Sub

    Private Sub ОбаЛистокпечататьToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles ОбаЛистокпечататьToolStripMenuItem1.Click
        Dim f As New List(Of String)({"ЗАК", "ПЕР", "РЕЙС"})
        ЗапускэксельAsync(f)


        'Dim d As New Thread(Sub() Запускэксель("ПЕР"))
        'd.IsBackground = True
        'd.Start()
        'Dim d1 As New Thread(Sub() Запускэксель("ЗАК"))
        'd1.IsBackground = True
        'd1.Start()
        'Dim d2 As New Thread(Sub() Запускэксель("РЕЙС"))
        'd2.IsBackground = True
        'd2.Start()
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
        'timer2.Change(Timeout.Infinite, Timeout.Infinite)  'остановка таймера
        f.Show()
        'UpdListForTime(ComboBox1.Text)  'запуск таймера
    End Sub

    Private Sub ОплатаToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ОплатаToolStripMenuItem.Click
        If НомРес = Nothing Then
            MessageBox.Show("Выберите рейс!", Рик)
            Exit Sub
        End If
        bl = False
        Dim f As New Отчет(НомРес)

        'timer2.Change(Timeout.Infinite, Timeout.Infinite)  'остановка таймера
        f.ShowDialog()
        'UpdListForTime(ComboBox1.Text)  'запуск таймера
    End Sub

    Private Sub ШтрафыToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ШтрафыToolStripMenuItem.Click
        Dim f As New Штрафы
        'timer2.Change(Timeout.Infinite, Timeout.Infinite)  'остановка таймера
        f.ShowDialog()
        'UpdListForTime(ComboBox1.Text)  'запуск таймера
        If f.ШтрафКлиент1 = True Then
            ШтрафКлиент = True
        End If

        If f.ШтрафПер1 = True Then
            ШтрафПер = True
        End If
    End Sub

    Private Sub АктСчетИРазбивкаToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles АктСчетИРазбивкаToolStripMenuItem.Click
        If ComboBox3.Text = "" Or TextBox1.Text = "" Then
            MessageBox.Show("Выберите рейс!", Рик)
            Exit Sub
        End If
        Dim f As New ДопФорма(НомРес, TextBox1.Text)
        'timer2.Change(Timeout.Infinite, Timeout.Infinite)  'остановка таймера
        f.ShowDialog()
        'UpdListForTime(ComboBox1.Text)  'запуск таймера
        If f.ОбнвлExcel = True Then
            ДокиОбновление(f.Num)
        End If
    End Sub

    Private Sub ПеревозчикToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles ПеревозчикToolStripMenuItem1.Click
        Dim f As New НовыйКлиент
        'timer2.Change(Timeout.Infinite, Timeout.Infinite)  'остановка таймера
        f.ShowDialog()
        'UpdListForTime(ComboBox1.Text)  'запуск таймера
        COM3()
    End Sub

    Private Sub ПеревозчикToolStripMenuItem2_Click(sender As Object, e As EventArgs) Handles ПеревозчикToolStripMenuItem2.Click
        Dim f As New НовыйПеревоз
        'timer2.Change(Timeout.Infinite, Timeout.Infinite)  'остановка таймера
        f.ShowDialog()
        'UpdListForTime(ComboBox1.Text)  'запуск таймера
        COM4()
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
        Dim f As New ПредоплатаКлиент
        'timer2.Change(Timeout.Infinite, Timeout.Infinite)  'остановка таймера
        f.ShowDialog()
        'UpdListForTime(ComboBox1.Text)  'запуск таймера
    End Sub

    Private Sub ToolStripMenuItem2_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem2.Click
        Dim f As New List(Of String)({"СФ"})
        ЗапускэксельAsync(f)
        'Dim d As New Thread(Sub() Запускэксель("СФ"))
        'd.IsBackground = True
        'd.Start()
    End Sub

    Private Sub ToolStripMenuItem3_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem3.Click
        Dim f As New List(Of String)({"СФ", "СФ"})
        ЗапускэксельAsync(f)
        'Dim d As New Thread(Sub() Запускэксель("СФ"))
        'd.IsBackground = True
        'd.Start()

        'Dim d1 As New Thread(Sub() Запускэксель("СФ"))
        'd1.IsBackground = True
        'd1.Start()


    End Sub

    Private Sub ToolStripMenuItem4_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem4.Click
        Dim f As New List(Of String)({"АКТ"})
        ЗапускэксельAsync(f)
        'Dim d As New Thread(Sub() Запускэксель("АКТ"))
        'd.IsBackground = True
        'd.Start()
    End Sub

    Private Sub ToolStripMenuItem5_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem5.Click
        Dim f As New List(Of String)({"АКТ", "АКТ"})
        ЗапускэксельAsync(f)


        'Dim d As New Thread(Sub() Запускэксель("АКТ"))
        'd.IsBackground = True
        'd.Start()

        'Dim d1 As New Thread(Sub() Запускэксель("АКТ"))
        'd1.IsBackground = True
        'd1.Start()
    End Sub

    Private Sub Button8_Click_1(sender As Object, e As EventArgs) Handles Button8.Click
        comdjx1 = ComboBox1.Text
        Dim f As New ПоискВРейсах(_год:=ComboBox1.Text)

        'timer2.Change(Timeout.Infinite, Timeout.Infinite)  'остановка таймера
        f.ShowDialog()
        'UpdListForTime(ComboBox1.Text)  'запуск таймера

        ComboBox1.Text = f.Годс
        ПерегрЛист1()

        If f.NumbCor IsNot Nothing Then
            If f.NumbCor.Length > 0 Then
                Dim ind = ListBox1.FindString(f.NumbCor)
                If ind > 0 Then
                    ListBox1.SetSelected(ind, True)
                    ClkLst()
                End If

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
        'timer2.Change(Timeout.Infinite, Timeout.Infinite)  'остановка таймера
        f.ShowDialog()
        'UpdListForTime(ComboBox1.Text)  'запуск таймера
    End Sub

    Private Sub ВодительToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ВодительToolStripMenuItem.Click
        Dim f As New ВодитДан
        'timer2.Change(Timeout.Infinite, Timeout.Infinite)  'остановка таймера
        f.ShowDialog()
        'UpdListForTime(ComboBox1.Text)  'запуск таймера
    End Sub

    Private Sub Условия20ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles Условия20ToolStripMenuItem.Click
        Dim f As New ДопПроц
        'timer2.Change(Timeout.Infinite, Timeout.Infinite)  'остановка таймера
        f.ShowDialog()
        'UpdListForTime(ComboBox1.Text)  'запуск таймера

    End Sub

    Private Sub ComboBox1_SelectedValueChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedValueChanged
        timer2.Change(Timeout.Infinite, Timeout.Infinite)  'остановка таймера

        ПерегрЛист1(, ComboBox1.Text)


        UpdListForTime(ComboBox1.Text)  'запуск таймера
    End Sub

    Private Sub ПоискОбщийToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ПоискОбщийToolStripMenuItem.Click
        Dim f As New ПоискПолный
        'timer2.Change(Timeout.Infinite, Timeout.Infinite)  'остановка таймера
        f.ShowDialog()
        'UpdListForTime(ComboBox1.Text)  'запуск таймера
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs)
        Dim f As New Штрафы
        'timer2.Change(Timeout.Infinite, Timeout.Infinite)  'остановка таймера
        f.ShowDialog()
        'UpdListForTime(ComboBox1.Text)  'запуск таймера

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
        Dim mo As New AllUpd


        Dim xlapp As Microsoft.Office.Interop.Excel.Application
        xlapp = New Microsoft.Office.Interop.Excel.Application

        Dim XXX = "Provider='SQLOLEDB';Data Source=178.124.211.175,52891;Network Library=DBMSSOCN;Initial Catalog=Rickmans;User ID=Rickmans;Password=Zf6VpP37Ol"

        Dim CON As New ADODB.Connection
        Dim RS As New ADODB.Recordset
        Dim RS1 As New ADODB.Recordset
        Dim RS2 As New ADODB.Recordset
        Dim RS3 As New ADODB.Recordset

        Dim com3 As String = ComboBox3.Text
        Dim com4 As String = ComboBox4.Text

        Dim strSQL, strSQL1, strSQL2, strSQL3 As String
        strSQL = "select * from РейсыКлиента WHERE НомерРейса=" & СлРейс & ""
        strSQL1 = "select * from РейсыПеревозчика WHERE НомерРейса=" & СлРейс & ""
        strSQL2 = "select * from Клиент WHERE НазваниеОрганизации='" & com3 & "'"
        strSQL3 = "select * from Перевозчики WHERE Названиеорганизации='" & com4 & "'"

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


        Dim КоличРейсЗак As Integer
        Dim КоличРейсПер As Integer
        Using db As New dbAllDataContext()
            КоличРейсЗак = db.РейсыКлиента.Where(Function(x) x.НомерРейса = СлРейс).Select(Function(x) x.КоличРейсов).FirstOrDefault()
            КоличРейсПер = db.РейсыПеревозчика.Where(Function(x) x.НомерРейса = СлРейс).Select(Function(x) x.КоличРейсов).FirstOrDefault()
        End Using

        If КоличРейсЗак = 0 Then
            Throw New System.Exception("Нет данных!.'ДокиОбновление'")
            Return
        End If
        If КоличРейсПер = 0 Then
            Throw New System.Exception("Нет данных!.'ДокиОбновление'")
            Return
        End If

        Dim com1 As String = ComboBox1.Text



        If ПрИзмНазКл = True Then
            Dim G As String = ПутьРейса

            Try
                IO.File.Copy("Z:\RICKMANS\" & com1 & "\" & ПутьРейса, "Z:\RICKMANS\" & com1 & "\" & НомРес & " " & com3 & " " & КоличРейсЗак & " - " & com4 & " " & КоличРейсПер & ".xlsm", True)
            Catch ex As Exception
                IO.File.Delete("Z:\RICKMANS\" & com1 & "\" & НомРес & " " & com3 & " " & КоличРейсЗак & " - " & com4 & " " & КоличРейсПер & ".xlsm")
                IO.File.Copy("Z:\RICKMANS\" & com1 & "\" & ПутьРейса, "Z:\RICKMANS\" & com1 & "\" & НомРес & " " & com3 & " " & КоличРейсЗак & " - " & com4 & " " & КоличРейсПер & ".xlsm", True)
            End Try

            IO.File.Delete("Z:\RICKMANS\" & com1 & "\" & G)
            ПрИзмНазКл = False
            ПутьРейса = НомРес & " " & com3 & " " & КоличРейсЗак & " - " & com4 & " " & КоличРейсПер & ".xlsm"
        End If

        If ПрИзмНазПер = True Then
            Dim G As String = ПутьРейса
            Try
                IO.File.Copy("Z:\RICKMANS\" & com1 & "\" & ПутьРейса, "Z:\RICKMANS\" & com1 & "\" & НомРес & " " & com3 & " " & КоличРейсЗак & " - " & com4 & " " & КоличРейсПер & ".xlsm", True)
            Catch ex As Exception
                IO.File.Delete("Z:\RICKMANS\" & com1 & "\" & НомРес & " " & com3 & " " & КоличРейсЗак & " - " & com4 & " " & КоличРейсПер & ".xlsm")
                IO.File.Copy("Z:\RICKMANS\" & com1 & "\" & ПутьРейса, "Z:\RICKMANS\" & com1 & "\" & НомРес & " " & com3 & " " & КоличРейсЗак & " - " & com4 & " " & КоличРейсПер & ".xlsm", True)
            End Try

            IO.File.Delete("Z:\RICKMANS\" & com1 & "\" & G)
            ПрИзмНазПер = False
            ПутьРейса = НомРес & " " & com3 & " " & КоличРейсЗак & " - " & com4 & " " & КоличРейсПер & ".xlsm"
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



    Private Sub Button6_Click(sender As Object, e As EventArgs)

        If MessageBox.Show("Удалить " & НомРес & " Рейс?", Рик, MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.No Then
            Exit Sub
        End If

        Diction()

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
        If IO.File.Exists("Z:\RICKMANS\" & ComboBox1.Text & "\" & ПутьРейса) Then
            IO.File.Copy("Z:\RICKMANS\" & ComboBox1.Text & "\" & ПутьРейса, "Z:\RICKMANS\" & ComboBox1.Text & "\СОРВАННЫЕ ЗАГРУЗКИ\" & ПутьРейса, True)
            IO.File.Delete("Z:\RICKMANS\" & ComboBox1.Text & "\" & ПутьРейса)
            MessageBox.Show("Рейс полностью удалён!", Рик)
        Else
            MessageBox.Show("Рейс не найден!", Рик)
        End If

        ClearDictionAsync()


        Очистка()
        ПерегрЛист1()
    End Sub
    Private Async Sub PredzagAsync()
        Await Task.Run(Sub() Predzag())
    End Sub

    Private Sub ИзменитьНомерРейсаToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles ИзменитьНомерРейсаToolStripMenuItem1.Click
        Dim f As New ИзменитьНомерРейса(НомРес, lst1all, list1selПуть)
        f.ShowDialog()
        Dim k = f.newPt
        If lst1all Is Nothing Then Return
        If k Is Nothing Then Return
        lst1all.Remove(list1selПуть) 'удаляем старый путь
        lst1all.Add(k) 'добавляем новый путь
        Dim f2 = lst1all.OrderBy(Function(x) x.Путь).ToList()
        lst1all.Clear()

        ListBox1.BeginUpdate()
        For Each b In f2
            lst1all.Add(b)
        Next
        ListBox1.SelectedItem = lst1all.Last
        ListBox1.EndUpdate()
    End Sub

    Private Sub УдалитьToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles УдалитьToolStripMenuItem.Click
        If MessageBox.Show("Удалить " & НомРес & " Рейс?", Рик, MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.No Then
            Exit Sub
        End If

        Diction()

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
        If IO.File.Exists("Z:\RICKMANS\" & ComboBox1.Text & "\" & ПутьРейса) Then
            IO.File.Copy("Z:\RICKMANS\" & ComboBox1.Text & "\" & ПутьРейса, "Z:\RICKMANS\" & ComboBox1.Text & "\СОРВАННЫЕ ЗАГРУЗКИ\" & ПутьРейса, True)
            IO.File.Delete("Z:\RICKMANS\" & ComboBox1.Text & "\" & ПутьРейса)
            MessageBox.Show("Рейс полностью удалён!", Рик)
        Else
            MessageBox.Show("Рейс не найден!", Рик)
        End If
        ClearDictionAsync()
        Очистка()
        ПерегрЛист1()
    End Sub

    Private Sub СоздатьНовыйToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles СоздатьНовыйToolStripMenuItem.Click

        PredzagAsync()
        Diction()
        Отмена = 0
        ПерезагрЛист1 = 0
        If Проверка() = 1 Then Exit Sub
        ПредExcelAsync()
        Me.Cursor = Cursors.WaitCursor
        If НомРес > 0 Then
            If MessageBox.Show("Создать рейс?", Рик, MessageBoxButtons.YesNo) = DialogResult.Yes Then


                НовыйРейсГлавная()
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

    Private Sub ИзменитьДействующийToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ИзменитьДействующийToolStripMenuItem.Click
        PredzagAsync()
        Diction()
        Отмена = 0
        ПерезагрЛист1 = 0
        If Проверка() = 1 Then Exit Sub
        Me.Cursor = Cursors.WaitCursor

        If НомРес > 0 Then
            If MessageBox.Show("Изменить данные этого рейса?", Рик, MessageBoxButtons.YesNo) = DialogResult.Yes Then
                РедакцияСтарогоРейса()
                plsForPrint = False
                If Отмена = 1 Then
                    Me.Cursor = Cursors.Default
                    Exit Sub
                End If
                MessageBox.Show("Рейс изменен!", Рик)
                'ПерегрДанныхИзБазы()
                ПерегрЛист1()
            End If
        End If

        If НомРес = Nothing Then
            ПерезагрЛист1 = 1
            НовыйРейсГлавная()
            Очистка()
            ПерезагрЛист1 = 0
        End If
        ClearDictionAsync()
        Button2.BackColor = Color.LightBlue
        Button7.BackColor = Color.LightBlue
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub Predzag()
        Dim mo As New AllUpd
        Do While AllClass.РейсыКлиента Is Nothing
            mo.РейсыКлиентаAll()
        Loop
        Do While AllClass.РейсыПеревозчика Is Nothing
            mo.РейсыПеревозчикаAll()
        Loop

        Do While AllClass.Перевозчики Is Nothing
            mo.ПеревозчикиAll()
        Loop


    End Sub
    Private Async Sub DictionAsync()
        Await Task.Run(Sub() Diction())
    End Sub

    Private Sub Label14_MouseMove(sender As Object, e As MouseEventArgs) Handles Label14.MouseMove

        If NewListDynamFlag = True Then Return


        listbxDyn = New ListBox With {.Size = New Size(150, 200), .Location = New Point(1023, 77)}
        '.Location = New Point(100, 77),  '1023, 77
        'f.Left = 1023 : f.Top = 20

        listbxDyn.MultiColumn = True
        listbxDyn.Visible = True
        Me.Controls.Add(listbxDyn)
        listbxDyn.BringToFront()
        ResumeLayout()

        Dim год = ComboBox1.Text
        Dim f1 = (IO.Directory.GetFiles("Z:\RICKMANS\" & год, "*.xls*", IO.SearchOption.TopDirectoryOnly)).ToList()
        Dim f2 As New List(Of Integer)
        For Each b In f1
            f2.Add(Strings.Left(IO.Path.GetFileName(b), 3))
        Next
        Dim f3 = f2.OrderBy(Function(x) x).ToList()
        Dim fs = f3.FirstOrDefault
        Dim ls = f3.LastOrDefault
        Dim col = ls - fs
        Dim f4 = Enumerable.Range(fs, col).Except(f2).ToList()
        listbxDyn.BeginUpdate()

        If f4 IsNot Nothing Then
            For Each b In f4
                listbxDyn.Items.Add(b)
            Next
        End If
        listbxDyn.EndUpdate()

        NewListDynamFlag = True


    End Sub

    Private Sub Label14_MouseLeave(sender As Object, e As EventArgs) Handles Label14.MouseLeave

        If ClicnewLstDyn = False Then
            Me.Controls.Remove(listbxDyn)
            NewListDynamFlag = False
        End If


    End Sub

    Private Sub Label14_MouseClick(sender As Object, e As MouseEventArgs) Handles Label14.MouseClick
        Dim f = e.Button
        If f = MouseButtons.Left Then
            If ClicnewLstDyn = False Then
                ClicnewLstDyn = True
            Else
                ClicnewLstDyn = False
            End If
        End If
    End Sub

    Private Sub Diction()


        If arrtbox.Any Then
            arrtbox.Clear()
        End If

        If arrtmask.Any Then
            arrtmask.Clear()
        End If

        If arrtcom.Any Then
            arrtcom.Clear()
        End If

        If arrtRichbox.Any Then
            arrtRichbox.Clear()
        End If



        For Each gh In Me.Controls.OfType(Of GroupBox)  'перебираем в groupbox
            For Each tx In gh.Controls.OfType(Of TextBox)
                arrtbox.Add(tx.Name, tx.Text)
            Next
            For Each tx In gh.Controls.OfType(Of ComboBox)
                arrtcom.Add(tx.Name, tx.Text)
            Next

            For Each ts In gh.Controls.OfType(Of MaskedTextBox)
                arrtmask.Add(ts.Name, ts.Text)
            Next

            For Each ts In gh.Controls.OfType(Of RichTextBox)
                arrtRichbox.Add(ts.Name, ts.Text)
            Next


        Next

        Dim Ctrl As Control


        For Each Ctrl In Me.Controls 'перебираем текстбоксы вне  groupbox
            If TypeName(Ctrl) = "TextBox" Then
                arrtbox.Add(Ctrl.Name, Ctrl.Text)
                'Ctrl.Value = "бла-бла-бла"
            End If

            If TypeName(Ctrl) = "ComboBox" Then
                arrtcom.Add(Ctrl.Name, Ctrl.Text)
                'Ctrl.Value = "бла-бла-бла"
            End If

            If TypeName(Ctrl) = "RichTextBox" Then
                arrtRichbox.Add(Ctrl.Name, Ctrl.Text)
                'Ctrl.Value = "бла-бла-бла"
            End If

            If TypeName(Ctrl) = "MaskedTextBox" Then
                arrtmask.Add(Ctrl.Name, Ctrl.Text)
                'Ctrl.Value = "бла-бла-бла"
            End If

        Next



    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        PredzagAsync()

        Отмена = 0
        ПерезагрЛист1 = 0
        If Проверка() = 1 Then Exit Sub
        Me.Cursor = Cursors.WaitCursor
        If НомРес > 0 Then
            Dim Res As DialogResult = MessageBox.Show("Выберите 'Да'- если, хотите заменить данные этого рейса" & vbCrLf & " Выберите 'Нет'- если, хотите создать новый рейс" & vbCrLf & "Выберите 'Отмена'- если, хотите выйти", Рик, MessageBoxButtons.YesNoCancel)
            If Res = DialogResult.No Then
                Diction()
                ПредExcelAsync()
                НовыйРейсГлавная()
            End If
            If Res = DialogResult.Yes Then
                Diction()
                РедакцияСтарогоРейса()
                plsForPrint = False
                If Отмена = 1 Then
                    Me.Cursor = Cursors.Default
                    Exit Sub
                End If
                ClearDiction()
                MessageBox.Show("Рейс изменен!", Рик)
                'ПерегрДанныхИзБазы()
                ПерегрЛист1()
            End If
            If Res = DialogResult.Cancel Then
                Me.Cursor = Cursors.Default
                'УничтожениеExcelAsync()
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
    Private Async Sub ClearDictionAsync()
        Await Task.Run(Sub() ClearDiction())
    End Sub
    Private Sub ClearDiction()
        If arrtbox.Any Then
            arrtbox.Clear()
        End If

        If arrtmask.Any Then
            arrtmask.Clear()
        End If

        If arrtcom.Any Then
            arrtcom.Clear()
        End If

        If arrtRichbox.Any Then
            arrtRichbox.Clear()
        End If
    End Sub
    Private Async Sub УничтожениеExcelAsync()
        Await Task.Run(Sub() УничтожениеExcel())
    End Sub
    Private Sub УничтожениеExcel()
        Do While pls = False

        Loop
        Dim f As New List(Of Object)({xlapp, xlworkbook, xlworksheet})
        releaseobject(xlapp, f)
    End Sub

    Private Sub НовыйРейсГлавная()
        НовыйРейс()
        If Отмена = 1 Then Exit Sub

        Доки()
        'ПредExcelAsync() 'создаем новый обьект эксель
        'ПерегрДанныхИзБазы()
        MessageBox.Show("Рейс оформлен!", Рик)
        If ПерезагрЛист1 = 0 Then
            ПерегрЛист1(NewPutForListInNewRejs)
        End If

    End Sub

    Private Sub Рейс_Closed(sender As Object, e As EventArgs) Handles Me.Closed
        timer2.Change(Timeout.Infinite, Timeout.Infinite)  'остановка таймера
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