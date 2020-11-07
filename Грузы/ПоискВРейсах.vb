Imports System.Threading
'Imports Microsoft.Office.Interop.Excel
Imports System.IO
Imports System.Data.Linq.SqlClient

Public Class ПоискВРейсах
    Private Delegate Sub comb11()
    Private Delegate Sub comb12()
    Dim ГотовыеДанные As New List(Of String)
    Dim СписокПапок() As String
    Dim fl4 As New List(Of String)()
    Dim flPath4 As New List(Of String)()
    Dim FilesПолнПуть2() As String
    Dim Files45() As String
    Private Год As String
    Private lst1 As List(Of Listbx1Class)
    Private lst1sort As List(Of Listbx1Class)
    Private bslst1 As BindingSource
    Public NumbCor As String
    Sub New(ByVal _год As String)
        Год = _год
        ' Этот вызов является обязательным для конструктора.
        InitializeComponent()

        ' Добавить код инициализации после вызова InitializeComponent().

    End Sub
    Public Class Listbx1Class
        Public Property КорПуть As String
        Public Property ПолнПуть As String
        Public Property Клиент As String
        Public Property КлиентТелефон As String
        Public Property КлиентСтавка As String
        Public Property Перевозчик As String
        Public Property ПеревозчикТелефон As String
        Public Property ПеревозчикСтавка As String
        Public Property Маршрут As String
        Public Property ДатаЗагрузки As String
        Public Property Местозагрузки As String
        Public Property Автомобиль As String
        Public Property Водитель As String
        Public Property ДатаВыгрузки As String
        Public Property Местовыгрузки As String



    End Class
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
        Do While AllClass.Клиент Is Nothing
            mo.КлиентAll()
        Loop
        Do While AllClass.Перевозчики Is Nothing
            mo.ПеревозчикиAll()
        Loop
    End Sub

    Private Sub ПоискВРейсах_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Cursor = Cursors.WaitCursor
        ПредзагрузкаAsync()
        lst1 = New List(Of Listbx1Class)
        lst1sort = New List(Of Listbx1Class)
        bslst1 = New BindingSource
        bslst1.DataSource = lst1sort
        ListBox1.DataSource = bslst1
        ListBox1.DisplayMember = "КорПуть"

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



        Dim f = Directory.GetFiles("Z:\RICKMANS\" & Год, "*.xls*", IO.SearchOption.AllDirectories).OrderBy(Function(x) x).ToList()
        Dim f3 = f.Select(Function(x) Strings.Left(Path.GetFileName(x), 3))

        If f IsNot Nothing Then
            For Each b In f
                Dim f4 As New РейсыПеревозчика
                Dim f5 As New РейсыКлиента
                Dim f41 As New Перевозчики
                Dim f51 As New Клиент
                Dim inu As String = Nothing

                inu = Strings.Left(Path.GetFileName(b), 3)
                If inu Is Nothing Then Continue For

                f4 = AllClass.РейсыПеревозчика.Where(Function(x) x.НомерРейса = inu).Select(Function(x) x).FirstOrDefault()
                If f4 Is Nothing Then Continue For

                f5 = AllClass.РейсыКлиента.Where(Function(x) x.НомерРейса IsNot Nothing And x.НомерРейса = inu).Select(Function(x) x).FirstOrDefault()
                If f5 Is Nothing Then Continue For

                f41 = AllClass.Перевозчики.Where(Function(x) x.Названиеорганизации IsNot Nothing And x.Названиеорганизации = f4.НазвОрганизации).Select(Function(x) x).FirstOrDefault()
                f51 = AllClass.Клиент.Where(Function(x) x.НазваниеОрганизации = f5.НазвОрганизации).Select(Function(x) x).FirstOrDefault()

                Dim f2 As New Listbx1Class With {.ПолнПуть = b, .КорПуть = Path.GetFileName(b), .Автомобиль = f4.НомерАвтомобиля,
                                      .Водитель = f4.Водитель, .ДатаВыгрузки = f4.ДатаПодачиПодРастаможку, .ДатаЗагрузки = f4.ДатаПодачиПодЗагрузку,
                                      .Маршрут = f4.Маршрут, .Местовыгрузки = f4.ТочнАдресРазгр, .Местозагрузки = f4.ТочныйАдресЗагрузки, .Перевозчик = f4.НазвОрганизации,
                                      .ПеревозчикСтавка = f4.СтоимостьФрахта, .ПеревозчикТелефон = f41.Контактное_лицо & " - " & f41.Телефон, .Клиент = f5.НазвОрганизации, .КлиентСтавка = f5.СтоимостьФрахта,
                                      .КлиентТелефон = f51.Контактное_лицо & " - " & f51.Телефон}
                lst1.Add(f2)
            Next



        End If

        Me.Cursor = Cursors.Default
    End Sub
    'Private Sub list1sbor()


    '    Dim gth3 As String
    '    Dim f As String = Год

    '    FilesПолнПуть2 = (IO.Directory.GetFiles("Z:\RICKMANS\" & Год, "*.xls*", IO.SearchOption.TopDirectoryOnly))
    '    Files45 = (IO.Directory.GetFiles("Z:\RICKMANS\" & Год, "*.xls*", IO.SearchOption.TopDirectoryOnly))
    '    For n As Integer = 0 To Files45.Length - 1
    '        gth3 = ""
    '        gth3 = IO.Path.GetFileName(Files45(n))
    '        Files45(n) = gth3
    '    Next
    '    ListBox1.Items.AddRange(Files45)
    '    Label14.Text = Files45.Length

    'End Sub

    Private Sub lab14(ByVal d As Integer)
        If Label14.InvokeRequired Then
            Me.Invoke(New comb12(Sub() lab14(d)))
        Else
            Label14.Text = d
        End If

    End Sub
    Private Sub Listdelegate(ByVal d As List(Of String))
        If ListBox1.InvokeRequired Then
            Me.Invoke(New comb11(Sub() Listdelegate(d)))
        Else
            ListBox1.Items.Clear()
            For m As Integer = 0 To d.Count - 1
                If Not d(m) = Nothing Then
                    ListBox1.Items.Add(d(m))
                End If
            Next
            'ListBox1.Text = Files3.Last
        End If
    End Sub
    Private Sub ПоискВБазе(ByVal d As String, ByVal m As Integer)

        Dim mo As New AllUpd
        Do While AllClass.РейсыКлиента Is Nothing
            mo.РейсыКлиентаAll()
        Loop
        Do While AllClass.РейсыПеревозчика Is Nothing
            mo.РейсыПеревозчикаAll()
        Loop


        Dim np As New List(Of РейсыПеревозчика)
        Dim nk As New List(Of РейсыКлиента)

        'пересобираем коллекцию и убираем пустые значения
        For Each b In AllClass.РейсыПеревозчика
            If b.Маршрут Is Nothing Then
                Continue For
            Else
                np.Add(b)
            End If
        Next
        For Each b In AllClass.РейсыКлиента
            If b.Маршрут Is Nothing Then
                Continue For
            Else
                nk.Add(b)
            End If
        Next

        If lst1sort IsNot Nothing Then
            lst1sort.Clear()
        End If






        If m = 1 Then

            Dim txt1 As String = TextBox1.Text
            Dim f As New List(Of РейсыКлиента)
            For Each b In nk
                If b.Маршрут.ToUpper.Contains(txt1.ToUpper) Then
                    f.Add(b)
                End If
            Next

            Dim f2 = lst1.Select(Function(b) CType(Strings.Left(b.КорПуть, 3), Integer)).Intersect(f.Select(Function(x) CType(x.НомерРейса, Integer)))
            For Each b In f2
                Dim f5 = lst1.Where(Function(x) CType(Strings.Left(x.КорПуть, 3), Integer) = b).Select(Function(x) x).FirstOrDefault()
                lst1sort.Add(f5)
            Next

            bslst1.ResetBindings(False)


        ElseIf m = 2 Then

            Dim txt1 As String = TextBox2.Text
            Dim f As New List(Of РейсыКлиента)
            For Each b In nk
                If b.НазвОрганизации.ToUpper.Contains(txt1.ToUpper) Then
                    f.Add(b)
                End If
            Next

            Dim f2 = lst1.Select(Function(b) CType(Strings.Left(b.КорПуть, 3), Integer)).Intersect(f.Select(Function(x) CType(x.НомерРейса, Integer)))

            For Each b In f2
                Dim f5 = lst1.Where(Function(x) CType(Strings.Left(x.КорПуть, 3), Integer) = b).Select(Function(x) x).FirstOrDefault()
                lst1sort.Add(f5)
            Next

            bslst1.ResetBindings(False)

        ElseIf m = 3 Then

            Dim txt1 As String = TextBox3.Text
            Dim f As New List(Of РейсыКлиента)
            For Each b In nk
                If b.ТочныйАдресЗагрузки.ToUpper.Contains(txt1.ToUpper) Then
                    f.Add(b)
                End If
            Next

            Dim f2 = lst1.Select(Function(b) CType(Strings.Left(b.КорПуть, 3), Integer)).Intersect(f.Select(Function(x) CType(x.НомерРейса, Integer)))

            For Each b In f2
                Dim f5 = lst1.Where(Function(x) CType(Strings.Left(x.КорПуть, 3), Integer) = b).Select(Function(x) x).FirstOrDefault()
                lst1sort.Add(f5)
            Next

            bslst1.ResetBindings(False)

        Else

            Dim txt1 As String = TextBox4.Text
            Dim f As New List(Of РейсыПеревозчика)
            For Each b In np
                If b.НазвОрганизации.ToUpper.Contains(txt1.ToUpper) Then
                    f.Add(b)
                End If
            Next

            Dim f2 = lst1.Select(Function(b) CType(Strings.Left(b.КорПуть, 3), Integer)).Intersect(f.Select(Function(x) CType(x.НомерРейса, Integer)))

            For Each b In f2
                Dim f5 = lst1.Where(Function(x) CType(Strings.Left(x.КорПуть, 3), Integer) = b).Select(Function(x) x).FirstOrDefault()
                lst1sort.Add(f5)
            Next

            bslst1.ResetBindings(False)

        End If

        Label14.Text = ListBox1.Items.Count
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        TextBox3.Text = ""
        TextBox2.Text = ""
        'list1sbor()

        ПоискВБазе("Маршрут", 1)
        ЧистимРичи()
    End Sub
    Private Sub ЧистимРичи()
        For Each tx1 In Controls.OfType(Of GroupBox)
            For Each tx In tx1.Controls.OfType(Of RichTextBox)
                tx.Text = ""
            Next
        Next
        For Each tx1 In Controls.OfType(Of GroupBox)
            For Each tx2 In tx1.Controls.OfType(Of GroupBox)
                For Each tx In tx2.Controls.OfType(Of RichTextBox)
                    tx.Text = ""
                Next
            Next
        Next
    End Sub
    Private Sub ПоискПомаршрутуВариант1()
        Dim _Excel As New Microsoft.Office.Interop.Excel.Application
        Dim Лист As Microsoft.Office.Interop.Excel.Worksheet
        Dim Книга As Microsoft.Office.Interop.Excel.Workbook
        _Excel.Visible = False

        For x As Integer = 0 To flPath4.Count - 1
            Try
                Книга = _Excel.Workbooks.Open(flPath4(x)) 'Открываем книгу
                Лист = CType(Книга.Worksheets("ЗАК"), Microsoft.Office.Interop.Excel.Worksheet)
                If Лист.Range("C7").Value.ToString.Contains(TextBox1.Text) Then
                    ГотовыеДанные.Add(flPath4(x))
                End If
                _Excel.Workbooks(1).Close(False)
                _Excel.Quit()
            Catch ex As Exception

            End Try

        Next

        ' сохранение и закрытие
        _Excel.Quit()

        'Первый лист книги
        TextBox1.Text = Лист.Range("B5").Value
    End Sub
    Private Sub ПоискПоМаршруту()
        Dim xlapp1 As Application
        Dim xlworkbook1 As Microsoft.Office.Interop.Excel.Workbook
        Dim xlworksheet1 As Microsoft.Office.Interop.Excel.Worksheet
        'Dim misvalue As Object = Reflection.Missing.Value





        'For x As Integer = 0 To flPath4.Count - 1
        '    Try
        '        xlapp1 = New Microsoft.Office.Interop.Excel.Application With {
        '    .Visible = False
        '}
        '        xlworkbook1 = xlapp1.Microsoft.Office.Interop.Excel.Workbooks.Open(flPath4(x),, True) 'Открываем книгу
        '        xlworksheet1 = CType(xlapp1.Worksheets("ЗАК"), Worksheet)
        '        If xlworksheet1.Range("C7").Value.ToString.Contains(TextBox1.Text) Then
        '            ГотовыеДанные.Add(flPath4(x))
        '        End If
        '        xlapp1.Quit()
        '        releaseobject(xlapp1)
        '    Catch ex As Exception

        '    End Try

        'Next
        ListBox1.Items.Clear()
        For n As Integer = 0 To ГотовыеДанные.Count - 1
            ListBox1.Items.Add(ГотовыеДанные(n))
        Next
    End Sub
    Private Sub ВесьСписок()
        Dim f As Integer = CType(Strings.Left(ListBox1.SelectedItem.ToString, 3), Integer)
        Dim ds As DataTable = Selects3(StrSql:="SELECT * FROM РейсыКлиента WHERE НомерРейса=" & f & "")
        Dim ds1 As DataTable = Selects3(StrSql:="SELECT * FROM РейсыПеревозчика WHERE НомерРейса=" & f & "")
        Dim ds2 As DataTable = Selects3(StrSql:="SELECT [Контактное лицо],Телефон FROM Клиент WHERE НазваниеОрганизации='" & ds.Rows(0).Item(1).ToString & "'")
        Dim ds3 As DataTable = Selects3(StrSql:="SELECT [Контактное лицо],Телефон FROM Перевозчики WHERE Названиеорганизации='" & ds1.Rows(0).Item(1).ToString & "'")

        RichTextBox1.Text = ds.Rows(0).Item(1).ToString
        RichTextBox2.Text = ds1.Rows(0).Item(1).ToString
        RichTextBox3.Text = ds.Rows(0).Item(4).ToString
        RichTextBox4.Text = ds.Rows(0).Item(5).ToString
        RichTextBox5.Text = ds.Rows(0).Item(9).ToString
        RichTextBox6.Text = ds.Rows(0).Item(12).ToString
        RichTextBox7.Text = ds.Rows(0).Item(14).ToString
        RichTextBox8.Text = ds.Rows(0).Item(7).ToString
        RichTextBox9.Text = ds.Rows(0).Item(16).ToString
        RichTextBox10.Text = ds2.Rows(0).Item(0).ToString & " (" & ds2.Rows(0).Item(1).ToString & ")"
        RichTextBox11.Text = ds3.Rows(0).Item(0).ToString & " (" & ds3.Rows(0).Item(1).ToString & ")"
        RichTextBox12.Text = ds.Rows(0).Item(17).ToString
        RichTextBox13.Text = ds1.Rows(0).Item(17).ToString
    End Sub

    Private Sub ListBox1_MouseClick(sender As Object, e As MouseEventArgs) Handles ListBox1.MouseClick
        For Each tx1 In Controls.OfType(Of GroupBox)
            For Each tx In tx1.Controls.OfType(Of RichTextBox)
                tx.Text = String.Empty
            Next
        Next

        'If TextBox1.Text = "" And TextBox2.Text = "" And TextBox3.Text = "" And TextBox4.Text = "" Then
        '    ВесьСписок()
        'End If
        If ListBox1.SelectedIndex = -1 Then Return

        Dim f1 = lst1sort.ElementAt(ListBox1.SelectedIndex)



        RichTextBox1.Text = f1.Клиент
        RichTextBox2.Text = f1.Перевозчик
        RichTextBox3.Text = f1.Маршрут
        RichTextBox4.Text = f1.ДатаЗагрузки
        RichTextBox5.Text = f1.Местозагрузки
        RichTextBox6.Text = f1.Автомобиль
        RichTextBox7.Text = f1.Водитель
        RichTextBox8.Text = f1.ДатаВыгрузки
        RichTextBox9.Text = f1.Местовыгрузки
        RichTextBox10.Text = f1.КлиентТелефон
        RichTextBox11.Text = f1.ПеревозчикТелефон
        RichTextBox12.Text = f1.КлиентСтавка
        RichTextBox13.Text = f1.ПеревозчикСтавка
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        TextBox1.Text = ""
        TextBox2.Text = ""

        'list1sbor()
        ПоискВБазе("ТочныйАдресЗагрузки", 3)
        ЧистимРичи()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        TextBox1.Text = ""
        TextBox3.Text = ""

        'list1sbor()
        ПоискВБазе("НазвОрганизации", 2)
        ЧистимРичи()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        TextBox1.Text = ""
        TextBox3.Text = ""
        TextBox2.Text = ""
        'list1sbor()
        ПоискВБазе("НазвОрганизации", 4)
        ЧистимРичи()
    End Sub

    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown
        If e.KeyCode = Keys.Enter Then
            Button8.PerformClick()
        End If
    End Sub

    Private Sub TextBox3_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox3.KeyDown
        If e.KeyCode = Keys.Enter Then
            Button2.PerformClick()
        End If
    End Sub

    Private Sub TextBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox2.KeyDown
        If e.KeyCode = Keys.Enter Then
            Button1.PerformClick()
        End If
    End Sub

    Private Sub TextBox4_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox4.KeyDown
        If e.KeyCode = Keys.Enter Then
            Button3.PerformClick()
        End If
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        If ListBox1.SelectedIndex = -1 Then
            MessageBox.Show("Выберите рейс!", Рик)
            Return
        End If

        Dim f = lst1sort.ElementAt(ListBox1.SelectedIndex)
        NumbCor = f.КорПуть
        Close()
    End Sub
End Class