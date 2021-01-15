Public Class ВыборкаЖурнал

    Private com1all As New List(Of IDNaz)
    Private bscom1 As New BindingSource

    Private com2all As New List(Of IDNaz)
    Private bscom2 As New BindingSource

    Private Grid1all As New List(Of Grid1Class)
    Private bsGrid1 As New BindingSource

    Private DataAll As New List(Of СтраныClass)
    Sub New()

        ' Этот вызов является обязательным для конструктора.
        InitializeComponent()
        ПредзагрузкаAsync()
        ' Добавить код инициализации после вызова InitializeComponent().

    End Sub
    Private Sub ВыборкаЖурнал_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        bscom1.DataSource = com1all
        ComboBox1.DataSource = bscom1
        ComboBox1.DisplayMember = "Naz"

        bscom2.DataSource = com2all
        ComboBox2.DataSource = bscom2
        ComboBox2.DisplayMember = "Naz"

        bsGrid1.DataSource = Grid1all
        Grid1.DataSource = bsGrid1
        GridView(Grid1)
        Grid1.Columns(0).Width = 60
        Grid1.Columns(2).Width = 80
        Grid1.Columns(4).Width = 100
        Grid1.Columns(8).Visible = False

        LoadComboAsync()

    End Sub
    Private Async Sub ПредзагрузкаAsync()
        Await Task.Run(Sub() Предзагрузка())
    End Sub
    Private Sub Предзагрузка()
        Dim mo As New AllUpd

        Do While AllClass.Страна Is Nothing
            mo.СтранаAll()
        Loop

        Do While AllClass.ЖурналДата Is Nothing
            mo.ЖурналДатаAll()
        Loop

        Do While AllClass.ЖурналКлиентГруз Is Nothing
            mo.ЖурналКлиентГрузAll()
        Loop

        Do While AllClass.ЖурналКлиентМаршрут Is Nothing
            mo.ЖурналКлиентМаршрутAll()
        Loop

        Do While AllClass.SkypeКлиентПредложение Is Nothing
            mo.SkypeКлиентПредложениеAll()
        Loop

        Do While AllClass.SkypeПеревозчикПредложение Is Nothing
            mo.SkypeПеревозчикПредложениеAll()
        Loop

        Do While AllClass.Клиент Is Nothing
            mo.КлиентAll()
        Loop

        ПодготовкаАктуальныхДанных()
    End Sub
    Private Sub ПодготовкаВсехДанных()


        Dim f = (From b In (From x In AllClass.ЖурналКлиентГруз   'выборка из таблицы журнал грузы 
                            Join z In AllClass.ЖурналДата On z.Код Equals x.КодЖурналДата
                            Join y In AllClass.ЖурналКлиентМаршрут On x.Код Equals y.КодЖурналКлиентГруз
                            Order By x.ДатаЗагрузки Descending
                            Where x.Экспедитор = Экспедитор
                            Select x, y, z).ToList()
                 Group b By Keys = New With {Key b.y.СтранаПогрузки, Key b.y.СтранаВыгрузки}
                      Into Group
                 Select New СтраныClass With {.СтранаЗагрузка = Keys.СтранаПогрузки, .СтранаВыгрузки = Keys.СтранаВыгрузки, .ТипВыборки = "Груз", .lst = (From n In Group
                                                                                                                                                          Order By n.x.ДатаЗагрузки
                                                                                                                                                          Select New Grid1Class With {.Организация = n.x?.Клиент, .Актуальность = n.x?.РезультатРаботы,
                                                                                                                                       .Груз = "Маршрут " & vbCrLf & n.y.СтранаПогрузки & " - " & n.y.СтранаВыгрузки & vbCrLf & n.x?.НаименованиеГруза & vbCrLf & "Загрузка - " & IIf(n.x?.ДатаЗагрузки Is Nothing, String.Empty, n.x?.ДатаЗагрузки) _
                                                                                                                                       & vbCrLf & "Выгрузка - " & IIf(n.x?.ДатаВыгрузки Is Nothing, String.Empty, n.x?.ДатаВыгрузки) &
                                                                                                                                       vbCrLf & "Вес - " & n.x?.Вес & vbCrLf & "Обьем - " & n.x?.Обьем & vbCrLf & "Ставка - " & n.y?.Ставка, .Дата = IIf(n.z?.Дата Is Nothing, String.Empty, n.z?.Дата), .КодЖурналГруз = n.x.Код, .ТипАвто = n.x.ТипАвто,
                                                                                                                                        .Контакт = (From v In AllClass.Клиент
                                                                                                                                                    Where v.НазваниеОрганизации = n.x.Клиент
                                                                                                                                                    Select v?.Контактное_лицо & vbCrLf & v?.Телефон).FirstOrDefault()}).ToList()}).ToList()


        If DataAll IsNot Nothing Then
            DataAll.Clear()
        End If

        If f IsNot Nothing Then
            DataAll.AddRange(f)
        End If

        Dim f2 = (From x In AllClass.SkypeПеревозчикПредложение 'добавляем скайп данные по авто
                  Order By x.Дата Descending
                  Where x.Экспедитор = Экспедитор
                  Select x).ToList()

        If f2 IsNot Nothing Then
            For Each b In AllClass.Страна.OrderBy(Function(x) x.Страна).ToList()
                Dim f3 = (From v In f2
                          Where v.Сообщение?.ToUpper.Contains(b.Страна?.ToUpper)
                          Select v).ToList()
                If f3 IsNot Nothing Then
                    If f3.Count > 0 Then
                        Dim m As New СтраныClass With {.ТипВыборки = "Автомобиль", .СтранаЗагрузка = b.Страна, .СтранаВыгрузки = b.Страна, .lst = (From c In f3
                                                                                                                                                   Order By c.Дата Descending
                                                                                                                                                   Select New Grid1Class With {.Организация = c.Перевозчик, .Skype = c.Сообщение}).ToList()}

                        DataAll.Add(m)
                    End If

                End If

            Next
        End If

        Dim f4 = (From x In AllClass.SkypeКлиентПредложение 'добавляем скайп данные по грузу
                  Order By x.Дата Descending
                  Where x.Экспедитор = Экспедитор
                  Select x).ToList()

        If f4 IsNot Nothing Then
            For Each b In AllClass.Страна.OrderBy(Function(x) x.Страна).ToList()
                Dim f5 = (From v In f4
                          Where v.Сообщение?.ToUpper.Contains(b.Страна?.ToUpper)
                          Select v).ToList()
                If f5 IsNot Nothing Then
                    If f5.Count > 0 Then
                        Dim m As New СтраныClass With {.ТипВыборки = "Груз", .СтранаЗагрузка = b.Страна, .СтранаВыгрузки = b.Страна, .lst = (From c In f5
                                                                                                                                             Order By c.Дата Descending
                                                                                                                                             Select New Grid1Class With {.Организация = c.Клиент, .Груз = c.Сообщение}).ToList()}

                        DataAll.Add(m)
                    End If

                End If

            Next
        End If



    End Sub
    Private Sub ПодготовкаАктуальныхДанных()


        Dim f = (From b In (From x In AllClass.ЖурналКлиентГруз   'выборка из таблицы журнал грузы 
                            Join z In AllClass.ЖурналДата On z.Код Equals x.КодЖурналДата
                            Join y In AllClass.ЖурналКлиентМаршрут On x.Код Equals y.КодЖурналКлиентГруз
                            Order By x.ДатаЗагрузки Descending
                            Where x.Экспедитор = Экспедитор And x.ОтоброжатьВТаблицеЖурнала Is Nothing
                            Select x, y, z).ToList()
                 Group b By Keys = New With {Key b.y.СтранаПогрузки, Key b.y.СтранаВыгрузки}
                      Into Group
                 Select New СтраныClass With {.СтранаЗагрузка = Keys.СтранаПогрузки, .СтранаВыгрузки = Keys.СтранаВыгрузки, .ТипВыборки = "Груз", .lst = (From n In Group
                                                                                                                                                          Order By n.x.ДатаЗагрузки
                                                                                                                                                          Select New Grid1Class With {.Организация = n.x?.Клиент, .Актуальность = n.x?.РезультатРаботы,
                                                                                                                                       .Груз = "Маршрут " & vbCrLf & n.y.СтранаПогрузки & " - " & n.y.СтранаВыгрузки & vbCrLf & n.x?.НаименованиеГруза & vbCrLf & "Загрузка - " & IIf(n.x?.ДатаЗагрузки Is Nothing, String.Empty, n.x?.ДатаЗагрузки) _
                                                                                                                                       & vbCrLf & "Выгрузка - " & IIf(n.x?.ДатаВыгрузки Is Nothing, String.Empty, n.x?.ДатаВыгрузки) &
                                                                                                                                       vbCrLf & "Вес - " & n.x?.Вес & vbCrLf & "Обьем - " & n.x?.Обьем & vbCrLf & "Ставка - " & n.y?.Ставка, .Дата = IIf(n.z?.Дата Is Nothing, String.Empty, n.z?.Дата), .КодЖурналГруз = n.x.Код, .ТипАвто = n.x.ТипАвто,
                                                                                                                                        .Контакт = (From v In AllClass.Клиент
                                                                                                                                                    Where v.НазваниеОрганизации = n.x.Клиент
                                                                                                                                                    Select v?.Контактное_лицо & vbCrLf & v?.Телефон).FirstOrDefault()}).ToList()}).ToList()


        If DataAll IsNot Nothing Then
            DataAll.Clear()
        End If

        If f IsNot Nothing Then
            DataAll.AddRange(f)
        End If

        Dim f2 = (From x In AllClass.SkypeПеревозчикПредложение 'добавляем скайп данные по авто
                  Order By x.Дата Descending
                  Where x.Экспедитор = Экспедитор And x?.Дата >= Now.AddDays(-15)
                  Select x).ToList()

        If f2 IsNot Nothing Then
            For Each b In AllClass.Страна.OrderBy(Function(x) x.Страна).ToList()
                Dim f3 = (From v In f2
                          Where v.Сообщение?.ToUpper.Contains(b.Страна?.ToUpper)
                          Select v).ToList()
                If f3 IsNot Nothing Then
                    If f3.Count > 0 Then
                        Dim m As New СтраныClass With {.ТипВыборки = "Автомобиль", .СтранаЗагрузка = b.Страна, .СтранаВыгрузки = b.Страна, .lst = (From c In f3
                                                                                                                                                   Order By c.Дата Descending
                                                                                                                                                   Select New Grid1Class With {.Организация = c.Перевозчик, .Skype = c.Сообщение}).ToList()}

                        DataAll.Add(m)
                    End If

                End If

            Next
        End If

        Dim f4 = (From x In AllClass.SkypeКлиентПредложение 'добавляем скайп данные по грузу
                  Order By x.Дата Descending
                  Where x.Экспедитор = Экспедитор And x?.Дата >= Now.AddDays(-15)
                  Select x).ToList()

        If f4 IsNot Nothing Then
            For Each b In AllClass.Страна.OrderBy(Function(x) x.Страна).ToList()
                Dim f5 = (From v In f4
                          Where v.Сообщение?.ToUpper.Contains(b.Страна?.ToUpper)
                          Select v).ToList()
                If f5 IsNot Nothing Then
                    If f5.Count > 0 Then
                        Dim m As New СтраныClass With {.ТипВыборки = "Груз", .СтранаЗагрузка = b.Страна, .СтранаВыгрузки = b.Страна, .lst = (From c In f5
                                                                                                                                             Order By c.Дата Descending
                                                                                                                                             Select New Grid1Class With {.Организация = c.Клиент, .Груз = c.Сообщение}).ToList()}

                        DataAll.Add(m)
                    End If

                End If

            Next
        End If



    End Sub
    Private Async Sub LoadComboAsync()
        Await Task.Run(Sub() LoadCombo())
        bscom1.ResetBindings(False)
        bscom2.ResetBindings(False)
        ComboBox1.Text = String.Empty
        ComboBox2.Text = String.Empty

    End Sub
    Private Sub LoadCombo()
        Dim mo As New AllUpd
        Do While AllClass.Страна Is Nothing
            mo.СтранаAll()
        Loop
        If com1all IsNot Nothing Then
            com1all.Clear()
        End If
        If com2all IsNot Nothing Then
            com2all.Clear()
        End If



        Dim f = AllClass.Страна.OrderBy(Function(x) x.Страна).Select(Function(x) x.Страна).Distinct.ToList()
        If f IsNot Nothing Then
            For Each b In f
                Dim m As New IDNaz With {.Naz = b}
                com1all.Add(m)
                com2all.Add(m)
            Next

        End If
    End Sub
    Public Class Grid1Class
        Public Property Номер As Integer
        Public Property Организация As String
        Public Property Дата As String
        Public Property Контакт As String
        Public Property ТипАвто As String
        Public Property Груз As String
        Public Property Skype As String
        Public Property Актуальность As String
        Public Property КодЖурналГруз As String
    End Class
    Public Class СтраныClass

        Public Property ТипВыборки As String
        Public Property СтранаЗагрузка As String
        Public Property СтранаВыгрузки As String

        Public Property lst As List(Of Grid1Class)


    End Class

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If ComboBox1.Text.Length = 0 Then
            MessageBox.Show("Выберите страну загрузки!", Рик)
            Return
        End If
        Dim f1 As IDNaz = ComboBox1.SelectedItem

        If Grid1all IsNot Nothing Then
            Grid1all.Clear()
        End If


        If DataAll IsNot Nothing Then
            If ComboBox3.Text.Length = 0 Then
                Dim f = (From x In DataAll
                         Where x.СтранаЗагрузка = f1.Naz
                         Select x).ToList()
                If f IsNot Nothing Then
                    Dim i As Integer = 1
                    For Each b In f
                        For Each b1 In b.lst
                            b1.Номер = i
                            i += 1
                            Grid1all.Add(b1)
                        Next
                    Next
                    bsGrid1.ResetBindings(False)
                End If
            Else
                Dim f = (From x In DataAll
                         Where x.СтранаЗагрузка = f1.Naz And x.ТипВыборки = ComboBox3.Text
                         Select x).ToList()
                If f IsNot Nothing Then
                    Dim i As Integer = 1
                    For Each b In f
                        For Each b1 In b.lst
                            b1.Номер = i
                            i += 1
                            Grid1all.Add(b1)
                        Next
                    Next
                    bsGrid1.ResetBindings(False)
                End If

            End If
        End If


    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        Dim f1 As IDNaz = ComboBox2.SelectedItem

        If f1 Is Nothing Then
            MessageBox.Show("Выберите страну загрузки!", Рик)
            Return
        End If

        If Grid1all IsNot Nothing Then
            Grid1all.Clear()
        End If


        If DataAll IsNot Nothing Then
            If ComboBox3.Text.Length = 0 Then
                Dim f = (From x In DataAll
                         Where x.СтранаВыгрузки = f1.Naz
                         Select x).ToList()
                If f IsNot Nothing Then
                    Dim i As Integer = 1
                    For Each b In f
                        For Each b1 In b.lst
                            b1.Номер = i
                            i += 1
                            Grid1all.Add(b1)
                        Next
                    Next
                    bsGrid1.ResetBindings(False)
                End If
            Else
                Dim f = (From x In DataAll
                         Where x.СтранаВыгрузки = f1.Naz And x.ТипВыборки = ComboBox3.Text
                         Select x).ToList()
                If f IsNot Nothing Then
                    Dim i As Integer = 1
                    For Each b In f
                        For Each b1 In b.lst
                            b1.Номер = i
                            i += 1
                            Grid1all.Add(b1)
                        Next
                    Next
                    bsGrid1.ResetBindings(False)
                End If

            End If
        End If
    End Sub
    Private Async Sub ПодготовкаВсехДанныхAsync()
        Await Task.Run(Sub() ПодготовкаВсехДанных())
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked = True Then
            ПодготовкаВсехДанныхAsync()
            CheckBox2.Checked = False
        End If
    End Sub
    Private Async Sub ПодготовкаАктуальныхДанныхAsync(Optional d As Boolean = False)
        Await Task.Run(Sub() ПодготовкаАктуальныхДанных())
        If d = True Then
            АктуалДанные()
            bsGrid1.ResetBindings(False)
        End If
    End Sub

    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox2.CheckedChanged
        If CheckBox2.Checked = True Then
            ПодготовкаАктуальныхДанных()
            CheckBox1.Checked = False
        End If

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        ПодготовкаАктуальныхДанныхAsync(True)
    End Sub
    Private Sub АктуалДанные()
        If Grid1all IsNot Nothing Then
            Grid1all.Clear()
        End If


        If DataAll IsNot Nothing Then
            Dim i As Integer = 1
            For Each b In DataAll
                For Each b1 In b.lst
                    b1.Номер = i
                    i += 1
                    Grid1all.Add(b1)
                Next
            Next

        End If

    End Sub
End Class