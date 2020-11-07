Public Class ЖурналДобавитьГруз
    Private com1 As List(Of IDNaz)
    Private bscom1 As BindingSource
    Private com2 As List(Of IDNaz)
    Private bscom2 As BindingSource
    Private Grid1All As ГрузыClass
    Private bsGrid1 As BindingSource
    Private Grid2All As List(Of МаршрутыClass)
    Private bsGrid2 As BindingSource

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged



    End Sub
    Private Async Sub ПредзагрузкаAsync()
        Await Task.Run(Sub() Предзагрузка())
    End Sub
    Private Sub Предзагрузка()
        Dim mo As New AllUpd

        Do While AllClass.ЖурналКлиентГруз Is Nothing
            mo.ЖурналКлиентГрузAll()
        Loop


        Do While AllClass.ЖурналКлиентМаршрут Is Nothing
            mo.ЖурналКлиентМаршрутAll()
        Loop

        Do While AllClass.ЖурналДата Is Nothing
            mo.ЖурналДатаAll()
        Loop
        Do While AllClass.Клиент Is Nothing
            mo.КлиентAll()
        Loop
    End Sub

    Private Sub ЖурналДобавитьГруз_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Label2.Text = Now.Date
        com1 = New List(Of IDNaz)
        bscom1 = New BindingSource
        bscom1.DataSource = com1
        ComboBox1.DataSource = bscom1
        ComboBox1.DisplayMember = "Naz"

        com2 = New List(Of IDNaz)
        bscom2 = New BindingSource
        bscom2.DataSource = com2
        ComboBox2.DataSource = bscom2
        ComboBox2.DisplayMember = "Naz"
        ComboBox2.Text = String.Empty

        Dim mo As New AllUpd
        Do While AllClass.Клиент Is Nothing
            mo.КлиентAll()
        Loop

        Do While AllClass.ЖурналКлиентГруз Is Nothing
            mo.ЖурналКлиентГрузAll()
        Loop


        Do While AllClass.ЖурналКлиентСписок Is Nothing
            mo.ЖурналКлиентСписокAll()
        Loop


        Dim f = AllClass.Клиент.OrderBy(Function(x) x.НазваниеОрганизации).Select(Function(x) New IDNaz With {.Naz = x.НазваниеОрганизации}).Distinct.ToList()
        Dim f1 = AllClass.ЖурналКлиентГруз.OrderBy(Function(x) x.Клиент).Select(Function(x) New IDNaz With {.Naz = x.Клиент}).Distinct.ToList()
        Dim f4 = AllClass.ЖурналКлиентСписок.OrderBy(Function(x) x.Клиент).Select(Function(x) New IDNaz With {.Naz = x.Клиент}).Distinct.ToList()
        Dim f2 = f1.Select(Function(x) x.Naz).ToList().Union(f.Select(Function(x) x.Naz).ToList())
        Dim f5 = f2.OrderBy(Function(x) x).Select(Function(x) x).ToList().Union(f4.OrderBy(Function(x) x).Select(Function(x) x.Naz).ToList())
        For Each b In f5
            Dim f3 As New IDNaz With {.Naz = b}
            com1.Add(f3)
        Next
        bscom1.ResetBindings(False)
        ComboBox1.Text = String.Empty


        Grid1All = New ГрузыClass
        bsGrid1 = New BindingSource
        bsGrid1.DataSource = Grid1All
        Grid1.DataSource = bsGrid1
        GridView(Grid1)
        Grid1.SelectionMode = DataGridViewSelectionMode.CellSelect
        Grid1.Columns(7).HeaderText = "Тип Погрузки"
        Grid1.Columns(8).HeaderText = "Паллеты штук"
        Grid1.Columns(9).HeaderText = "Размер паллет"
        Grid1.Columns(11).HeaderText = "Тип авто"
        Grid1.Columns(12).HeaderText = "Доп информация"
        Grid1.Columns(0).Visible = False
        Grid1.Columns(1).Width = "200"
        Grid1.Columns(2).Width = "80"
        Grid1.Columns(2).HeaderText = "Вес (кг)"


        Grid2All = New List(Of МаршрутыClass)
        bsGrid2 = New BindingSource
        bsGrid2.DataSource = Grid2All
        Grid2.DataSource = bsGrid2
        GridView(Grid2)
        Grid2.SelectionMode = DataGridViewSelectionMode.CellSelect

        Grid2.Columns(1).HeaderText = "Страна Погрузки"
        Grid2.Columns(2).HeaderText = "Страна Выгрузки"
        Grid2.Columns(3).HeaderText = "Город Погрузки"
        Grid2.Columns(4).HeaderText = "Город Выгрузки"
        Grid2.Columns(5).HeaderText = "Квадрат Погрузки"
        Grid2.Columns(6).HeaderText = "Квадрат Выгрузки"
        Grid2.Columns(7).HeaderText = "Таможня Отправления"
        Grid2.Columns(8).HeaderText = "Таможня Назначения"
        Grid2.Columns(11).HeaderText = "Доп информация"
        Grid2.Columns(0).Visible = False
        Grid2.Columns(9).Width = "100"

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If TextBox1.Text.Length = 0 Then
            MessageBox.Show("Заполните поле название!", Рик)
            Return
        End If
        If TextBox2.Text.Length = 0 Then
            MessageBox.Show("Заполните поле телефон!", Рик)
            Return
        End If
        If TextBox3.Text.Length = 0 Then
            MessageBox.Show("Заполните поле конт.лицо!", Рик)
            Return
        End If
        Dim txt1 = Trim(TextBox1.Text)
        Dim f1 = com1.Where(Function(x) x.Naz = txt1).Select(Function(x) x).FirstOrDefault()
        If f1 IsNot Nothing Then
            MessageBox.Show("Такой клиент уже существует!", Рик)
            Return
        End If


        Dim f As New ЖурналКлиентСписок
        With f
            .Клиент = txt1
            .КонтактноеЛицо = TextBox2.Text
            .Телефон = TextBox3.Text
        End With
        Using db As New dbAllDataContext()
            db.ЖурналКлиентСписок.InsertOnSubmit(f)
            db.SubmitChanges()
        End Using

        Dim f4 As New IDNaz With {.Naz = txt1}
        com1.Add(f4)
        bscom1.ResetBindings(False)
        ComboBox1.Text = txt1

        TextBox1.Text = String.Empty
        TextBox2.Text = String.Empty
        TextBox3.Text = String.Empty
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If ComboBox1.Text.Length = 0 Then
            MessageBox.Show("Выберите название клиента!", Рик)
            Return
        End If

        If MaskedTextBox1.MaskCompleted = False Then
            MessageBox.Show("Выберите дату загрузки!", Рик)
            Return
        End If
        If MaskedTextBox2.MaskCompleted = False Then
            MessageBox.Show("Выберите дату выгрузки!", Рик)
            Return
        End If

        Dim f7 As New ЖурналДата
        Dim f As IDNaz = ComboBox1.SelectedItem
        Using db As New dbAllDataContext()
            Dim f1 = db.ЖурналДата.Where(Function(x) x.Дата = Label2.Text).Select(Function(x) x).FirstOrDefault()
            If f1 Is Nothing Then
                Dim f2 As New ЖурналДата
                f2.Дата = Label2.Text
                db.ЖурналДата.InsertOnSubmit(f2)
                db.SubmitChanges()

                f7.Дата = f2.Дата
                f7.Код = f2.Код
            Else
                f7.Дата = f1.Дата
                f7.Код = f1.Код
            End If


            Dim f3 As New ЖурналКлиентГруз
            With f3
                .КодЖурналДата = f7.Код
                .Вес = Grid1All.Вес
                .Высота = Grid1All.Высота
                .Длина = Grid1All.Длина
                .ДополнитИнформация = Grid1All.ДополнитИнформация
                .Клиент = f.Naz
                .НаименованиеГруза = Grid1All.Наименование
                .Ширина = Grid1All.Ширина
                .Обьем = Grid1All.Обьем
                .РазмерПаллет = Grid1All.РазмерПаллет
                .ПаллетыШтук = Grid1All.ПаллетыШтук
                .ТипПогрузки = Grid1All.ТипПогрузки
                .ADR = Grid1All.ADR
                .ДатаЗагрузки = MaskedTextBox1.Text
                .ДатаВыгрузки = MaskedTextBox2.Text
                .ТипАвто = Grid1All.ТипАвто
            End With
            db.ЖурналКлиентГруз.InsertOnSubmit(f3)
            db.SubmitChanges()

            Dim f4 As New List(Of ЖурналКлиентМаршрут)
            For Each b In Grid2All
                Dim f5 As New ЖурналКлиентМаршрут
                With f5
                    .КодЖурналКлиентГруз = f3.Код
                    .Клиент = f.Naz
                    .СтранаПогрузки = b.СтранаПогрузки
                    .СтранаВыгрузки = b.СтранаВыгрузки
                    .ГородПогрузки = b.ГородПогрузки
                    .ГородВыгрузки = b.ГородВыгрузки
                    .КвадратПогрузки = b.КвадратПогрузки
                    .КвадратВыгрузки = b.КвадратВыгрузки
                    .ТаможняОтправления = b.ТаможняОтправления
                    .ТаможняНазначения = b.ТаможняНазначения
                    .Ставка = b.Ставка
                    .EX = b.EX
                    .ДополнитИнформация = b.ДополнитИнформация

                End With
                f4.Add(f5)
            Next
            db.ЖурналКлиентМаршрут.InsertAllOnSubmit(f4)
            db.SubmitChanges()
        End Using
        Dim mo As New AllUpd
        mo.ЖурналДатаAllAsync()
        mo.ЖурналКлиентГрузAllAsync()
        mo.ЖурналКлиентМаршрутAllAsync()

        Close()

    End Sub

    Private Sub MaskedTextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles MaskedTextBox1.KeyDown
        If e.KeyCode = Keys.Enter Then
            MaskedTextBox2.Focus()
        End If
    End Sub



    Private Sub ComboBox1_SelectedValueChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedValueChanged
        Dim f As IDNaz = ComboBox1.SelectedItem
        If com2 IsNot Nothing Then
            com2.Clear()
        End If
        If f Is Nothing Then
            Return
        End If
        Dim mo As New AllUpd

        Do While AllClass.ЖурналКлиентГруз Is Nothing
            mo.ЖурналКлиентГрузAll()
        Loop


        Do While AllClass.ЖурналКлиентМаршрут Is Nothing
            mo.ЖурналКлиентМаршрутAll()
        Loop

        Do While AllClass.ЖурналДата Is Nothing
            mo.ЖурналДатаAll()
        Loop
        If com2 IsNot Nothing Then
            com2.Clear()
        End If

        Dim f1 = (From x In AllClass.ЖурналКлиентГруз
                  Join y In AllClass.ЖурналКлиентМаршрут On x.Код Equals y.КодЖурналКлиентГруз
                  Join z In AllClass.ЖурналДата On z.Код Equals x.КодЖурналДата
                  Where x.Клиент = f.Naz
                  Select x, y, z).ToList()
        ' Select Case New IDNaz With {.ID = x.Код, .Naz = z.Дата & ": " & IIf(x.НаименованиеГруза Is Nothing, String.Empty, x.НаименованиеГруза) & " (" & y.СтранаПогрузки & " - " & y.СтранаВыгрузки & ")"}).ToList()
        Dim f2 = (From x In f1
                  Group x By Keys = New With {Key x.x.Код, Key .P = x.z.Дата & ": " & IIf(x.x.НаименованиеГруза Is Nothing, String.Empty, x.x.НаименованиеГруза) & " (" & x.y.СтранаПогрузки & " - " & x.y.СтранаВыгрузки & ")"}
                      Into Group
                  Select New IDNaz With {.ID = Keys.Код, .Naz = Keys.P}).ToList()

        If f2 IsNot Nothing Then
            If f2.Count > 0 Then
                com2.AddRange(f2)
                bscom2.ResetBindings(False)
            End If
        End If
        ComboBox2.Text = String.Empty
    End Sub

    Private Sub ComboBox2_SelectedValueChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedValueChanged
        Dim f As IDNaz = ComboBox2.SelectedItem
        If Grid1All IsNot Nothing Then
            Grid1All = Nothing
        End If

        If Grid2All IsNot Nothing Then
            Grid2All.Clear()
        End If
        If f Is Nothing Then
            Return
        End If

        Dim mo As New AllUpd

        Do While AllClass.ЖурналКлиентГруз Is Nothing
            mo.ЖурналКлиентГрузAll()
        Loop


        Do While AllClass.ЖурналКлиентМаршрут Is Nothing
            mo.ЖурналКлиентМаршрутAll()
        Loop

        Do While AllClass.ЖурналДата Is Nothing
            mo.ЖурналДатаAll()
        Loop

        Dim f1 = From x In AllClass.ЖурналДата
                 Join y In AllClass.ЖурналКлиентГруз On x.Код Equals y.КодЖурналДата
                 Join z In AllClass.ЖурналКлиентМаршрут On y.Код Equals z.КодЖурналКлиентГруз
                 Where y.Код = f.ID
                 Select New ГрузыClass With {.ADR = y.ADR, .Вес = y.Вес, .Высота = y.Высота, .Длина = y.Длина}




    End Sub
End Class