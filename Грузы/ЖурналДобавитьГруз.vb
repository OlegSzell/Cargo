Public Class ЖурналДобавитьГруз
    Private com1All As List(Of IDNaz)
    Private com1Sel As List(Of IDNaz)
    Private com1 As List(Of IDNaz)
    Private bscom1 As BindingSource
    Private com2 As List(Of Com2ЖурналДобавитьГрузClass)
    Private bscom2 As BindingSource
    Private Grid1All As List(Of ГрузыClass)
    Private bsGrid1 As BindingSource
    Private Grid2All As List(Of МаршрутыClass)
    Private bsGrid2 As BindingSource
    Private AutoСтрана As AutoCompleteStringCollection
    Private AutoРегионы As AutoCompleteStringCollection
    Public ForGrid2журналс As Grid2ЖурналClass
    Public Flag As Boolean = False
    Private Rezult As String = Nothing

    Private Sub AutoCompl()
        Dim mo As New AllUpd
        Do While AllClass.Страна Is Nothing
            mo.СтранаAll()
        Loop
        Do While AllClass.РегионыРоссии Is Nothing
            mo.РегионыРоссииAll()
        Loop
        AutoСтрана = New AutoCompleteStringCollection
        AutoРегионы = New AutoCompleteStringCollection

        Dim f = AllClass.Страна.OrderBy(Function(x) x.Страна).Select(Function(x) x.Страна).Distinct()
        For Each b In f
            AutoСтрана.Add(b)
        Next


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

        Do While AllClass.Страна Is Nothing
            mo.СтранаAll()
        Loop
        Do While AllClass.РегионыРоссии Is Nothing
            mo.РегионыРоссииAll()
        Loop

        Do While AllClass.ФайлыExcelВсе Is Nothing
            mo.ФайлыExcelВсеAll()
        Loop
    End Sub

    Private Sub ЖурналДобавитьГруз_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ПредзагрузкаAsync()


        Label2.Text = Now.Date
        com1All = New List(Of IDNaz)
        com1Sel = New List(Of IDNaz)
        com1 = New List(Of IDNaz)
        bscom1 = New BindingSource
        bscom1.DataSource = com1
        ComboBox1.DataSource = bscom1
        ComboBox1.DisplayMember = "Naz"

        com2 = New List(Of Com2ЖурналДобавитьГрузClass)  'хранит все данные для гридов
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
        Dim f1 = AllClass.ЖурналКлиентГруз.OrderBy(Function(x) x.Клиент).Where(Function(x) x.Экспедитор = Экспедитор).Select(Function(x) New IDNaz With {.Naz = x.Клиент}).Distinct.ToList()
        Dim f4 = AllClass.ЖурналКлиентСписок.OrderBy(Function(x) x.Клиент).Select(Function(x) New IDNaz With {.Naz = x.Клиент}).Distinct.ToList()
        Dim f2 = f1.Select(Function(x) x.Naz).ToList().Union(f.Select(Function(x) x.Naz).ToList()).ToList()
        Dim f5 = f2.Union(f4.Select(Function(x) x.Naz)).ToList()
        Dim f6 = f5.OrderBy(Function(x) x).ToList()
        For Each b In f6
            Dim f3 As New IDNaz With {.Naz = b}
            com1All.Add(f3)
        Next

        Dim f7 = (From y In ((From x In AllClass.ЖурналКлиентГруз
                              Where x.Экспедитор = Экспедитор
                              Order By x.Клиент
                              Select x.Клиент).Distinct())
                  Select New IDNaz With {.Naz = y}).ToList()


        'Dim f8 = f7.Select(Function(x) New IDNaz With {.Naz = x}).ToList()
        com1.AddRange(f7)
        com1Sel.AddRange(f7)
        bscom1.ResetBindings(False)
        ComboBox1.Text = String.Empty


        Grid1All = New List(Of ГрузыClass)
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
        Grid1.Columns(3).Visible = False
        Grid1.Columns(4).Visible = False
        Grid1.Columns(5).Visible = False
        Grid1.Columns(6).Visible = False
        Grid1.Columns(8).Visible = False
        Grid1.Columns(9).Visible = False
        Grid1.Columns(10).Visible = False
        Grid1.Columns(12).Visible = False
        Grid1.Columns(1).MinimumWidth = "200"
        Grid1.Columns(2).MinimumWidth = "100"
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


        AutoCompl()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Dim mo As New AllUpd
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
            .ДопИнфо = Rezult
        End With
        Using db As New dbAllDataContext()
            db.ЖурналКлиентСписок.InsertOnSubmit(f)
            db.SubmitChanges()
        End Using

        Dim f4 As New IDNaz With {.Naz = txt1}
        com1All.Add(f4)
        Dim ms As New List(Of IDNaz)
        ms.AddRange(com1All.OrderBy(Function(x) x.Naz).ToList())

        com1All.Clear()
        com1All.AddRange(ms)
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

        If MessageBox.Show("Добавить груз?", Рик, MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.No Then
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
                .Вес = Grid1All(0).Вес
                .Высота = Grid1All(0).Высота
                .Длина = Grid1All(0).Длина
                .ДополнитИнформация = Grid1All(0).ДополнитИнформация
                .Клиент = f.Naz
                .НаименованиеГруза = Grid1All(0).Наименование
                .Ширина = Grid1All(0).Ширина
                .Обьем = Grid1All(0).Обьем
                .РазмерПаллет = Grid1All(0).РазмерПаллет
                .ПаллетыШтук = Grid1All(0).ПаллетыШтук
                .ТипПогрузки = Grid1All(0).ТипПогрузки
                .ADR = Grid1All(0).ADR
                .ДатаЗагрузки = MaskedTextBox1.Text
                .ДатаВыгрузки = MaskedTextBox2.Text
                .ТипАвто = Grid1All(0).ТипАвто
                .Экспедитор = Экспедитор

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
        mo.ЖурналКлиентГрузAll()
        mo.ЖурналКлиентМаршрутAll()
        ForGrid2журнал(f.Naz)
        Flag = True
        Close()

    End Sub
    Private Sub ForGrid2журнал(ByVal Клиент As String)

        Dim marsh As String = Nothing
        For Each b In Grid2All
            If marsh Is Nothing Then
                marsh = "( " & b.СтранаПогрузки & " ) " & b.ГородПогрузки & " - " & "( " & b.СтранаВыгрузки & " ) " & b.ГородВыгрузки
            Else
                marsh = marsh & vbCrLf & "( " & b.СтранаПогрузки & " ) " & b.ГородПогрузки & " - " & "( " & b.СтранаВыгрузки & " ) " & b.ГородВыгрузки
            End If
        Next
        ForGrid2журналс = New Grid2ЖурналClass

        Dim f As New Grid2ЖурналClass With {.Дата = CDate(Label2.Text), .ДатаЗагрузки = MaskedTextBox1.Text, .Клиент = Клиент,
            .Страны = marsh, .Груз = Grid1All(0).Наименование}
        If f IsNot Nothing Then
            ForGrid2журналс = f
        End If

    End Sub

    Private Sub MaskedTextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles MaskedTextBox1.KeyDown
        If e.KeyCode = Keys.Enter Then
            MaskedTextBox2.Focus()
        End If
    End Sub


    Private Sub ComboBox2_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles ComboBox2.SelectionChangeCommitted
        'Dim f As IDNaz = ComboBox2.SelectedItem
        'If Grid1All IsNot Nothing Then
        '    Grid1All = Nothing
        'End If

        'If Grid2All IsNot Nothing Then
        '    Grid2All.Clear()
        'End If
        'If f Is Nothing Then
        '    Return
        'End If

        'Dim mo As New AllUpd

        'Do While AllClass.ЖурналКлиентГруз Is Nothing
        '    mo.ЖурналКлиентГрузAll()
        'Loop


        'Do While AllClass.ЖурналКлиентМаршрут Is Nothing
        '    mo.ЖурналКлиентМаршрутAll()
        'Loop

        'Do While AllClass.ЖурналДата Is Nothing
        '    mo.ЖурналДатаAll()
        'Loop

        'Dim f1 = From x In AllClass.ЖурналДата
        '         Join y In AllClass.ЖурналКлиентГруз On x.Код Equals y.КодЖурналДата
        '         Join z In AllClass.ЖурналКлиентМаршрут On y.Код Equals z.КодЖурналКлиентГруз
        '         Where y.Код = f.ID
        '         Select New ГрузыClass With {.ADR = y.ADR, .Вес = y.Вес, .Высота = y.Высота, .Длина = y.Длина}
        Dim m1 As Com2ЖурналДобавитьГрузClass = ComboBox2.SelectedItem
        Dim k As New ГрузыClass With {.ADR = m1.ADR, .Вес = m1.Вес, .Высота = m1.Высота, .Длина = m1.Длина, .ДополнитИнформация = m1.ДополнитИнформация,
            .Наименование = m1.Наименование, .Номер = m1.Номер, .Обьем = m1.Обьем, .ПаллетыШтук = m1.ПаллетыШтук, .РазмерПаллет = m1.РазмерПаллет,
            .ТипАвто = m1.ТипАвто, .ТипПогрузки = m1.ТипПогрузки, .Ширина = m1.Ширина}
        If Grid1All IsNot Nothing Then
            Grid1All.Clear()
        End If
        If k IsNot Nothing Then
            If k.Длина IsNot Nothing Then
                If k.Длина.Length > 0 Then
                    Grid1.Columns(4).Visible = True
                End If
            Else
                Grid1.Columns(4).Visible = False
            End If
            If k.Высота IsNot Nothing Then
                If k.Высота.Length > 0 Then
                    Grid1.Columns(6).Visible = True
                End If
            Else
                Grid1.Columns(6).Visible = False

            End If

            If k.Ширина IsNot Nothing Then
                If k.Ширина.Length > 0 Then
                    Grid1.Columns(5).Visible = True
                End If
            Else
                Grid1.Columns(5).Visible = False
            End If

            If k.Обьем IsNot Nothing Then
                If k.Обьем.Length > 0 Then
                    Grid1.Columns(3).Visible = True
                End If
            Else
                Grid1.Columns(3).Visible = False
            End If

            If k.ПаллетыШтук > 0 Then

                Grid1.Columns(8).Visible = True
            Else
                Grid1.Columns(8).Visible = False

            End If

            If k.РазмерПаллет IsNot Nothing Then
                If k.РазмерПаллет.Length > 0 Then
                    Grid1.Columns(9).Visible = True
                End If
            Else
                Grid1.Columns(9).Visible = False
            End If

            If k.ADR IsNot Nothing Then
                If k.ADR.Length > 0 Then
                    Grid1.Columns(10).Visible = True
                End If
            Else
                Grid1.Columns(10).Visible = False
            End If

            If k.ДополнитИнформация IsNot Nothing Then
                If k.ДополнитИнформация.Length > 0 Then
                    Grid1.Columns(12).Visible = True
                End If
            Else
                Grid1.Columns(12).Visible = False
            End If


        End If

        Grid1All.Add(k)

        bsGrid1.ResetBindings(False)

        If Grid2All IsNot Nothing Then
            Grid2All.Clear()
        End If
        MaskedTextBox1.Text = m1.ДатаЗагрузки
        MaskedTextBox2.Text = m1.ДатаВыгрузки


        Dim i As Integer = 1
        For Each b In m1.Lst
            Dim m2 As New МаршрутыClass With {.EX = b.EX, .ГородВыгрузки = b.ГородВыгрузки, .ГородПогрузки = b.ГородПогрузки, .ДополнитИнформация = b.ДополнитИнформация,
                .КвадратВыгрузки = b.КвадратВыгрузки, .КвадратПогрузки = b.КвадратПогрузки, .Номер = i, .Ставка = b.Ставка, .СтранаВыгрузки = b.СтранаВыгрузки,
                .СтранаПогрузки = b.СтранаПогрузки, .ТаможняНазначения = b.ТаможняНазначения, .ТаможняОтправления = b.ТаможняОтправления}
            Grid2All.Add(m2)
            i += 1
        Next
        bsGrid2.ResetBindings(False)
    End Sub

    Private Sub ComboBox1_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles ComboBox1.SelectionChangeCommitted
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
                  Select New Com2ЖурналДобавитьГрузClass With {.Naz = Keys.P, .КодЖурнГруз = Keys.Код, .Наименование = Group(0).x.НаименованиеГруза, .Вес = Group(0).x.Вес,
                      .Обьем = Group(0).x.Обьем, .Длина = Group(0).x.Длина, .Ширина = Group(0).x.Ширина, .Высота = Group(0).x.Высота,
                      .ТипПогрузки = Group(0).x.ТипПогрузки, .ПаллетыШтук = Group(0).x.ПаллетыШтук, .РазмерПаллет = Group(0).x.РазмерПаллет, .ADR = Group(0).x.ADR,
                      .ТипАвто = Group(0).x.ТипАвто, .ДополнитИнформация = Group(0).x.ДополнитИнформация, .ДатаЗагрузки = Group(0).x.ДатаЗагрузки,
                      .ДатаВыгрузки = Group(0).x.ДатаВыгрузки, .Lst = (From b In Group
                                                                       Select b.y).ToList()}).ToList()


        If f2 IsNot Nothing Then
            Dim k = f2.OrderByDescending(Function(x) x.Naz).ToList()
            If k.Count > 0 Then
                com2.AddRange(k)
                bscom2.ResetBindings(False)
            End If
        End If
        ComboBox2.Text = String.Empty
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked = True Then
            If com1 IsNot Nothing Then
                com1.Clear()
            End If
            com1.AddRange(com1All)
        Else
            If com1 IsNot Nothing Then
                com1.Clear()
            End If
            com1.AddRange(com1Sel)
        End If
        bscom1.ResetBindings(False)
    End Sub



    'Private Sub Grid2_EditingControlShowing_1(sender As Object, e As DataGridViewEditingControlShowingEventArgs) Handles Grid2.EditingControlShowing

    '    If Grid2.CurrentCell.ColumnIndex = 1 Then
    '        Dim prodCode As TextBox = TryCast(e.Control, TextBox)
    '        If prodCode IsNot Nothing Then
    '            prodCode.AutoCompleteMode = AutoCompleteMode.SuggestAppend
    '            prodCode.AutoCompleteCustomSource = AutoСтрана
    '            prodCode.AutoCompleteSource = AutoCompleteSource.CustomSource
    '            'Else

    '            '    prodCode.AutoCompleteCustomSource = Nothing
    '        End If


    '    End If
    'End Sub

    Private Sub Grid2_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles Grid2.CellEndEdit
        If Grid2.CurrentCell.ColumnIndex = 1 Or Grid2.CurrentCell.ColumnIndex = 2 Then
            Dim f = AllClass.Страна.Where(Function(x) x.Страна.ToUpper.Contains(Grid2.CurrentCell.Value.ToString.ToUpper)).Select(Function(x) x.Страна).FirstOrDefault()
            If f IsNot Nothing Then
                If f.Length > 0 Then
                    Grid2.CurrentCell.Value = f
                End If
            End If
        End If

        If Grid2.CurrentCell.ColumnIndex = 5 Then
            If Grid2.CurrentRow.Cells(1).Value IsNot Nothing Then
                Dim stran = Grid2.CurrentRow.Cells(1).Value
                Dim stran2 = Grid2.CurrentRow.Cells(5).Value
                If stran2 IsNot Nothing Then
                    Dim k = (From x In AllClass.Страна
                             Join y In AllClass.РегионыРоссии On x.Код Equals y.Страны
                             Where x.Страна = stran And y?.Регионы?.ToUpper.Contains(stran2?.ToString.ToUpper)
                             Select y.Регионы).FirstOrDefault()
                    If k IsNot Nothing Then
                        If k.Length > 0 Then
                            Grid2.CurrentCell.Value = k
                        End If
                    End If
                End If


            End If
        End If

        If Grid2.CurrentCell.ColumnIndex = 6 Then
            If Grid2.CurrentRow.Cells(2).Value IsNot Nothing Then
                Dim stran = Grid2.CurrentRow.Cells(2).Value
                Dim stran2 = Grid2.CurrentRow.Cells(6).Value
                If stran2 IsNot Nothing Then
                    Dim k = (From x In AllClass.Страна
                             Join y In AllClass.РегионыРоссии On x.Код Equals y.Страны
                             Where x.Страна = stran And y?.Регионы?.ToUpper.Contains(stran2?.ToString.ToUpper)
                             Select y.Регионы).FirstOrDefault()
                    If k IsNot Nothing Then
                        If k.Length > 0 Then
                            Grid2.CurrentCell.Value = k
                        End If
                    End If
                End If


            End If
        End If


    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim com2sell As Com2ЖурналДобавитьГрузClass = ComboBox2.SelectedItem
        If com2sell Is Nothing Then
            MessageBox.Show("Выберите груз?", Рик)
            Return
        End If

        If com2sell.КодЖурнГруз = 0 Then
            MessageBox.Show("Выберите груз?", Рик)
            Return
        End If

        If MessageBox.Show("Сохранить изменение?", Рик, MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.No Then
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

        Dim mo As New AllUpd

        Do While AllClass.ЖурналКлиентГруз Is Nothing
            mo.ЖурналКлиентГрузAll()
        Loop


        Do While AllClass.ЖурналКлиентМаршрут Is Nothing
            mo.ЖурналКлиентМаршрутAll()
        Loop







        Dim f7 As New ЖурналДата
        Dim f As IDNaz = ComboBox1.SelectedItem
        Using db As New dbAllDataContext()
            Dim f5 = (From x In db.ЖурналКлиентГруз
                      Join y In db.ЖурналКлиентМаршрут On x.Код Equals y.КодЖурналКлиентГруз
                      Where x.Код = com2sell.КодЖурнГруз And x.Экспедитор = Экспедитор
                      Select x, y).FirstOrDefault()

            If f5 IsNot Nothing Then 'обновляем данные
                With f5.x
                    .ADR = Grid1All(0).ADR
                    .НаименованиеГруза = Grid1All(0).Наименование
                    .Вес = Grid1All(0).Вес
                    .Высота = Grid1All(0).Высота
                    .Длина = Grid1All(0).Длина
                    .ДополнитИнформация = Grid1All(0).ДополнитИнформация
                    .Ширина = Grid1All(0).Ширина
                    .Обьем = Grid1All(0).Обьем
                    .РазмерПаллет = Grid1All(0).РазмерПаллет
                    .ПаллетыШтук = Grid1All(0).ПаллетыШтук
                    .ТипПогрузки = Grid1All(0).ТипПогрузки
                    .ДатаЗагрузки = MaskedTextBox1.Text
                    .ДатаВыгрузки = MaskedTextBox2.Text
                    .ТипАвто = Grid1All(0).ТипАвто
                    db.SubmitChanges()
                End With

                Dim f10 = db.ЖурналКлиентМаршрут.Where(Function(x) x.КодЖурналКлиентГруз = com2sell.КодЖурнГруз).Select(Function(x) x).ToList()
                If f10 IsNot Nothing Then
                    If f10.Count > 0 Then
                        db.ЖурналКлиентМаршрут.DeleteAllOnSubmit(f10) 'удаляем данные из таблицы 
                        db.SubmitChanges()

                    End If
                End If


                Dim f8 As New List(Of ЖурналКлиентМаршрут) 'добавляем новые данные
                For Each b In Grid2All
                    Dim f9 As New ЖурналКлиентМаршрут
                    With f9
                        .КодЖурналКлиентГруз = com2sell.КодЖурнГруз
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
                    f8.Add(f9)
                Next
                db.ЖурналКлиентМаршрут.InsertAllOnSubmit(f8)
                db.SubmitChanges()
            End If
        End Using
        mo.ЖурналКлиентГрузAll()
        mo.ЖурналКлиентМаршрутAll()
        ForGrid2журнал(f.Naz)
        Close()


    End Sub

    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click
        If String.IsNullOrEmpty(TextBox4.Text) Then
            MessageBox.Show("Напишите в поле искомое значение", Рик)
            Return
        End If
        Dim txt As String = TextBox4.Text
        Dim mo As New AllUpd
        Do While AllClass.ФайлыExcelВсе Is Nothing
            mo.ФайлыExcelВсеAll()
        Loop
        Do While AllClass.РейсыКлиента Is Nothing
            mo.РейсыКлиентаAll()
        Loop


        Dim f1 = AllClass.ФайлыExcelВсе.Where(Function(x) x.Груз IsNot Nothing).Select(Function(x) x).ToList()


        For Each b In AllClass.РейсыКлиента
            f1.Add(New ФайлыExcelВсе With {.Клиент = b.НазвОрганизации, .Маршрут = b.Маршрут, .ДатаЗагрузки = b.ДатаПодачиПодЗагрузку,
                   .ДатаПодРастаможку = b.ДатаПодачиПодРастаможку, .АдресЗагрузки = b.ТочныйАдресЗагрузки, .АдресЗатаможки = b.АдресЗатаможки, .АдресРастаможки = b.ТочнАдресРаста,
                   .АдресВыгрузки = b.ТочнАдресРазгр, .Авто = b.ТипТрСредства, .Груз = b.НаименованиеГруза, .ДопУсловия = b.ДопУсловия, .Рейс = b.НомерРейса & " от " & b.ДатаПодачиПодЗагрузку, .ID = b.Код})

        Next

        Dim f3 = f1.Where(Function(x) x.Груз IsNot Nothing).Select(Function(x) x).ToList()

        Dim f2 = f3.Where(Function(x) x.Груз.ToUpper.Contains(txt.ToUpper)).Select(Function(x) x).ToList()

        Dim f4 As New ОбщийПоиск(f2)
        f4.ShowDialog()

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        If ComboBox2.Text.Length = 0 Then
            MessageBox.Show("Выберите рейс для удаления!", Рик)
            Return
        End If
        If com2.Count = 0 Then
            MessageBox.Show("Нет данных для удаления!", Рик)
            Return
        End If
        Dim f As Com2ЖурналДобавитьГрузClass = ComboBox2.SelectedItem
        If MessageBox.Show("Удалить груз - " & f.Naz & "?", Рик, MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.No Then
            Return
        End If

        DelГрузAsync(f.КодЖурнГруз)
        com2.Remove(f)
        bscom2.ResetBindings(False)
        MessageBox.Show("Данные удалены!", Рик)
        ComboBox2.Text = String.Empty
        Grid1All.Clear()
        bsGrid1.ResetBindings(False)
        Grid2All.Clear()
        bsGrid2.ResetBindings(False)
        MaskedTextBox1.Text = Nothing
        MaskedTextBox2.Text = Nothing
    End Sub
    Private Async Sub DelГрузAsync(ByVal d As Integer)
        Await Task.Run(Sub() DelГруз(d))
    End Sub
    Private Sub DelГруз(ByVal d As Integer)
        Using db As New dbAllDataContext()
            Dim f = db.ЖурналКлиентГруз.Where(Function(x) x.Код = d).FirstOrDefault()
            If f IsNot Nothing Then
                db.ЖурналКлиентГруз.DeleteOnSubmit(f)
                db.SubmitChanges()
                Dim mo As New AllUpd
                mo.ЖурналКлиентГрузAll()
            End If
        End Using
    End Sub
    Public Property DGVsender As DataGridView

    Private Sub Grid2_MouseDown(sender As Object, e As MouseEventArgs) Handles Grid2.MouseDown
        If e.Button = MouseButtons.Right Then
            ContextMenuStrip1.Show(MousePosition, ToolStripDropDownDirection.Right)
            DGVsender = TryCast(sender, DataGridView)
        End If


    End Sub

    Private Sub УдалитьСтрокуToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles УдалитьСтрокуToolStripMenuItem.Click


        If MessageBox.Show("Удалить строку?", Рик, MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.No Then
            Return
        End If

        Grid2All.RemoveAt(DGVsender.CurrentRow.Index)
        bsGrid2.ResetBindings(False)
    End Sub

    Private Sub Grid1_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles Grid1.CellEndEdit
        If e.ColumnIndex = 1 Then
            Grid1All.ElementAt(e.RowIndex).Наименование = Grid1All.ElementAt(e.RowIndex).Наименование & " - (" & Now.ToLongTimeString & ")"
        End If
    End Sub

    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox2.CheckedChanged
        'дЛина
        If CheckBox2.Checked = True Then
            Grid1.Columns(4).Visible = True
        Else
            Grid1.Columns(4).Visible = False
        End If
    End Sub

    Private Sub CheckBox3_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox3.CheckedChanged
        If CheckBox3.Checked = True Then
            Grid1.Columns(5).Visible = True
        Else
            Grid1.Columns(5).Visible = False
        End If
    End Sub

    Private Sub CheckBox4_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox4.CheckedChanged
        If CheckBox4.Checked = True Then 'высота
            Grid1.Columns(6).Visible = True
        Else
            Grid1.Columns(6).Visible = False
        End If
    End Sub

    Private Sub CheckBox5_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox5.CheckedChanged
        If CheckBox5.Checked = True Then
            Grid1.Columns(3).Visible = True
        Else
            Grid1.Columns(3).Visible = False
        End If
    End Sub

    Private Sub CheckBox6_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox6.CheckedChanged
        If CheckBox6.Checked = True Then
            Grid1.Columns(8).Visible = True
        Else
            Grid1.Columns(8).Visible = False
        End If
    End Sub

    Private Sub CheckBox7_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox7.CheckedChanged
        If CheckBox7.Checked = True Then
            Grid1.Columns(9).Visible = True
        Else
            Grid1.Columns(9).Visible = False
        End If
    End Sub

    Private Sub CheckBox8_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox8.CheckedChanged
        If CheckBox8.Checked = True Then
            Grid1.Columns(10).Visible = True
        Else
            Grid1.Columns(10).Visible = False
        End If
    End Sub

    Private Sub CheckBox9_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox9.CheckedChanged
        If CheckBox9.Checked = True Then
            Grid1.Columns(12).Visible = True
        Else
            Grid1.Columns(12).Visible = False
        End If
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Dim f As New ЖурналДобавитьГрузДопДанные
        f.ShowDialog()
        Rezult = f.Rezult
    End Sub
End Class