﻿Imports System.ComponentModel
Imports System.Threading

Public Class Журнал
    'Inherits MetroFramework.Forms.MetroForm

    Private comALLS As List(Of Com1ЖурналClass) 'общие данные главный лист

    Private com1All As List(Of IDNaz)
    Private bscom1 As BindingSource
    Private com2All As List(Of IDNaz)
    Private bscom2 As BindingSource
    Private com5All As List(Of Com1ЖурналClass)
    Private bscom5 As BindingSource
    Private Grid2all As List(Of Grid2ЖурналClass)
    Private bsgrid2 As BindingSource
    Private listbox1All As List(Of String)
    Private bslistbox1 As BindingSource
    Private Grid1all As List(Of Grid1ЖурналClass)
    Private bsgrid1 As BindingSource
    Private Grid3all As List(Of Grid3ЖурналClass)
    Private Grid3sel As BindingList(Of Grid3ЖурналClass)
    Private Grid3selAll As List(Of Grid3ЖурналClass)
    Private bsgrid3 As BindingSource
    Private Grid4all As BindingList(Of Grid3ЖурналClass)
    Private bsgrid4 As BindingSource
    Private Delegate Sub grd1()
    Private com3All As List(Of IDNaz)
    Private bscom3 As BindingSource
    Private com4All As List(Of IDNaz)
    Private bscom4 As BindingSource

    Sub New()

        ' Этот вызов является обязательным для конструктора.
        InitializeComponent()

        Grid4all = New BindingList(Of Grid3ЖурналClass)
        bsgrid4 = New BindingSource
        bsgrid4.DataSource = Grid4all
        Grid4.DataSource = bsgrid4
        GridView(Grid4)
        Grid4.Columns(1).Visible = False
        Grid4.Visible = False
        Grid4.DefaultCellStyle.Font = New Font("Calibri", 10)

        ' Добавить код инициализации после вызова InitializeComponent().

    End Sub

    Private Sub Grids()
        Grid1all = New List(Of Grid1ЖурналClass)
        bsgrid1 = New BindingSource
        bsgrid1.DataSource = Grid1all
        Grid1.DataSource = bsgrid1
        GridView(Grid1)
        Grid1.Columns(0).ReadOnly = True
        Grid1.Columns(1).ReadOnly = True
        Grid1.Columns(2).ReadOnly = True
        Grid1.Columns(3).ReadOnly = True
        Grid1.Columns(3).Width = 60
        Grid1.Columns(8).Visible = False
        Grid1.Columns(9).Visible = False
        Grid1.Columns(10).Visible = False
        CheckBox1.Checked = True
        Grid1.DefaultCellStyle.Font = New Font("Calibri", 10)


        Grid3selAll = New List(Of Grid3ЖурналClass)
        Grid3sel = New BindingList(Of Grid3ЖурналClass)
        bsgrid3 = New BindingSource
        bsgrid3.DataSource = Grid3sel
        Grid3.DataSource = bsgrid3
        GridView(Grid3)
        Grid3.Columns(1).Visible = False
        Grid3.Visible = False
        Grid3all = New List(Of Grid3ЖурналClass)
        Grid3.DefaultCellStyle.Font = New Font("Calibri", 10)

        'Grid4all = New BindingList(Of Grid3ЖурналClass)
        'bsgrid4 = New BindingSource
        'bsgrid4.DataSource = Grid4all
        'Grid4.DataSource = bsgrid4
        'GridView(Grid4)
        'Grid4.Columns(1).Visible = False
        'Grid4.Visible = False



        Do While AllClass.ПеревозчикиБаза Is Nothing
            Dim mo As New AllUpd
            mo.ПеревозчикиБазаAll()
        Loop
        Dim i As Integer = 1
        Dim f = AllClass.ПеревозчикиБаза.OrderBy(Function(x) x.Наименование_фирмы).Select(Function(x) New Grid3ЖурналClass With
                                                                                              {.КонтДанные = x.Контактное_лицо & vbCrLf & x.Телефоны,
                                                                                              .Перевозчик = x.Наименование_фирмы, .Страны = x.Страны_перевозок,
                                                                                              .Регионы = x.Регионы, .Транспорт = x.Вид_авто & ":" & vbCrLf & "-" &
                                                                                              x.Тоннаж & " т." & vbCrLf & "-" & x.Объем & " м3." & vbCrLf & "Кол: - " & x.Кол_во_авто & vbCrLf &
                                                                                              "ADR: - " & x.ADR, .Примечание = x.Примечание, .КодПеревозчика = x.ID}).ToList()
        Cursor = Cursors.WaitCursor

        If f IsNot Nothing Then
            If f.Count > 0 Then
                For Each b In f
                    b.Номер = i
                    i += 1
                    'Grid4all.Add(b)

                Next
                Grid3all.AddRange(f)



            End If
        End If
        Cursor = Cursors.Default

    End Sub
    Public Async Sub grdasync()
        Await Task.Run(Sub() grd())
    End Sub

    Private Sub grd()
        If Grid4.InvokeRequired Then
            Me.Invoke(New grd1(AddressOf grd))
        Else
            For Each b In Grid3all
                Grid4all.Add(b)
            Next
        End If
    End Sub
    Private Sub Loaded()
        Me.MdiParent = MDIParent1
        Label1.Text = Now.Date
        ПредзагрузкаAsync()

        Grid2Met()

        Grids()

        comALLS = New List(Of Com1ЖурналClass)
        com1All = New List(Of IDNaz)
        bscom1 = New BindingSource
        bscom1.DataSource = com1All
        ComboBox1.DataSource = bscom1
        ComboBox1.DisplayMember = "Naz"
        ComboBox1.Text = String.Empty


        com3All = New List(Of IDNaz)
        bscom3 = New BindingSource
        bscom3.DataSource = com3All
        ComboBox3.DataSource = bscom3
        ComboBox3.DisplayMember = "Naz"
        ComboBox3.Text = String.Empty

        com4All = New List(Of IDNaz)
        bscom4 = New BindingSource
        bscom4.DataSource = com4All
        ComboBox4.DataSource = bscom4
        ComboBox4.DisplayMember = "Naz"
        ComboBox4.Text = String.Empty


        com2All = New List(Of IDNaz)
        bscom2 = New BindingSource
        bscom2.DataSource = com2All
        ComboBox2.DataSource = bscom2
        ComboBox2.DisplayMember = "Naz"
        ComboBox2.Text = String.Empty

        com5All = New List(Of Com1ЖурналClass)
        bscom5 = New BindingSource
        bscom5.DataSource = com5All
        ComboBox5.DataSource = bscom5
        ComboBox5.DisplayMember = "Naz"
        ComboBox5.Text = String.Empty

        listbox1All = New List(Of String)
        bslistbox1 = New BindingSource
        bslistbox1.DataSource = listbox1All
        ListBox1.DataSource = bslistbox1

        ComboBox4.BeginInvoke(New MethodInvoker(Sub() Com4metod())) 'поток параллельно с контролами

        grdasync()

        ComboBox1.BeginInvoke(New MethodInvoker(Sub() Osnowa())) 'поток параллельно с контролами




        Grid1.Columns(7).Visible = False
        Grid1.Columns(0).Width = 60
        Grid3.Columns(0).Width = 60
        Grid4.Columns(0).Width = 60
        Grid4.Columns(1).MinimumWidth = 150

    End Sub
    Private Sub Com4metod()
        For i As Integer = 1 To 12
            Dim n As New IDNaz With {.ID = i, .Naz = MonthName(i).ToString()}
            com4All.Add(n)
        Next
        bscom4.ResetBindings(False)
        ComboBox4.Text = String.Empty
    End Sub



    Private Sub Журнал_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Loaded()


    End Sub
    Private Sub Osnowa()
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


        Do While AllClass.ЖурналКлиентСписок Is Nothing
            mo.ЖурналКлиентСписокAll()
        Loop

        Do While AllClass.ЖурналПеревозчик Is Nothing
            mo.ЖурналПеревозчикAll()
        Loop




        Dim i As Integer = 1
        Dim f1 = (From b In ((From x In AllClass.ЖурналКлиентГруз
                              Join y In AllClass.ЖурналКлиентМаршрут On x.Код Equals y.КодЖурналКлиентГруз
                              Join z In AllClass.ЖурналДата On z.Код Equals x.КодЖурналДата
                              Where x.Экспедитор = Экспедитор
                              Select x, y, z).ToList())
                  Group b By Keys = New With {Key b.x.Код, Key b.z.Дата,
                     Key .P = b.z.Дата & ": " & IIf(b.x.НаименованиеГруза Is Nothing, String.Empty, b.x.НаименованиеГруза) & ". Место загрузки: " & b.y.ГородПогрузки & " (" & b.y.СтранаПогрузки & " - " & b.y.СтранаВыгрузки & ")", Key b.x.Клиент}
                      Into Group
                  Select New Com1ЖурналClass With {.СписокДат = Keys.Дата, .КодГруза = Keys.Код, .Организации = Keys.Клиент, .Naz = Keys.P, .Наименование = Group(0).x.НаименованиеГруза, .Вес = Group(0).x.Вес,
        .Обьем = Group(0).x.Обьем, .Длина = Group(0).x.Длина, .Ширина = Group(0).x.Ширина, .Высота = Group(0).x.Высота, .Результат = Group(0).x.РезультатРаботы, .ДатаРезультата = IIf(Group(0)?.x?.ДатаРезультата Is Nothing, String.Empty, Group(0)?.x?.ДатаРезультата),
            .ТипПогрузки = Group(0).x.ТипПогрузки, .ПаллетыШтук = Group(0).x.ПаллетыШтук, .РазмерПаллет = Group(0).x.РазмерПаллет, .ADR = Group(0).x.ADR,
                     .ТипАвто = Group(0).x.ТипАвто, .ДополнитИнформация1 = Group(0).x.ДополнитИнформация, .ДатаЗагрузки = IIf(Group(0).x.ДатаЗагрузки Is Nothing, Nothing, Group(0).x.ДатаЗагрузки), .ДатаВыгрузки = IIf(Group(0).x.ДатаВыгрузки Is Nothing, Nothing, Group(0).x.ДатаВыгрузки),
                      .Lst = (From n In Group
                              Select n.y).ToList()}).ToList()



        Dim f2 = (From x In AllClass.ЖурналДата
                  Join y In AllClass.ЖурналКлиентГруз On x.Код Equals y.КодЖурналДата
                  Where y.Экспедитор = Экспедитор
                  Order By x.Дата Descending
                  Select x.Дата).Distinct().ToList()
        For Each b In f2
            Dim mk As New IDNaz With {.Naz = b}
            com1All.Add(mk)
        Next

        Dim m As New List(Of Com1ЖурналClass)

        Dim f10 = f1.Select(Function(x) x.Организации).Distinct().ToList() 'выбираем без повтора

        For Each f3 In f10 'перебираем в двух разных таблицах
            m.AddRange(AllClass.Клиент.Where(Function(x) x.НазваниеОрганизации = f3).Select(Function(x) New Com1ЖурналClass With {.КонтТелефон = x.Телефон, .КонтЛицо = x.Контактное_лицо, .Организации = x.НазваниеОрганизации}).ToList())
            m.AddRange(AllClass.ЖурналКлиентСписок.Where(Function(x) x.Клиент = f3).Select(Function(x) New Com1ЖурналClass With {.КонтТелефон = x.Телефон, .КонтЛицо = x.КонтактноеЛицо, .Организации = x.Клиент}).ToList())
        Next

        If m IsNot Nothing Then 'вставляем в большую таблицу
            If m.Count > 0 Then
                For Each f9 In m
                    Dim ml = f1.Where(Function(x) x.Организации = f9.Организации).Select(Function(x) x).ToList()
                    If ml IsNot Nothing Then
                        If ml.Count > 0 Then
                            For Each b In ml
                                b.КонтТелефон = f9.КонтТелефон
                                b.КонтЛицо = f9.КонтЛицо
                            Next

                        End If
                    End If
                Next
            End If

        End If


        If f1 IsNot Nothing Then

            If f1.Count > 0 Then
                comALLS.AddRange(f1)
            End If

        End If
        bscom1.ResetBindings(False)
        ComboBox1.Text = String.Empty

        Dim f5 = (From x In AllClass.ЖурналКлиентГруз
                  Where x.Экспедитор = Экспедитор
                  Order By x.Клиент
                  Select x.Клиент).Distinct().ToList()
        If f5 IsNot Nothing Then
            For Each b In f5
                Dim f As New IDNaz With {.Naz = b}
                com3All.Add(f)
            Next
            bscom3.ResetBindings(False)
            ComboBox3.Text = String.Empty
        End If


    End Sub
    Private Sub Grid2Met()
        Grid2all = New List(Of Grid2ЖурналClass)
        bsgrid2 = New BindingSource
        bsgrid2.DataSource = Grid2all
        Grid2.DataSource = bsgrid2
        GridView(Grid2)
        Grid2.Columns(0).Width = 50
        Grid2.Columns(0).HeaderText = "№"
        Grid2.Columns(2).Width = 80
        Grid2.Columns(3).Width = 80
        Grid2.Columns(6).Visible = False
        Grid2.Columns(7).Visible = False
        Grid2.Columns(8).Visible = False

        Grid2.Columns(3).HeaderText = "Дата загрузки"

        Dim mo As New AllUpd
        Do While AllClass.ЖурналДата Is Nothing
            mo.ЖурналДатаAll()
        Loop
        Do While AllClass.ЖурналКлиентГруз Is Nothing
            mo.ЖурналКлиентГрузAll()
        Loop
        Do While AllClass.ЖурналКлиентМаршрут Is Nothing
            mo.ЖурналКлиентМаршрутAll()
        Loop

        Dim f = (From x In AllClass.ЖурналДата
                 Join y In AllClass.ЖурналКлиентГруз On x.Код Equals y.КодЖурналДата
                 Join z In AllClass.ЖурналКлиентМаршрут On y.Код Equals z.КодЖурналКлиентГруз
                 Where y.Экспедитор = Экспедитор And y.ОтоброжатьВТаблицеЖурнала Is Nothing
                 Order By x.Дата Descending
                 Select x, y, z).ToList

        Dim f1 = (From x In f
                  Group x By Keys = New With {Key x.y.Клиент, Key x.x.Дата, Key x.y.ДатаЗагрузки, Key x.y.НаименованиеГруза}
                      Into Group
                  Select New Grid2ЖурналClass With {.Клиент = Keys.Клиент, .СтранаПогрузки = Group(0).z.СтранаПогрузки, .СтранаВыгрузки = Group(0).z.СтранаВыгрузки, .КодЖурналГруз = Group(0).y.Код, .Дата = Keys.Дата, .ДатаЗагрузки = IIf(Keys.ДатаЗагрузки Is Nothing, Nothing, Keys.ДатаЗагрузки), .Груз = Keys.НаименованиеГруза,
                     .МаршрутList = (From p In Group
                                     Select "( " & p.z.СтранаПогрузки & " ) " & p.z.ГородПогрузки & " - " & "( " & p.z.СтранаВыгрузки & " ) " & p.z.ГородВыгрузки).ToList()}).ToList()



        If Grid2all IsNot Nothing Then
            Grid2all.Clear()
        End If
        Dim i As Integer = 1
        For Each b In f1
            Dim k As String = Nothing
            For Each b1 In b.МаршрутList
                If k Is Nothing Then
                    k = b1
                Else
                    k = k & vbCrLf & b1
                End If

            Next
            Dim f3 As New Grid2ЖурналClass With {.Дата = b.Дата, .ДатаЗагрузки = b.ДатаЗагрузки, .Клиент = b.Клиент, .Номер = i, .Mаршрут = k, .Груз = b.Груз, .КодЖурналГруз = b.КодЖурналГруз}

            i += 1
            Grid2all.Add(f3)
        Next

        bsgrid2.ResetBindings(False)

    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click

        Dim f As New ЖурналДобавитьГруз
        f.ShowDialog()
        If f.Flag = True Then
            Loaded()
        End If

        Grid2Met()
    End Sub
    Private Async Sub ПредзагрузкаAsync()
        Await Task.Run(Sub() Предзагрузка())
    End Sub
    Private Sub Предзагрузка()
        Dim mo As New AllUpd
        Do While AllClass.Клиент Is Nothing
            mo.КлиентAll()
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

        Do While AllClass.ЖурналКлиентСписок Is Nothing
            mo.ЖурналКлиентСписокAll()
        Loop
        Do While AllClass.ЖурналПеревозчик Is Nothing
            mo.ЖурналПеревозчикAll()
        Loop
        Do While AllClass.ПеревозчикиБаза Is Nothing
            mo.ПеревозчикиБазаAll()
        Loop
    End Sub

    Private Sub ComboBox1_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles ComboBox1.SelectionChangeCommitted
        Dim f2 As IDNaz = ComboBox1.SelectedItem
        Dim f = comALLS.Where(Function(x) x.СписокДат = f2.Naz).Select(Function(x) x.Организации).Distinct.ToList()
        If com2All IsNot Nothing Then
            com2All.Clear()
        End If

        If f IsNot Nothing Then
            If f.Count > 0 Then
                For Each b In f
                    Dim m As New IDNaz With {.Naz = b}
                    com2All.Add(m)
                Next
            End If
        End If
        bscom2.ResetBindings(False)
        ComboBox2.Text = String.Empty
    End Sub

    Private Sub ComboBox2_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles ComboBox2.SelectionChangeCommitted
        Dim f As IDNaz = ComboBox2.SelectedItem
        Dim f1 As IDNaz = ComboBox1.SelectedItem
        Dim f2 = comALLS.Where(Function(x) x.Организации = f.Naz And x.СписокДат = f1.Naz).Select(Function(x) x).ToList()
        If com5All IsNot Nothing Then
            com5All.Clear()
        End If

        If f2 IsNot Nothing Then
            If f2.Count > 0 Then
                com5All.AddRange(f2)
                bscom5.ResetBindings(False)
                ComboBox5.Text = String.Empty
            End If
        End If
    End Sub
    Private Sub comb5Sel()
        Cursor = Cursors.WaitCursor
        Dim f2 As Com1ЖурналClass
        f2 = ComboBox5.SelectedItem
        If f2 Is Nothing Then
            f2 = comALLS.Where(Function(x) x.Naz.ToUpper.Contains(ComboBox5.Text.ToUpper)).Select(Function(x) x).FirstOrDefault()
        End If


        If listbox1All IsNot Nothing Then
            listbox1All.Clear()
        End If


        ListBox1.BeginUpdate()
        'ListBox1.DrawMode = DrawMode.OwnerDrawFixed
        If f2 Is Nothing Then Return

        With f2
            listbox1All.Add("- - ДАННЫЕ ПО ГРУЗУ - -")
            listbox1All.Add("|Наименование груза: " & vbCrLf & .Наименование)
            listbox1All.Add("|Размеры груза: " & vbCrLf & "|Длина: " & .Длина & ", |Ширина: " & .Ширина & ", |Высота: " & .Высота)
            listbox1All.Add("|Дата загрузки: " & .ДатаЗагрузки & "|Дата выгрузки: " & .ДатаВыгрузки)
            listbox1All.Add("|Вес: " & .Вес & " кг." & "|Обем: " & .Обьем & " м3.")
            listbox1All.Add("|Кол.пал: " & .ПаллетыШтук & " шт., " & "|Размер.пал: " & .РазмерПаллет & " м.")
            listbox1All.Add("|СТАВКА: " & .Lst(0).Ставка)
            listbox1All.Add(vbCrLf & "___________________________________" & vbCrLf)
            listbox1All.Add("- - ДАННЫЕ ПО МЕСТУ ЗАГРУЗКИ - -")
            For Each b In f2.Lst
                listbox1All.Add("|Место загрузки: " & "(" & b.СтранаПогрузки & ") " & b.ГородПогрузки & " (" & b.КвадратПогрузки & ")")
                listbox1All.Add("|Место выгрузки: " & "(" & b.СтранаВыгрузки & ") " & b.ГородВыгрузки & " (" & b.КвадратВыгрузки & ")")
                listbox1All.Add("|Там.отпр: " & vbCrLf & b.ТаможняОтправления)
                listbox1All.Add("|Там.Назнач: " & vbCrLf & b.ТаможняНазначения)
                listbox1All.Add("|Доп.информация: " & b.ДополнитИнформация)
            Next
            listbox1All.Add(vbCrLf & "___________________________________" & vbCrLf)
            listbox1All.Add("- - ДАННЫЕ ПО КЛИЕНТУ - -" & vbCrLf)
            listbox1All.Add("| РЕЗУЛЬТАТ: " & .Результат & "| ДАТА: " & .ДатаРезультата)
            listbox1All.Add("|Конт.тел: " & .КонтТелефон & ", " & "|Конт.лицо: " & .КонтЛицо)


        End With


        bslistbox1.ResetBindings(False)

        ListBox1.EndUpdate()

        If Grid1all IsNot Nothing Then
            Grid1all.Clear()
            bsgrid1.ResetBindings(False)
        End If

        Dim i As Integer = 1
        Dim f3 = AllClass.ЖурналПеревозчик.Where(Function(x) x.КодЖурналКлиентГруз = f2.КодГруза).Select(Function(x) New Grid1ЖурналClass _
                                                                                                         With {.Дата = IIf(x.Дата Is Nothing, Nothing, x.Дата),
                                                                                                         .ДопИнформация = x.ДопИнформация, .Контакт = x.КонтДанные,
                                                                                                         .Перевозчик = x.Организация,
                                                                                                         .Состояние = x.Состояние, .Ставка = x.Ставкапервозчика,
                                                                                                         .КодПеревозчика = x.Кодперевозчик, .КодЖурналПеревозчик = x.Код,
                                                                                                         .Skype = x.Skype?.ToString, .SkypeDate = x.SkypeDate?.ToString}).ToList()
        For Each b In f3
            If b.Skype?.Length > 0 Then
                b.ДопИнформация = b.ДопИнформация & vbCrLf & "Skype +"
            End If
        Next

        If f3 IsNot Nothing Then
            If f3.Count > 0 Then
                For Each b In f3
                    b.Номер = i
                    i += 1
                Next
                Grid1all.AddRange(f3)
                bsgrid1.ResetBindings(False)
            End If
        End If






        If Grid3all IsNot Nothing Then 'заполняем grid3 таблицу по регионам
            Dim f5 As New List(Of Grid3ЖурналClass)
            For Each b In f2.Lst
                If b.СтранаПогрузки IsNot Nothing Then
                    f5.AddRange(Grid3all.Where(Function(x) x.Страны.ToUpper.Contains(b.СтранаПогрузки.ToUpper)).Select(Function(x) x).ToList())
                End If
                If b.КвадратПогрузки IsNot Nothing Then
                    f5.AddRange(Grid3all.Where(Function(x) x.Регионы.ToUpper.Contains(b.КвадратПогрузки.ToUpper)).Select(Function(x) x).ToList())
                End If

            Next
            Dim f7 = f5.OrderBy(Function(x) x.Перевозчик).Select(Function(x) x.КодПеревозчика).Distinct().ToList()
            If f7 IsNot Nothing Then
                If f7.Count > 0 Then
                    Grid3sel.Clear()
                    Dim i1 As Integer = 1
                    For Each b In f7
                        Dim k = Grid3all.Where(Function(x) x.КодПеревозчика = b).Select(Function(x) x).FirstOrDefault()
                        k.Номер = i1
                        Grid3sel.Add(k)
                        i1 += 1
                    Next
                    If Grid3selAll IsNot Nothing Then
                        Grid3selAll.Clear()
                    End If
                    Grid3selAll.AddRange(Grid3sel)
                End If

            End If
        End If

        Cursor = Cursors.Default
    End Sub

    Private Sub ComboBox5_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles ComboBox5.SelectionChangeCommitted
        comb5Sel()
    End Sub

    Private Sub Label8_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub TextBox4_TextChanged(sender As Object, e As EventArgs)

    End Sub

    Private Sub Button12_Click(sender As Object, e As EventArgs)

    End Sub



    Private Sub CheckBox3_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox3.CheckedChanged
        If CheckBox3.Checked = True Then
            Cursor = Cursors.WaitCursor
            Grid4.Visible = True
            CheckBox1.Checked = False
            CheckBox2.Checked = False
            Grid1.Visible = False
            Grid3.Visible = False
            Cursor = Cursors.Default
            Button11.Enabled = True
            Button12.Enabled = False
        Else
            Button11.Enabled = False
        End If
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked = True Then
            Grid1.Visible = True
            CheckBox3.Checked = False
            CheckBox2.Checked = False
            Grid3.Visible = False
            Grid4.Visible = False
            Button11.Enabled = False
            Button12.Enabled = False
        End If
    End Sub

    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox2.CheckedChanged
        If CheckBox2.Checked = True Then
            Grid3.Visible = True
            CheckBox3.Checked = False
            CheckBox1.Checked = False
            Grid1.Visible = False
            Grid4.Visible = False
            Button11.Enabled = False
            Button12.Enabled = True
        Else
            Button12.Enabled = False
        End If
    End Sub

    Private Sub Grid4_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles Grid4.CellDoubleClick
        If e.RowIndex = -1 Then Return
        Dim f3 As Grid3ЖурналClass = Grid4all.ElementAt(e.RowIndex)

        If MessageBox.Show("Добавить перевозчика в список выбранных?", Рик, MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.No Then
            Return
        End If

        If Grid1all IsNot Nothing Then
            Dim f1 = Grid1all.Where(Function(x) x.КодПеревозчика = f3.КодПеревозчика).Select(Function(x) x).FirstOrDefault()
            If f1 IsNot Nothing Then
                MessageBox.Show("Такой перевозчик уже есть в выбранных!", Рик)
                Return
            End If
        End If

        If String.IsNullOrEmpty(ComboBox5.Text) = True Then
            MessageBox.Show("Выберите 'Заказ клиента'!", Рик)
            Return
        End If



        If String.IsNullOrEmpty(ComboBox5.Text) = False Then
            Dim f5 As Com1ЖурналClass = ComboBox5.SelectedItem


            Dim f As New Grid1ЖурналClass With {.Дата = Now, .ДопИнформация = f3.Примечание, .КодПеревозчика = f3.КодПеревозчика, .Контакт = f3.КонтДанные, .Перевозчик = f3.Перевозчик}
            Dim id = Grid1AddToBase(f, f5.КодГруза)
            f.КодЖурналПеревозчик = id
            Grid1all.Add(f)
            Dim i As Integer = 1
            For Each b In Grid1all
                b.Номер = i
                i += 1
            Next

            bsgrid1.ResetBindings(False)


        End If


    End Sub
    'Private Async Sub Grid1AddToBaseAsync(ByVal d As Grid1ЖурналClass, ByVal d1 As Integer)
    '    Await Task.Run(Sub() Grid1AddToBase(d, d1))

    'End Sub
    Private Function Grid1AddToBase(ByVal d As Grid1ЖурналClass, ByVal d1 As Integer) As Integer
        Using db As New dbAllDataContext()
            Dim f As New ЖурналПеревозчик
            With f
                .Дата = d.Дата
                .ДопИнформация = d.ДопИнформация
                .Кодперевозчик = d.КодПеревозчика
                .КонтДанные = d.Контакт
                .Организация = d.Перевозчик
                .КодЖурналКлиентГруз = d1
            End With
            db.ЖурналПеревозчик.InsertOnSubmit(f)
            db.SubmitChanges()
            Dim mo As New AllUpd
            mo.ЖурналПеревозчикAllAsync()
            Return f.Код
        End Using
    End Function
    Private Sub ПоискВсе(ByVal kl As String)
        Dim f6 As New List(Of Grid3ЖурналClass)
        f6.AddRange(Grid3all.Where(Function(x) x.КонтДанные.ToUpper.Contains(kl.ToUpper)).Select(Function(x) x).ToList())
        f6.AddRange(Grid3all.Where(Function(x) x.Перевозчик.ToUpper.Contains(kl.ToUpper)).Select(Function(x) x).ToList())
        f6.AddRange(Grid3all.Where(Function(x) x.Примечание.ToUpper.Contains(kl.ToUpper)).Select(Function(x) x).ToList())
        f6.AddRange(Grid3all.Where(Function(x) x.Регионы.ToUpper.Contains(kl.ToUpper)).Select(Function(x) x).ToList())
        f6.AddRange(Grid3all.Where(Function(x) x.Страны.ToUpper.Contains(kl.ToUpper)).Select(Function(x) x).ToList())
        f6.AddRange(Grid3all.Where(Function(x) x.Транспорт.ToUpper.Contains(kl.ToUpper)).Select(Function(x) x).ToList())

        Dim f7 = f6.OrderBy(Function(x) x.Перевозчик).Select(Function(x) x.КодПеревозчика).Distinct().ToList()
        If f7 IsNot Nothing Then
            Dim f9 As New List(Of Grid3ЖурналClass)
            For Each b In f7
                Dim f8 = Grid3all.Where(Function(x) x.КодПеревозчика = b).Select(Function(x) x).FirstOrDefault()
                f9.Add(f8)
            Next
            If f9 IsNot Nothing Then
                Grid4all.Clear()
                Dim i As Integer = 1
                For Each b In f9
                    b.Номер = i
                    Grid4all.Add(b)
                    i += 1
                Next
            End If
        End If



    End Sub
    Private Sub ПоискРегионы(ByVal kl As String)
        Dim f6 As New List(Of Grid3ЖурналClass)
        f6.AddRange(Grid3selAll.Where(Function(x) x.КонтДанные.ToUpper.Contains(kl.ToUpper)).Select(Function(x) x).ToList())
        f6.AddRange(Grid3selAll.Where(Function(x) x.Перевозчик.ToUpper.Contains(kl.ToUpper)).Select(Function(x) x).ToList())
        f6.AddRange(Grid3selAll.Where(Function(x) x.Примечание.ToUpper.Contains(kl.ToUpper)).Select(Function(x) x).ToList())
        f6.AddRange(Grid3selAll.Where(Function(x) x.Регионы.ToUpper.Contains(kl.ToUpper)).Select(Function(x) x).ToList())
        f6.AddRange(Grid3selAll.Where(Function(x) x.Страны.ToUpper.Contains(kl.ToUpper)).Select(Function(x) x).ToList())
        f6.AddRange(Grid3selAll.Where(Function(x) x.Транспорт.ToUpper.Contains(kl.ToUpper)).Select(Function(x) x).ToList())

        Dim f7 = f6.OrderBy(Function(x) x.Перевозчик).Select(Function(x) x.КодПеревозчика).Distinct().ToList()
        If f7 IsNot Nothing Then
            Dim f9 As New List(Of Grid3ЖурналClass)
            For Each b In f7
                Dim f8 = Grid3selAll.Where(Function(x) x.КодПеревозчика = b).Select(Function(x) x).FirstOrDefault()
                f9.Add(f8)
            Next
            If f9 IsNot Nothing Then
                Grid3sel.Clear()
                Dim i As Integer = 1
                For Each b In f9
                    b.Номер = i
                    Grid3sel.Add(b)
                    i += 1
                Next
            End If
        End If



    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        If String.IsNullOrEmpty(TextBox2.Text) Then Return
        Dim kl As String = TextBox2.Text

        If CheckBox3.Checked = True Then
            ПоискВсе(kl)
        End If

        If CheckBox2.Checked = True Then
            If Grid3sel IsNot Nothing Then
                If Grid3sel.Count > 0 Then
                    ПоискРегионы(kl)
                End If

            End If
        End If
        TextBox2.Text = String.Empty
    End Sub



    Private Sub Grid3_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles Grid3.CellDoubleClick

        If e.RowIndex = -1 Then Return
        Dim f3 As Grid3ЖурналClass = Grid3sel.ElementAt(e.RowIndex)

        If MessageBox.Show("Добавить перевозчика в список выбранных?", Рик, MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.No Then
            Return
        End If

        If Grid1all IsNot Nothing Then
            Dim f1 = Grid1all.Where(Function(x) x.КодПеревозчика = f3.КодПеревозчика).Select(Function(x) x).FirstOrDefault()
            If f1 IsNot Nothing Then
                MessageBox.Show("Такой перевозчик уже есть в выбранных!", Рик)
                Return
            End If
        End If

        If String.IsNullOrEmpty(ComboBox5.Text) = True Then
            MessageBox.Show("Выберите 'Заказ клиента'!", Рик)
            Return
        End If



        If String.IsNullOrEmpty(ComboBox5.Text) = False Then
            Dim f5 As Com1ЖурналClass = ComboBox5.SelectedItem



            Dim f As New Grid1ЖурналClass With {.Дата = Now, .ДопИнформация = f3.Примечание, .КодПеревозчика = f3.КодПеревозчика, .Контакт = f3.КонтДанные, .Перевозчик = f3.Перевозчик}
            Dim id = Grid1AddToBase(f, f5.КодГруза)
            f.КодЖурналПеревозчик = id
            Grid1all.Add(f)
            Dim i As Integer = 1
            For Each b In Grid1all
                b.Номер = i
                i += 1
            Next
            bsgrid1.ResetBindings(False)

        End If
    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        If Grid3all IsNot Nothing Then
            Grid4all.Clear()
            Dim i As Integer = 1
            For Each b In Grid3all
                b.Номер = i
                Grid4all.Add(b)
                i += 1
            Next
        End If
    End Sub

    Private Sub Button12_Click_1(sender As Object, e As EventArgs) Handles Button12.Click
        If Grid3selAll IsNot Nothing Then
            If Grid3selAll.Count > 0 Then
                Grid3sel.Clear()
                Dim i As Integer = 1
                For Each b In Grid3selAll
                    b.Номер = i
                    Grid3sel.Add(b)
                    i += 1
                Next
            End If
        End If
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        If String.IsNullOrEmpty(TextBox1.Text) Then
            MessageBox.Show("Заполните поле поиска!", Рик)
            Return
        End If
        Dim txt1 As String = TextBox1.Text
        Dim f3 As New List(Of ForSeachtxt1)
        If comALLS IsNot Nothing Then



            If comALLS.Count > 0 Then
                For Each b In comALLS

                    Dim lstt As String = Nothing
                    For Each b1 In b.Lst
                        Dim var2 = String.Join(",", b1.ГородВыгрузки, b1.ГородВыгрузки, b1.ДополнитИнформация, b1.КвадратВыгрузки, b1.КвадратПогрузки, b1.Клиент, b1.Ставка, b1.СтранаВыгрузки,
b1.СтранаПогрузки, b1.ТаможняНазначения, b1.ТаможняОтправления)
                        lstt &= var2
                    Next

                    Dim var = String.Join(",", (b.ADR, b.EX, b.Высота, b.Вес, b.ГородВыгрузки, b.ГородПогрузки, b.ДатаВыгрузки, b.ДатаЗагрузки, b.Длина, b.ДополнитИнформация, b.ДополнитИнформация1, b.КвадратВыгрузки, lstt,
                                          b.КвадратПогрузки, b.КонтЛицо, b.КонтТелефон, b.Наименование, b.Обьем, b.Организации, b.ПаллетыШтук, b.РазмерПаллет, b.СписокДат, b.Ставка, b.СтранаВыгрузки, b.СтранаПогрузки, b.ТаможняНазначения, b.ТаможняОтправления,
                                          b.ТипАвто, b.ТипПогрузки, b.Ширина))

                    If var.ToUpper.Contains(txt1.ToUpper) Then
                        Dim ml As New ForSeachtxt1 With {.Данные = b, .Строка = var}
                        f3.Add(ml)
                    End If
                Next


            End If
        End If


        If f3 IsNot Nothing Then
            If f3.Count > 0 Then
                Dim f9 As New ЖурналДопФормаВыбора(f3)
                f9.ShowDialog()
                Dim f11 = f9.Sel
                If f11 IsNot Nothing Then
                    Dim mo As New AllUpd
                    Do While AllClass.ЖурналДата Is Nothing
                        mo.ЖурналДатаAll()
                    Loop
                    Do While AllClass.ЖурналКлиентГруз Is Nothing
                        mo.ЖурналКлиентГрузAll()
                    Loop
                    Dim f12 = (From x In AllClass.ЖурналДата
                               Join y In AllClass.ЖурналКлиентГруз On x.Код Equals y.КодЖурналДата
                               Join z In AllClass.ЖурналКлиентМаршрут On y.Код Equals z.КодЖурналКлиентГруз
                               Where y.Код = f11.КодЖурналГруз
                               Select x, y, z).FirstOrDefault()


                    ComboBox1.Text = f12.x.Дата
                    ComboBox2.Text = f12.y.Клиент
                    ComboBox5.Text = f12.x.Дата & ": " & f12.y.НаименованиеГруза & " (" & f12.z.СтранаПогрузки & " - " & f12.z.СтранаВыгрузки & ")"
                    comb5Sel()
                End If
            End If
        End If





    End Sub
    Public Class ForSeachtxt1
        Public Property Строка As String
        Public Property Данные As Com1ЖурналClass
    End Class

    Private Sub Grid2_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles Grid2.CellDoubleClick
        If e.RowIndex = -1 Then Return
        Dim f = Grid2all.ElementAt(e.RowIndex)
        If f Is Nothing Then Return
        Dim k As New ЖурналВсплывДанныеОбщиеДляГрид2(f)
        k.ShowDialog()


    End Sub

    Private Sub Grid1_UserDeletingRow(sender As Object, e As DataGridViewRowCancelEventArgs) Handles Grid1.UserDeletingRow
        Dim f As Grid1ЖурналClass = Grid1all.ElementAt(e.Row.Index)
        Dim f5 As Com1ЖурналClass = ComboBox5.SelectedItem
        Using db As New dbAllDataContext()
            Dim f1 = db.ЖурналПеревозчик.Where(Function(x) x.КодЖурналКлиентГруз = f5.КодГруза And x.Кодперевозчик = f.КодПеревозчика).Select(Function(x) x).FirstOrDefault()
            If f1 IsNot Nothing Then
                db.ЖурналПеревозчик.DeleteOnSubmit(f1)
                db.SubmitChanges()
                Dim mo As New AllUpd
                mo.ЖурналКлиентГрузAllAsync()
            End If
        End Using


    End Sub

    Private Sub Grid1_UserDeletedRow(sender As Object, e As DataGridViewRowEventArgs) Handles Grid1.UserDeletedRow
        Dim i As Integer = 1
        For Each b In Grid1all
            b.Номер = i
            i += 1
        Next
    End Sub

    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click
        ДобПер = 1
        Dim f As New ДобавитьПеревозчика()
        f.ShowDialog()
    End Sub
    Public Property Grid1Clic As Grid1ЖурналClass = Nothing

    Private Sub Grid1_MouseDown(sender As Object, e As MouseEventArgs) Handles Grid1.MouseDown
        If e.Button = MouseButtons.Right Then
            ContextMenuStrip1.Show(MousePosition, ToolStripDropDownDirection.Right)
            Dim k = TryCast(sender, DataGridView)
            Grid1Clic = Grid1all.ElementAt(k.CurrentCell.RowIndex)
        End If
    End Sub

    Private Sub ОткрытьДанныеПеревозчикаToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ОткрытьДанныеПеревозчикаToolStripMenuItem.Click
        Dim f As Grid1ЖурналClass = Grid1all.ElementAt(Grid1.CurrentRow.Index)
        Dim mo As New AllUpd
        Do While AllClass.ПеревозчикиБаза Is Nothing
            mo.ПеревозчикиБазаAll()
        Loop
        If f IsNot Nothing Then
            Dim m As New ИзменПеревозчика(f)
            m.ShowDialog()


        End If
    End Sub
    Private Async Sub addРезультAsync(ByVal d As Integer, ByVal rezult As String)
        Await Task.Run(Sub() addРезульт(d, rezult))
        comb5Sel()
    End Sub
    Private Sub addРезульт(ByVal d As Integer, ByVal rezult As String)
        Dim mo As New AllUpd
        Using db As New dbAllDataContext
            Dim f = db.ЖурналКлиентГруз.Where(Function(x) x.Код = d).FirstOrDefault()
            f.РезультатРаботы = rezult
            f.ДатаРезультата = Now.ToShortDateString
            db.SubmitChanges()
        End Using
        mo.ЖурналКлиентГрузAll()
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Dim f1 As IDNaz = ComboBox1.SelectedItem
        Dim f2 As IDNaz = ComboBox2.SelectedItem
        Dim f5 As Com1ЖурналClass = ComboBox5.SelectedItem


        Dim f = (From x In comALLS
                 Where x.СписокДат = f1.Naz And x.Организации = f2.Naz And x.Naz = f5.Naz
                 Select x.КодГруза).FirstOrDefault()
        If f > 0 Then
            addРезультAsync(f, "КЛИЕНТ САМ ЗАКРЫЛ!")
        End If
        MessageBox.Show("Данные приняты!")

        Dim f7 = listbox1All.Where(Function(x) x.Contains("РЕЗУЛЬТАТ:")).FirstOrDefault()

        If f7 IsNot Nothing Then
            Dim f8 = listbox1All.IndexOf(f7)
            f7 = "| РЕЗУЛЬТАТ: " & "КЛИЕНТ САМ ЗАКРЫЛ!" & "| ДАТА: " & Now.ToShortDateString
            listbox1All(f8) = f7
            ListBox1.BeginUpdate()
            bslistbox1.ResetBindings(False)
            ListBox1.EndUpdate()
        End If




    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Dim f1 As IDNaz = ComboBox1.SelectedItem
        Dim f2 As IDNaz = ComboBox2.SelectedItem
        Dim f5 As Com1ЖурналClass = ComboBox5.SelectedItem

        Dim f = (From x In comALLS
                 Where x.СписокДат = f1.Naz And x.Организации = f2.Naz And x.Naz = f5.Naz
                 Select x.КодГруза).FirstOrDefault()
        If f > 0 Then
            addРезультAsync(f, "КЛИЕНТ ОТМЕНИЛ!")
        End If
        MessageBox.Show("Данные приняты!")

        Dim f7 = listbox1All.Where(Function(x) x.Contains("РЕЗУЛЬТАТ:")).FirstOrDefault()

        If f7 IsNot Nothing Then
            Dim f8 = listbox1All.IndexOf(f7)
            f7 = "| РЕЗУЛЬТАТ: " & "КЛИЕНТ ОТМЕНИЛ!" & "| ДАТА: " & Now.ToShortDateString
            listbox1All(f8) = f7
            ListBox1.BeginUpdate()
            bslistbox1.ResetBindings(False)
            ListBox1.EndUpdate()
        End If
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Dim f1 As IDNaz = ComboBox1.SelectedItem
        Dim f2 As IDNaz = ComboBox2.SelectedItem
        Dim f5 As Com1ЖурналClass = ComboBox5.SelectedItem

        Dim f = (From x In comALLS
                 Where x.СписокДат = f1.Naz And x.Организации = f2.Naz And x.Naz = f5.Naz
                 Select x.КодГруза).FirstOrDefault()
        If f > 0 Then
            addРезультAsync(f, "СДЕЛКА СОСТОЯЛАСЬ!")
        End If
        MessageBox.Show("Данные приняты!")

        Dim f7 = listbox1All.Where(Function(x) x.Contains("РЕЗУЛЬТАТ:")).FirstOrDefault()

        If f7 IsNot Nothing Then
            Dim f8 = listbox1All.IndexOf(f7)
            f7 = "| РЕЗУЛЬТАТ: " & "СДЕЛКА СОСТОЯЛАСЬ!" & "| ДАТА: " & Now.ToShortDateString
            listbox1All(f8) = f7
            ListBox1.BeginUpdate()
            bslistbox1.ResetBindings(False)
            ListBox1.EndUpdate()
        End If
    End Sub
    Private Sub UpdЖурналПеревозчикиGrid1(ByVal d As Integer, ByVal Sob As Integer, ByVal txt As String)
        Dim mo As New AllUpd
        Using db As New dbAllDataContext()
            Dim f = db.ЖурналПеревозчик.Where(Function(x) x.Код = d).FirstOrDefault()
            If f IsNot Nothing Then
                Select Case Sob
                    Case 4
                        f.Состояние = txt
                        db.SubmitChanges()
                        mo.ЖурналПеревозчикAll()
                    Case 5
                        f.Ставкапервозчика = txt
                        db.SubmitChanges()
                        mo.ЖурналПеревозчикAll()
                    Case 6
                        f.ДопИнформация = txt
                        db.SubmitChanges()
                        mo.ЖурналПеревозчикAll()
                End Select
            End If
        End Using

    End Sub

    Private Sub Grid1_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles Grid1.CellValueChanged
        Dim f = Grid1all.ElementAt(e.RowIndex)
        Dim Rnd As New Random(245)
        If e.ColumnIndex = 4 Then 'состояние


            Dim thed As New Thread(New ThreadStart(Sub() UpdЖурналПеревозчикиGrid1(f.КодЖурналПеревозчик, 4, f.Состояние)))
            thed.Name = Rnd.Next.ToString
            thed.Start()
        End If

        If e.ColumnIndex = 5 Then 'ставка
            Dim thed As New Thread(New ThreadStart(Sub() UpdЖурналПеревозчикиGrid1(f.КодЖурналПеревозчик, 5, f.Ставка)))
            thed.Name = Rnd.Next.ToString
            thed.Start()
        End If

        If e.ColumnIndex = 6 Then 'ДопИнформация
            Dim thed As New Thread(New ThreadStart(Sub() UpdЖурналПеревозчикиGrid1(f.КодЖурналПеревозчик, 6, f.ДопИнформация)))
            thed.Name = Rnd.Next.ToString
            thed.Start()
        End If
    End Sub


    Private Sub ДобавитьToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ДобавитьToolStripMenuItem.Click
        If Grid1Clic IsNot Nothing Then
            Dim f As New Skype(Grid1Clic, True)
            f.ShowDialog()
        End If
    End Sub

    Private Sub ПрочитатьToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ПрочитатьToolStripMenuItem.Click
        If Grid1Clic IsNot Nothing Then
            Dim f As New Skype(Grid1Clic, False)
            f.ShowDialog()
        End If
    End Sub

    Private Sub ComboBox3_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles ComboBox3.SelectionChangeCommitted
        Dim f As IDNaz = ComboBox3.SelectedItem
        If ComboBox1.Text.Length > 0 And ComboBox2.Text.Length > 0 Then
            Dim f1 As New ЖурналCombo3СводнаяПоКлиентам(f.Naz)
            f1.ShowDialog()
        End If

    End Sub

    Private Sub Button15_Click(sender As Object, e As EventArgs) Handles Button15.Click
        Dim f As New SkypeClientPredloj
        f.ShowDialog()
    End Sub

    Private Sub Button14_Click(sender As Object, e As EventArgs) Handles Button14.Click
        Dim f As New SkypePerevozPredloj
        f.ShowDialog()
    End Sub

    Private Sub Button16_Click(sender As Object, e As EventArgs) Handles Button16.Click
        If My.Computer.Name.ToString = "OLEGLAPTOP" Then
            Dim f As New Календарь
            f.WindowState = FormWindowState.Normal
            f.StartPosition = FormStartPosition.CenterScreen
            f.ShowDialog()
        End If
    End Sub

    Private Sub Button17_Click(sender As Object, e As EventArgs) Handles Button17.Click
        Dim f9 As New ПоискПолный()
        f9.WindowState = FormWindowState.Normal
        f9.StartPosition = FormStartPosition.CenterScreen
        f9.ShowDialog()
    End Sub

    Private Sub Grid2_CellClick(sender As Object, e As DataGridViewCellEventArgs)

    End Sub
    Public Property Grid2Clic As Grid2ЖурналClass = Nothing
    Private Sub Grid2_MouseDown(sender As Object, e As MouseEventArgs) Handles Grid2.MouseDown
        If e.Button = MouseButtons.Right Then
            ContextMenuStrip2.Show(MousePosition, ToolStripDropDownDirection.Right)
            Dim k = TryCast(sender, DataGridView)
            Grid2Clic = Grid2all.ElementAt(k.CurrentCell.RowIndex)
        End If
    End Sub

    Private Sub УдлаитьРейсИзСпискаToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles УдлаитьРейсИзСпискаToolStripMenuItem.Click
        УдалитьРейсИзтаблицыAsync()
        Grid2all.Remove(Grid2Clic)
        bsgrid2.ResetBindings(False)
    End Sub
    Private Async Sub УдалитьРейсИзтаблицыAsync()
        Await Task.Run(Sub() УдалитьРейсИзтаблицы())
    End Sub
    Private Sub УдалитьРейсИзтаблицы()
        Dim mo As New AllUpd
        Using db As New dbAllDataContext()
            Dim f = db.ЖурналКлиентГруз.Where(Function(x) x.Код = Grid2Clic.КодЖурналГруз).FirstOrDefault()
            If f IsNot Nothing Then
                f.ОтоброжатьВТаблицеЖурнала = "False"
                db.SubmitChanges()
                mo.ЖурналКлиентГрузAll()
            End If
        End Using

    End Sub

    Private Sub ДобавитьСобытиеВКалендарьToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ДобавитьСобытиеВКалендарьToolStripMenuItem.Click
        My.Computer.Clipboard.Clear()
        Dim comb2x As IDNaz = ComboBox2.SelectedItem
        Dim comb5x As Com1ЖурналClass = ComboBox5.SelectedItem
        Dim m As String = Grid1Clic.Перевозчик & vbCrLf & Grid1Clic.Контакт & vbCrLf & " - авто для клиента - " _
& vbCrLf & comb2x.Naz & " - груз - " & vbCrLf & comb5x.Наименование & vbCrLf & "(" & comb5x.Lst(0).СтранаПогрузки & ")" & vbCrLf & " - тип авто - " & vbCrLf & comb5x.ТипАвто

        My.Computer.Clipboard.SetText(m)
        MessageBox.Show("Данные в буфере")
    End Sub





    'Private Sub ListBox1_DrawItem(sender As Object, e As DrawItemEventArgs) Handles ListBox1.DrawItem
    '    'e.DrawBackground()
    '    'Dim g As Graphics = e.Graphics
    '    'g.FillRectangle(New SolidBrush(Color.Gray), e.Bounds)
    '    'Dim lb As ListBox = DirectCast(sender, ListBox)
    '    'Dim clr As Color = Color.White
    '    'If e.Index Mod 2 = 1 Then clr = Color.Navy
    '    'g.DrawString(lb.Items(e.Index).ToString(), e.Font, New SolidBrush(clr), New PointF(e.Bounds.X, e.Bounds.Y))
    '    'If e.State And DrawItemState.Selected Then
    '    '    e.Graphics.FillRectangle(SystemBrushes.HotTrack, e.Bounds)
    '    '    g.DrawString(lb.Items(e.Index).ToString(), e.Font, New SolidBrush(clr), New PointF(e.Bounds.X, e.Bounds.Y))
    '    'End If

    '    'e.DrawFocusRectangle()
    'End Sub
End Class