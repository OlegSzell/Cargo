Imports System.ComponentModel

Public Class ПоискПолный
    Private Grid1All As BindingList(Of Grid1Class)
    Private bsGrid1All As BindingSource
    Private list1 As BindingList(Of IDNaz)
    Private bslist1 As BindingSource
    Private Пути As List(Of Путь)
    Private Flag As String = Nothing
    Private VremGrid1All As List(Of Grid1Class)
    Private Sub ПоискПолный_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.MdiParent = MDIParent1
        ПредзагрузкаAsync()
        Grid1All = New BindingList(Of Grid1Class)
        bsGrid1All = New BindingSource
        bsGrid1All.DataSource = Grid1All
        Grid1.DataSource = bsGrid1All
        GridView(Grid1)
        Grid1.Columns(0).Width = 50
        Grid1.Columns(1).MinimumWidth = 150
        Grid1.Columns(2).MinimumWidth = 150
        Grid1.Columns(4).MinimumWidth = 150
        Grid1.Columns(9).Visible = False
        Grid1.Columns(10).Visible = False
        Grid1.Columns(11).Visible = False

        list1 = New BindingList(Of IDNaz)
        bslist1 = New BindingSource With {
            .DataSource = list1
        }
        ListBox1.DataSource = bslist1
        ListBox1.DisplayMember = "Naz"
        Пути = New List(Of Путь)
        ПутьМетодAsync()

        VremGrid1All = New List(Of Grid1Class)
    End Sub
    Private Async Sub ПутьМетодAsync()
        Await Task.Run(Sub() ПутьМетод())
    End Sub
    Private Sub ПутьМетод()
        Dim f = IO.Directory.GetFiles("Z:\RICKMANS\", "*.xls*", IO.SearchOption.AllDirectories).ToList()
        If f IsNot Nothing Then
            For Each b In f
                Dim m As New Путь With {.ПутьПолный = b, .НазваниеРейса = IO.Path.GetFileName(b)}
                Пути.Add(m)
            Next
        End If
    End Sub
    Public Class Путь
        Public Property ПутьПолный As String
        Public Property НазваниеРейса As String

    End Class
    Private Async Sub ПредзагрузкаAsync()
        Await Task.Run(Sub() Предзагрузка())
    End Sub
    Private Sub Предзагрузка()

        Dim mo As New AllUpd
        Do While AllClass.ПеревозчикиБаза Is Nothing
            mo.ПеревозчикиБазаAll()
        Loop
        Do While AllClass.Перевозчики Is Nothing
            mo.ПеревозчикиAll()
        Loop

        Do While AllClass.РейсыПеревозчика Is Nothing
            mo.РейсыПеревозчикаAll()
        Loop

        Do While AllClass.РейсыКлиента Is Nothing
            mo.РейсыКлиентаAll()
        Loop

        Do While AllClass.Клиент Is Nothing
            mo.КлиентAll()
        Loop
        Do While AllClass.ФайлыExcelВсе Is Nothing
            mo.ФайлыExcelВсеAll()
        Loop

    End Sub

    Private Sub Поиск()
        'Dim f6 As New List(Of Grid3ЖурналClass)
        'f6.AddRange(Grid3all.Where(Function(x) x.КонтДанные.ToUpper.Contains(kl.ToUpper)).Select(Function(x) x).ToList())
        'f6.AddRange(Grid3all.Where(Function(x) x.Перевозчик.ToUpper.Contains(kl.ToUpper)).Select(Function(x) x).ToList())
        'f6.AddRange(Grid3all.Where(Function(x) x.Примечание.ToUpper.Contains(kl.ToUpper)).Select(Function(x) x).ToList())
        'f6.AddRange(Grid3all.Where(Function(x) x.Регионы.ToUpper.Contains(kl.ToUpper)).Select(Function(x) x).ToList())
        'f6.AddRange(Grid3all.Where(Function(x) x.Страны.ToUpper.Contains(kl.ToUpper)).Select(Function(x) x).ToList())
        'f6.AddRange(Grid3all.Where(Function(x) x.Транспорт.ToUpper.Contains(kl.ToUpper)).Select(Function(x) x).ToList())

        'Dim f7 = f6.OrderBy(Function(x) x.Перевозчик).Select(Function(x) x.КодПеревозчика).Distinct().ToList()
        'If f7 IsNot Nothing Then
        '    Dim f9 As New List(Of Grid3ЖурналClass)
        '    For Each b In f7
        '        Dim f8 = Grid3all.Where(Function(x) x.КодПеревозчика = b).Select(Function(x) x).FirstOrDefault()
        '        f9.Add(f8)
        '    Next
        '    If f9 IsNot Nothing Then
        '        Grid4all.Clear()
        '        Dim i As Integer = 1
        '        For Each b In f9
        '            b.Номер = i
        '            Grid4all.Add(b)
        '            i += 1
        '        Next
        '    End If
        'End If

    End Sub


    Public Class Grid1Class
        Inherits DopInfo
        Public Property ID As Double
        Public Property Клиент As String
        Public Property Перевозчик As String
        Public Property Груз As String
        Public Property Маршрут As String
        Public Property Контакт As String
        Public Property Загрузка As String
        Public Property Выгрузка As String
        Public Property Примечание As String



    End Class
    Public Class DopInfo
        Public Property FileOpen As String
        Public Property Year As String
        Public Property OpenAll As String
    End Class


    Private Sub Metr3()
        If TextBox3.Text.Length = 0 Then
            Return
        End If




        Dim mo As New AllUpd
        Do While AllClass.ФайлыExcelВсе Is Nothing
            mo.ФайлыExcelВсеAll()
        Loop

        Do While AllClass.РейсыКлиента Is Nothing
            mo.РейсыКлиентаAll()
        Loop
        Do While AllClass.РейсыПеревозчика Is Nothing
            mo.РейсыПеревозчикаAll()
        Loop

        Do While AllClass.Клиент Is Nothing
            mo.КлиентAll()
        Loop
        'IIf(y.Клиент Is Nothing, String.Empty, y.Клиент.Remove(50)

        Dim f = (From y In (From x In AllClass.ФайлыExcelВсе
                            Where x.Груз IsNot Nothing
                            Select x).ToList()
                 Where y.Груз?.ToUpper.Contains(TextBox3.Text.ToUpper)
                 Select New Grid1Class With {.Груз = y.Груз, .Загрузка = y.АдресЗагрузки, .Выгрузка = y.АдресВыгрузки,
                        .Клиент = y.Клиент, .Маршрут = y.Маршрут & vbCrLf & "_____" & vbCrLf & y.Рейс,
                     .FileOpen = Пути.Where(Function(x) x?.ПутьПолный.ToUpper.Contains(Strings.Left(y.Рейс.ToUpper, y.Рейс.Length - 1))).Select(Function(x) x.ПутьПолный).FirstOrDefault(),
                     .Примечание = y?.ДопУсловия})?.ToList()

        For Each b9 In f  'убираем длинное примечание
            If b9.Примечание IsNot Nothing Then
                If b9.Примечание.Length > 50 Then
                    b9.Примечание = b9.Примечание.Substring(0, 50) & " ..."
                End If
            End If
        Next

        Dim f21 = (From x In AllClass.РейсыКлиента
                   Where x.НаименованиеГруза IsNot Nothing
                   Select x).ToList()



        Dim f1 = (From y In f21
                  Where y?.НаименованиеГруза?.ToUpper.Contains(TextBox3.Text.ToUpper)
                  Select New Grid1Class With {.Груз = y.НаименованиеГруза, .Клиент = y.НазвОрганизации, .Контакт = (From z In AllClass.Клиент
                                                                                                                    Where z.НазваниеОрганизации = y?.НазвОрганизации?.ToString
                                                                                                                    Select z.Контактное_лицо & " " & z.Телефон)?.FirstOrDefault,
                      .Выгрузка = y.ТочнАдресРазгр,
                      .Загрузка = y.ТочныйАдресЗагрузки,
                      .Маршрут = y.Маршрут & vbCrLf & "_____" & vbCrLf & (From y1 In (From x In Пути 'выбираем рейс
                                                                                      Where x?.НазваниеРейса?.ToUpper.Contains(y?.НомерРейса?.ToString.ToUpper)
                                                                                      Select x)?.ToList()
                                                                          Where y1?.НазваниеРейса?.ToUpper.Contains(y?.НазвОрганизации?.ToUpper)
                                                                          Select y1.НазваниеРейса)?.FirstOrDefault(),
                      .FileOpen = (From x In Пути
                                   Where x?.ПутьПолный?.ToUpper.Contains(Trim(Replace(.Маршрут, y.Маршрут & vbCrLf & "_____" & vbCrLf, ""))?.ToUpper)
                                   Select x.ПутьПолный)?.FirstOrDefault(),
                      .Перевозчик = (From c In AllClass.РейсыПеревозчика
                                     Where c.НомерРейса = y.НомерРейса
                                     Select c.НазвОрганизации).FirstOrDefault(),
                       .Примечание = y.ДопУсловия})?.ToList()


        For Each b10 In f1  'убираем длинное примечание
            If b10.Примечание IsNot Nothing Then
                If b10.Примечание.Length > 50 Then
                    b10.Примечание = b10.Примечание.Substring(0, 50) & " ..."
                End If
            End If
        Next



        For Each b In AllClass.Клиент    'перебираем и ищем название организации что бы вставитиь поменьше название
            Dim f4 = (From x In f
                      Where x.Клиент?.ToUpper.Contains(b.НазваниеОрганизации.ToUpper)
                      Select x).ToList()
            If f4 IsNot Nothing Then
                If f4.Count > 0 Then
                    For Each b1 In f4
                        b1.Клиент = b.НазваниеОрганизации
                    Next
                End If
            End If
        Next

        If Grid1All IsNot Nothing Then
            Grid1All.Clear()

        End If

        If list1 IsNot Nothing Then
            list1.Clear()

        End If

        Dim G1 As New List(Of Grid1Class)
        Dim i As Integer = 1
        If f IsNot Nothing Then
            For Each b In f
                b.ID = i
                G1.Add(b)
                i += 1
            Next
        End If

        If f1 IsNot Nothing Then
            For Each b In f1
                b.ID = i
                G1.Add(b)
                i += 1
            Next
        End If

        Dim f5 = (From x In (From y In G1
                             Order By y.Клиент
                             Select y).ToList
                  Group x By Keys = New With {Key x.Клиент}
                                 Into Group
                  Select New With {.Клиент = Keys.Клиент, .Coll = Group}).ToList()
        Dim i1 As Integer = 1
        For Each b In f5

            Dim contct = (From v In AllClass.Клиент
                          Where v.НазваниеОрганизации = b.Клиент
                          Select v.Контактное_лицо & vbCrLf & v.Телефон)?.FirstOrDefault

            If contct Is Nothing Then
                contct = "AU"
            End If
            Dim m As New Grid1Class With {.Клиент = b.Клиент, .Контакт = IIf(contct.Length = 2, (From n In AllClass.ФайлыExcelВсе
                                                                                                 Where n.Клиент?.ToUpper.Contains(b.Клиент.ToUpper)
                                                                                                 Select n.Телефон)?.FirstOrDefault(), contct), .ID = i1} 'вставляем данные в главные названия
            list1.Add(New IDNaz With {.ID = i1, .Naz = b.Клиент})
            Grid1All.Add(m)

            Dim il As Integer = 1
            For Each b1 In b.Coll 'вставляем данные в остальные
                Dim m1 As New Grid1Class With {.Выгрузка = b1.Выгрузка,
                    .Груз = b1.Груз,
                    .Загрузка = b1.Загрузка,
                    .Маршрут = b1.Маршрут,
                    .FileOpen = b1.FileOpen,
                    .Перевозчик = b1.Перевозчик,
                    .Примечание = b1.Примечание,
                    .ID = CDbl(CType(i1, String) & "," & CType(il, String))}
                Grid1All.Add(m1)
                il += 1
            Next





            i1 += 1
        Next

    End Sub

    Private Sub ListBox1_Click(sender As Object, e As EventArgs) Handles ListBox1.Click
        If ListBox1.SelectedIndex = -1 Then Return
        Dim f As IDNaz = list1.ElementAt(ListBox1.SelectedIndex)

        Dim f1 As New Grid1Class
        Dim f2 As Integer

        Select Case Flag
            Case "Перевозчик"
                f1 = Grid1All.Where(Function(x) x.Перевозчик = f.Naz).Select(Function(x) x).FirstOrDefault()
            Case "Клиент"
                f1 = Grid1All.Where(Function(x) x.Клиент = f.Naz).Select(Function(x) x).FirstOrDefault()
            Case "Груз"
                f1 = Grid1All.Where(Function(x) x.Клиент = f.Naz).Select(Function(x) x).FirstOrDefault()
        End Select



        f2 = Grid1All.IndexOf(f1)
        If f2 = -1 Then Return

        Grid1.ClearSelection()
        Grid1.FirstDisplayedScrollingRowIndex = f2
        Grid1.Rows(f2).Selected = True
    End Sub
    Private Sub ClearList1()
        If list1 IsNot Nothing Then
            list1.Clear()
        End If
    End Sub
    Private Sub Btn1()

        If TextBox1.Text.Length = 0 Then
            Return
        End If

        ClearList1()

        Flag = "Перевозчик"


        Dim mo As New AllUpd
        Do While AllClass.Клиент Is Nothing
            mo.КлиентAll()
        Loop

        Do While AllClass.Перевозчики Is Nothing
            mo.ПеревозчикиAll()
        Loop
        Do While AllClass.РейсыКлиента Is Nothing
            mo.РейсыКлиентаAll()
        Loop
        Do While AllClass.РейсыПеревозчика Is Nothing
            mo.РейсыПеревозчикаAll()
        Loop
        Do While AllClass.ФайлыExcelВсе Is Nothing
            mo.ФайлыExcelВсеAll()
        Loop
        Do While AllClass.ПеревозчикиБаза Is Nothing
            mo.ПеревозчикиБазаAll()
        Loop
        Dim fa = (From x In AllClass.ПеревозчикиБаза
                  Order By x.Наименование_фирмы
                  Where x?.Наименование_фирмы.ToUpper.Contains(TextBox1.Text.ToUpper)
                  Select x.Наименование_фирмы.ToUpper).ToList()

        Dim fb = (From x In AllClass.Перевозчики
                  Order By x.Названиеорганизации
                  Where x?.Названиеорганизации.ToUpper.Contains(TextBox1.Text.ToUpper)
                  Select x.Названиеорганизации.ToUpper).ToList()

        Dim f = fa.Union(fb).OrderBy(Function(x) x)

        If f IsNot Nothing Then
            If Grid1All IsNot Nothing Then
                Grid1All.Clear()
            End If
            Dim nm As String = Nothing
            Dim i As Integer = 1
            For Each b In f
                Dim k = AllClass.ПеревозчикиБаза.Where(Function(x) x.Наименование_фирмы.ToUpper = b).Select(Function(x) x).FirstOrDefault()
                If k IsNot Nothing Then
                    Dim contDan As String = IIf(k.Контактное_лицо Is Nothing, String.Empty, k.Контактное_лицо) & vbCrLf & IIf(k.Телефоны Is Nothing, String.Empty, k.Телефоны)
                    'создаем заглавную строку
                    Dim m As New Grid1Class With {.ID = i, .Перевозчик = k.Наименование_фирмы, .Контакт = contDan, .Загрузка = k.Страны_перевозок, .Примечание = k.Примечание}
                    nm = k.Наименование_фирмы
                    Grid1All.Add(m)
                Else
                    Dim k1 = AllClass.Перевозчики.Where(Function(x) x.Названиеорганизации.ToUpper = b).Select(Function(x) x).FirstOrDefault()
                    If k1 IsNot Nothing Then
                        Dim contDan1 As String = IIf(k1.Контактное_лицо Is Nothing, String.Empty, k1?.Контактное_лицо) & vbCrLf & IIf(k1.Телефон Is Nothing, String.Empty, k1?.Телефон)
                        Dim m4 As New Grid1Class With {.ID = i, .Перевозчик = k1.Названиеорганизации, .Контакт = contDan1}
                        nm = k1.Названиеорганизации
                        Grid1All.Add(m4)
                    End If
                End If
                list1.Add(New IDNaz With {.Naz = nm})

                'Первый этап выборка в РейсыПеревозчика
                Dim рейсыПеревозчикаВыборка = (From x In AllClass.РейсыПеревозчика
                                               Where x?.НазвОрганизации.ToUpper.Contains(b)
                                               Select x)?.ToList()
                Dim ik As Integer = 1
                'заполныем по каждому перевозчику данные по рейсам из РейсыПеревозчика
                If рейсыПеревозчикаВыборка IsNot Nothing Then

                    For Each b1 In рейсыПеревозчикаВыборка
                        Dim ClientName = (From v In AllClass.РейсыКлиента
                                          Where v.НомерРейса = b1?.НомерРейса
                                          Select v).FirstOrDefault()

                        Dim datYear As String = Nothing
                        If ClientName.ДатаПоручения Is Nothing Then
                            datYear = String.Empty
                        Else
                            datYear = CDate(ClientName.ДатаПоручения).Year.ToString
                        End If

                        Dim Primech As String = Nothing
                        If b1.ДопУсловия Is Nothing Then
                            Primech = String.Empty
                        Else
                            If b1?.ДопУсловия.Length > 0 Then
                                If b1?.ДопУсловия.Length > 50 Then
                                    Primech = b1?.ДопУсловия.Substring(0, 50)
                                Else
                                    Primech = b1?.ДопУсловия
                                End If

                            End If

                        End If


                        Dim flop = (From z In (From y In (From x In Пути  'выбираем номер рейса 
                                                          Where x?.НазваниеРейса.ToUpper.Contains(b1?.НомерРейса.ToString.ToUpper)
                                                          Select x).ToList()
                                               Where y?.НазваниеРейса.ToUpper.Contains(b1?.НазвОрганизации.ToUpper)
                                               Select y).ToList()
                                    Where z?.НазваниеРейса.ToUpper.Contains(ClientName.НазвОрганизации.ToUpper)
                                    Select z).FirstOrDefault

                        Dim f3 As New Grid1Class With {.ID = CDbl(CType(i, String) & "," & CType(ik, String)), .FileOpen = flop?.ПутьПолный, .Контакт = b1?.НаименованиеГруза, .Загрузка = b1?.ТочныйАдресЗагрузки, .Примечание = Primech, .Клиент = ClientName?.НазвОрганизации,
                            .Выгрузка = b1?.ТочнАдресРазгр, .Маршрут = b1?.Маршрут & vbCrLf & "____" & vbCrLf & flop?.НазваниеРейса & " - " & datYear}

                        Grid1All.Add(f3)
                        ik += 1
                    Next
                End If


                'Второй этап выборка в ФайлыexcelВсе
                Dim msd = (From x In AllClass.ФайлыExcelВсе
                           Where x.Рейс.ToUpper.Contains(b)
                           Select x).ToList()

                If msd IsNot Nothing Then
                    For Each b4 In msd
                        'Клиент короткое название
                        Dim corNameClient As String = b4?.Клиент
                        For Each b7 In AllClass.Клиент
                            If b4?.Клиент.ToUpper.Contains(b7?.НазваниеОрганизации.ToUpper) = True Then
                                corNameClient = b7?.НазваниеОрганизации
                                Exit For
                            End If
                        Next

                        'если перевозчик так же был как заказчик(виталюр)!
                        If b = corNameClient.ToUpper Then
                            Continue For
                        End If

                        Dim flp = (From z In Пути
                                   Where z?.НазваниеРейса.ToUpper.Contains(Strings.Left(b4?.Рейс.ToUpper, b4?.Рейс.Length - 1))
                                   Select z).FirstOrDefault()


                        Dim prm As String = Nothing
                        If b4?.ДопУсловия IsNot Nothing Then
                            If b4?.ДопУсловия.Length > 50 Then
                                prm = b4?.ДопУсловия.Substring(0, 50) & "..."
                            Else
                                prm = b4?.ДопУсловия
                            End If
                        End If


                        Dim dtYear As String = Nothing
                        If b4?.ДатаПоручения IsNot Nothing Then
                            dtYear = CDate(b4?.ДатаПоручения).Year.ToString
                        Else
                            dtYear = String.Empty

                        End If

                        Dim f5 As New Grid1Class With {.ID = CDbl(CType(i, String) & "," & CType(ik, String)), .FileOpen = flp?.ПутьПолный, .Контакт = b4?.Телефон, .Загрузка = b4?.АдресЗагрузки, .Примечание = prm, .Клиент = corNameClient,
                                      .Выгрузка = b4?.АдресВыгрузки, .Маршрут = b4?.Маршрут & vbCrLf & "____" & vbCrLf & flp?.НазваниеРейса & " - " & dtYear}

                        If Grid1All IsNot Nothing Then
                            Dim mkl = Grid1All.Where(Function(x) x?.Маршрут = f5.Маршрут).Select(Function(x) x).FirstOrDefault()
                            If mkl Is Nothing Then
                                Grid1All.Add(f5)
                            End If
                        Else
                            Grid1All.Add(f5)
                        End If



                        ik += 1

                    Next
                End If

                i += 1

            Next
        End If

        If VremGrid1All IsNot Nothing Then
            VremGrid1All.Clear()
            VremGrid1All.AddRange(Grid1All)
        End If
    End Sub

    Private Sub Button1_Click_1(sender As Object, e As EventArgs) Handles Button1.Click


        Btn1()

    End Sub
    Private Sub btn2()
        If TextBox2.Text.Length = 0 Then
            Return
        End If

        ClearList1()

        Flag = "Клиент"

        Dim mo As New AllUpd
        Do While AllClass.Клиент Is Nothing
            mo.КлиентAll()
        Loop
        Do While AllClass.РейсыКлиента Is Nothing
            mo.РейсыКлиентаAll()
        Loop
        Do While AllClass.РейсыПеревозчика Is Nothing
            mo.РейсыПеревозчикаAll()
        Loop
        Do While AllClass.ФайлыExcelВсе Is Nothing
            mo.ФайлыExcelВсеAll()
        Loop


        ' Dim f3 = AllClass.ФайлыExcelВсе.Where(Function(x) x.Клиент IsNot Nothing).Select(Function(x) x).ToList()

        Dim f = (From x In AllClass.Клиент
                 Order By x.НазваниеОрганизации
                 Where x.НазваниеОрганизации?.ToUpper.Contains(TextBox2.Text.ToUpper)
                 Select x).ToList()

        If f IsNot Nothing Then
            If Grid1All IsNot Nothing Then
                Grid1All.Clear()
            End If
            Dim i As Integer = 1



            For Each b In f
                Dim Vrem As New List(Of Grid1Class)
                Dim contDan As String = IIf(b.Контактное_лицо Is Nothing, String.Empty, b.Контактное_лицо) & vbCrLf & IIf(b.Телефон Is Nothing, String.Empty, b.Телефон)

                If contDan Is Nothing Then
                    contDan = "AU"
                End If

                If contDan.Length = 2 Then

                    Dim f2 = (From x In AllClass.ФайлыExcelВсе
                              Order By x.ДатаПоручения Descending
                              Where x.Клиент?.ToUpper.Contains(b.НазваниеОрганизации.ToUpper)
                              Select x).FirstOrDefault()
                    If f2 IsNot Nothing Then

                        contDan = IIf(f2.Телефон Is Nothing, String.Empty, f2.Телефон)

                    End If
                End If
                Dim m As New Grid1Class With {.ID = i, .Клиент = b.НазваниеОрганизации, .Контакт = contDan}
                Vrem.Add(m)
                list1.Add(New IDNaz With {.ID = i, .Naz = b.НазваниеОрганизации})

                'ищем рейсы по каждому клиенту и добавляем в группу

                Dim f3 = (From x In AllClass.РейсыКлиента  'текщая база
                          Where x.НазвОрганизации = b.НазваниеОрганизации
                          Select x).ToList()

                Dim f4 = (From x In AllClass.ФайлыExcelВсе  'старые рейсы
                          Where x.Клиент?.ToUpper.Contains(b.НазваниеОрганизации.ToUpper)
                          Select x).ToList()
                Dim ids As Integer = 1
                Dim im2 As Integer = 0


                If f3 IsNot Nothing Then
                    Dim f5 = (From x In f3 'перебираем и вставляем во временный лист
                              Group x By Keys = New With {Key x.НазвОрганизации}
                                 Into Group
                              Select New With {Keys.НазвОрганизации, .Coll = Group.ToList()}).ToList()

                    For Each b4 In f5   'поиск в рейсах клиента
                        For Each bi2 In b4.Coll
                            Dim PolnPut As String = Nothing 'для  FileOpen
                            Dim marsh As String = Nothing

                            Dim nazOrg = b4.НазвОрганизации 'выбиарем организацию
                            Dim nazPer = (From y In AllClass.РейсыПеревозчика 'выбиарем перевозчика
                                          Where y.НомерРейса = bi2.НомерРейса
                                          Select y.НазвОрганизации).FirstOrDefault()

                            If marsh Is Nothing Then
                                Dim s = (From a In (From c In Пути
                                                    Where Strings.Left(c?.НазваниеРейса, 3).ToUpper.Contains(CType(bi2.НомерРейса, String)?.ToUpper)
                                                    Select c).ToList()
                                         Where a.НазваниеРейса?.ToUpper.Contains(nazOrg.ToString.ToUpper)
                                         Select a).ToList() 'выбираем пути которые имеют номер рейса 



                                Dim s1 As String = (From x In s
                                                    Where x.НазваниеРейса?.ToUpper.Contains(nazPer?.ToUpper)
                                                    Select x.НазваниеРейса & " - " & x.ПутьПолный.Substring(12, 4) & "г.").FirstOrDefault()
                                If s1 IsNot Nothing Then
                                    marsh = bi2.Маршрут & vbCrLf & " ____" & vbCrLf & s1
                                End If
                                PolnPut = (From x In s
                                           Where x.НазваниеРейса?.ToUpper.Contains(nazPer?.ToUpper)
                                           Select x.ПутьПолный).FirstOrDefault()


                                'Else

                                '    Dim s = (From a In (From c In Пути
                                '                        Where Strings.Left(c?.НазваниеРейса, 3).ToUpper.Contains(CType(bi2.НомерРейса, String)?.ToUpper)
                                '                        Select c).ToList()
                                '             Where a.НазваниеРейса?.ToUpper.Contains(nazOrg.ToString.ToUpper)
                                '             Select a).ToList()
                                '    Dim s1 As String = (From x In s
                                '                        Where x.НазваниеРейса?.ToUpper.Contains(nazPer?.ToUpper)
                                '                        Select x.НазваниеРейса & " - " & x.ПутьПолный.Substring(12, 4) & "г.").FirstOrDefault()
                                '    If s1 IsNot Nothing Then
                                '        marsh = marsh & vbCrLf & bi2.Маршрут & vbCrLf & s1
                                '    End If


                            End If

                            Dim m2 As New Grid1Class With {.ID = CDbl(CType(i, String) & "," & CType(ids, String)), .Выгрузка = bi2.ТочнАдресРазгр,
                                .Груз = bi2.НаименованиеГруза, .Загрузка = bi2.ТочныйАдресЗагрузки, .Маршрут = marsh, .Примечание = bi2.ДопУсловия.Substring(0, 50),
                                .Перевозчик = nazPer, .FileOpen = PolnPut}
                            Vrem.Add(m2)
                            im2 += 1
                            ids += 1
                        Next


                    Next
                End If


                If f4 IsNot Nothing Then 'это из Excel
                    Dim f7 = (From x In f4
                              Group x By Keys = New With {Key x.Клиент}
                                 Into Group
                              Select New With {.Клиент2 = Keys.Клиент, .Coll = Group}).ToList()

                    Dim im3 As Integer = 0
                    For Each b5 In f7
                        For Each b1 In b5.Coll
                            Dim per As String
                            If b1.Перевозчик Is Nothing Then
                                per = String.Empty
                            Else
                                per = b1.Перевозчик
                            End If
                            Dim dt As String
                            If b1.ДатаПоручения Is Nothing Then
                                dt = String.Empty
                            ElseIf b1.ДатаПоручения.Length = 0 Then
                                dt = String.Empty
                            Else
                                dt = CDate(b1?.ДатаПоручения).Year.ToString
                            End If
                            Dim FlOp = Пути.Where(Function(c) c.НазваниеРейса.ToUpper.Contains(Strings.Left(b1?.Рейс.ToUpper, b1?.Рейс.Length - 1))).Select(Function(c) c.ПутьПолный).FirstOrDefault()

                            Dim prim As String
                            If b1?.ДопУсловия Is Nothing Then
                                prim = String.Empty
                            Else
                                If b1?.ДопУсловия.Length > 50 Then
                                    prim = b1?.ДопУсловия.Substring(0, 50)
                                Else
                                    prim = b1?.ДопУсловия
                                End If
                            End If

                            Dim m3 As New Grid1Class With {.ID = CDbl(CType(i, String) & "," & CType(ids, String)), .Выгрузка = b1?.АдресВыгрузки,
                                 .Груз = b1?.Груз, .Загрузка = b1?.АдресЗагрузки, .Маршрут = b1?.Маршрут & vbCrLf & " ____" & vbCrLf & b1.Рейс & " - год " & dt, .Примечание = prim,
                                 .Перевозчик = Trim(per), .FileOpen = FlOp}
                            Vrem.Add(m3)
                            im3 += 1
                            ids += 1
                        Next

                    Next
                End If


                For Each b3 In Vrem
                    Grid1All.Add(b3)
                Next

                i += 1
            Next



        End If

        If VremGrid1All IsNot Nothing Then
            VremGrid1All.Clear()
            VremGrid1All.AddRange(Grid1All)
        End If

    End Sub

    Private Sub Button2_Click_1(sender As Object, e As EventArgs) Handles Button2.Click
        btn2()

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        ClearList1()
        Flag = "Груз"
        Metr3()
        If VremGrid1All IsNot Nothing Then
            VremGrid1All.Clear()
            VremGrid1All.AddRange(Grid1All)
        End If
    End Sub

    Private Sub TextBox3_KeyDown_1(sender As Object, e As KeyEventArgs) Handles TextBox3.KeyDown
        If e.KeyCode = Keys.Enter Then
            If Button3.Text.Length > 0 Then
                Metr3()
            End If
        End If
    End Sub
    Private Async Sub Btn4()
        Flag = "Маршрут"

        If TextBox4.Text.Length = 0 Then
            Return
        End If
        Cursor = Cursors.WaitCursor
        If Grid1All IsNot Nothing Then
            Grid1All.Clear()
        End If


        Dim kl = Await MarshAsync(TextBox4.Text)

        If kl IsNot Nothing Then
            For Each b In kl
                Grid1All.Add(b)
            Next


            If VremGrid1All IsNot Nothing Then
                VremGrid1All.Clear()
                VremGrid1All.AddRange(Grid1All)
            End If
        End If
        Cursor = Cursors.Default
    End Sub


    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Btn4()
    End Sub
    Private Async Function MarshAsync(ByVal txt4 As String) As Task(Of List(Of Grid1Class))
        Return Await Task.Run(Function() Marsh(txt4))
    End Function
    Private Function Marsh(ByVal txt4 As String) As List(Of Grid1Class)


        ClearList1()
        Dim grd As New List(Of Grid1Class)

        Dim mo As New AllUpd
        Do While AllClass.Клиент Is Nothing
            mo.КлиентAll()
        Loop
        Do While AllClass.РейсыКлиента Is Nothing
            mo.РейсыКлиентаAll()
        Loop
        Do While AllClass.РейсыПеревозчика Is Nothing
            mo.РейсыПеревозчикаAll()
        Loop
        Do While AllClass.ФайлыExcelВсе Is Nothing
            mo.ФайлыExcelВсеAll()
        Loop

        Dim f = (From x In AllClass.РейсыКлиента
                 Order By x.НазвОрганизации
                 Where x.Маршрут?.ToUpper.Contains(txt4.ToUpper)
                 Select x).ToList()

        Dim f1 = (From x In AllClass.ФайлыExcelВсе
                  Where x.Маршрут?.ToUpper.Contains(txt4.ToUpper)
                  Select x).ToList()


        For Each b In AllClass.Клиент
            If f1 IsNot Nothing Then
                'короткое название организации
                Dim f3 = (From x In f1
                          Where x?.Клиент?.ToString.ToUpper.Contains(b?.НазваниеОрганизации?.ToUpper)
                          Select x).ToList()

                If f3 IsNot Nothing Then
                    For Each b1 In f3
                        f1(f1.IndexOf(b1)).Клиент = b.НазваниеОрганизации
                    Next

                End If
            End If
        Next



        'первый этап РейсыКлиента
        Dim i As Integer = 1
        For Each b1 In f

            Dim put As New Путь

            put = (From x In Пути
                   Where Strings.Left(x.НазваниеРейса, 3)?.ToUpper.Contains(b1?.НомерРейса.ToString.ToUpper) _
                       And Strings.Right(x.НазваниеРейса, x.НазваниеРейса.Length - 3)?.ToUpper.Contains(b1?.КоличРейсов.ToString.ToUpper) _
                       And x.НазваниеРейса?.ToUpper.Contains(b1?.НазвОрганизации.ToUpper)
                   Select x).FirstOrDefault()

            Dim mars As String = b1.Маршрут?.ToString & vbCrLf & "______" & vbCrLf & put?.НазваниеРейса & " - " & CDate(b1.ДатаПоручения?.ToString).Year.ToString
            'If put Is Nothing Then

            'End If



            Dim m As New Grid1Class With {.FileOpen = put?.ПутьПолный, .Выгрузка = Trim(b1.ТочнАдресРазгр?.ToString), .Груз = Trim(b1.НаименованиеГруза?.ToString),
                .Загрузка = Trim(b1.ТочныйАдресЗагрузки?.ToString), .Клиент = Trim(b1.НазвОрганизации?.ToString), .Контакт = Trim((From x In AllClass.Клиент
                                                                                                                                   Where x.НазваниеОрганизации?.ToString = b1.НазвОрганизации
                                                                                                                                   Select x.Контактное_лицо & vbCrLf & x.Телефон).FirstOrDefault()),
                                                                                                                      .Перевозчик = Trim((From x In AllClass.РейсыПеревозчика
                                                                                                                                          Where x?.НомерРейса = b1?.НомерРейса
                                                                                                                                          Select x.НазвОрганизации).FirstOrDefault()),
                                                                                                                                         .Примечание = Trim(b1.ДопУсловия?.ToString), .Маршрут = mars}

            If m?.Примечание?.Length > 50 Then
                m.Примечание = m.Примечание.Substring(0, 50) & " ..."
            End If
            If m?.Загрузка?.Length > 50 Then
                m.Загрузка = m.Загрузка.Substring(0, 50) & " ..."
            End If
            If m?.Выгрузка?.Length > 50 Then
                m.Выгрузка = m.Выгрузка.Substring(0, 50) & " ..."
            End If

            If m?.Груз?.Length > 150 Then
                m.Груз = m.Груз.Substring(0, 150) & " ..."
            End If


            grd.Add(m)

        Next

        'второй этап ФайлыExcelВсе
        For Each b2 In f1
            Dim put As New Путь

            put = (From x In Пути
                   Where x?.НазваниеРейса?.ToString.ToUpper.Contains(Strings.Left(b2?.Рейс.ToUpper, b2?.Рейс.Length - 1))
                   Select x).FirstOrDefault()




            Dim mars As String = b2.Маршрут?.ToString & vbCrLf & "______" & vbCrLf & put?.НазваниеРейса & " - "
            If b2.ДатаПоручения IsNot Nothing Then
                If b2.ДатаПоручения.Length > 10 Then
                    mars &= CDate(Strings.Left(b2.ДатаПоручения?.ToString, 10)).Year.ToString
                ElseIf b2.ДатаПоручения.Length = 10 Then
                    mars &= CDate(b2.ДатаПоручения?.ToString).Year.ToString
                End If
            End If





            Dim perev As String = b2?.Перевозчик?.ToString
            If perev IsNot Nothing Then
                If perev.Length > 50 Then
                    perev = perev.Substring(0, 50)
                End If
            End If


            Dim m As New Grid1Class With {.FileOpen = put?.ПутьПолный, .Выгрузка = Trim(b2.АдресВыгрузки?.ToString), .Груз = Trim(b2.Груз?.ToString),
                 .Загрузка = b2.АдресЗагрузки?.ToString, .Клиент = Trim(b2.Клиент?.ToString), .Контакт = (From x In AllClass.Клиент
                                                                                                          Where x.НазваниеОрганизации?.ToString = b2.Клиент
                                                                                                          Select x.Контактное_лицо & vbCrLf & x.Телефон).FirstOrDefault(),
                                                                                                                      .Перевозчик = Trim(perev),
                                                                                                                                          .Примечание = Trim(b2.ДопУсловия?.ToString), .Маршрут = mars}



            If m?.Примечание?.Length > 50 Then
                m.Примечание = m.Примечание.Substring(0, 50) & " ..."
            End If
            If m?.Загрузка?.Length > 50 Then
                m.Загрузка = m.Загрузка.Substring(0, 50) & " ..."
            End If
            If m?.Выгрузка?.Length > 50 Then
                m.Выгрузка = m.Выгрузка.Substring(0, 50) & " ..."
            End If
            If m?.Груз?.Length > 150 Then
                m.Груз = m.Груз.Substring(0, 150) & " ..."
            End If

            If grd IsNot Nothing Then
                Dim mgj = grd.Where(Function(x) x?.Маршрут?.ToString = m.Маршрут).Select(Function(x) x).FirstOrDefault()
                If mgj IsNot Nothing Then
                    Continue For
                Else
                    grd.Add(m)
                End If
            Else
                grd.Add(m)
            End If
            'grd.Add(m)

        Next
        Dim mk = grd.OrderBy(Function(x) x.Клиент).Select(Function(x) x).ToList()

        If mk IsNot Nothing Then
            For Each b7 In mk
                b7.ID = i
                b7.Груз = Trim(b7.Груз)
                b7.Загрузка = Trim(b7.Загрузка)
                b7.Выгрузка = Trim(b7.Выгрузка)
                b7.Груз = Trim(b7.Груз)
                i += 1
            Next

            'Dim mkl1 = (From x In mk
            '            Select x.Маршрут).Distinct().ToList()
            'If mkl1 IsNot Nothing Then
            '    Dim lop = mk.Where(Function(x) x.Маршрут).Select(Function(x) x.Маршрут).Intersect(mkl1)

            'End If

            Return mk
        End If


        Return Nothing


    End Function

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Flag = "Телефон"
        If VremGrid1All IsNot Nothing Then
            VremGrid1All.Clear()
            VremGrid1All.AddRange(Grid1All)
        End If
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Flag = "АдресЗагрузки"
        If VremGrid1All IsNot Nothing Then
            VremGrid1All.Clear()
            VremGrid1All.AddRange(Grid1All)
        End If
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Flag = "АдресВыгрузки"
        If VremGrid1All IsNot Nothing Then
            VremGrid1All.Clear()
            VremGrid1All.AddRange(Grid1All)
        End If
    End Sub

    Private Sub ОткрытьРейсToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ОткрытьРейсToolStripMenuItem.Click
        Dim f As Grid1Class = Grid1All.ElementAt(Grid1.CurrentRow.Index)
        If f IsNot Nothing Then
            If f.FileOpen IsNot Nothing Then
                Process.Start(f.FileOpen)
            End If

        End If
    End Sub

    Private Sub ОткрытьПолнуюИнформациюToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ОткрытьПолнуюИнформациюToolStripMenuItem.Click

            End Sub

    Private Sub Grid1_MouseDown(sender As Object, e As MouseEventArgs) Handles Grid1.MouseDown
        If e.Button = MouseButtons.Right Then
            If Grid1All IsNot Nothing Then
                If Grid1All.Count > 0 Then
                    ContextMenuStrip1.Show(MousePosition, ToolStripDropDownDirection.Right)
                End If

            End If

        End If
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        If TextBox8.Text.Length = 0 Then
            Return
        End If

        If VremGrid1All Is Nothing Then
            Return
        End If

        Grid1All.Clear()

        For Each b In VremGrid1All
            Dim lst As String() = {b.Выгрузка, b.Груз, b.Загрузка, b.Клиент, b.Контакт, b.Маршрут, b.Перевозчик, b.Примечание}
            Dim f1 As String = String.Join(" , ", lst)
            If f1?.ToUpper.Contains(TextBox8.Text.ToUpper) = True Then
                Grid1All.Add(b)
            End If
        Next

    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        If VremGrid1All IsNot Nothing Then
            Grid1All.Clear()
            For Each b In VremGrid1All
                Grid1All.Add(b)
            Next
        End If
    End Sub

    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown
        If e.KeyCode = Keys.Enter Then
            Btn1()
        End If
    End Sub

    Private Sub TextBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox2.KeyDown
        If e.KeyCode = Keys.Enter Then
            Btn2()
        End If
    End Sub

    Private Sub TextBox4_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox4.KeyDown
        If e.KeyCode = Keys.Enter Then
            Btn4()
        End If
    End Sub
End Class