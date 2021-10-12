Imports System.Collections.ObjectModel
Imports System.ComponentModel
Imports System.Globalization
Imports System.IO
Imports System.Runtime.CompilerServices
Imports System.Web.ApplicationServices
Imports Microsoft.Office.Interop
Imports Microsoft.VisualBasic.CompilerServices

Public Class ГотовыйОтчет
    Implements INotifyPropertyChanged

    Private _LS1 As ObservableCollection(Of String)
    Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged
    Public Sub OnPropertyChanged(<CallerMemberName> ByVal Optional propertyName As String = Nothing)
        RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(propertyName))
    End Sub
    Public Property LS1 As ObservableCollection(Of String)
        Get
            Return _LS1
        End Get
        Set(value As ObservableCollection(Of String))
            _LS1 = value
            OnPropertyChanged("LS1")
        End Set
    End Property
    Public Class Grid1Class
        Public Property Номер As Integer
        Public Property номер_Рейса As Integer
        Public Property Клиент As String
        Public Property Перевозчик As String
        Public Property маршрут As String
        Public Property дата_Загрузки As String
        Public Property ставка_Клиента As String
        Public Property ставка_Перевозчика As String
        Public Property дельта As String

    End Class

    Private grd1 As New List(Of Grid1Class)
    Private bsgrd1 As New BindingSource

    Dim LS2 As New List(Of String)
    Dim bs As New BindingSource
    Private Sub ГотовыйОтчет_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.MdiParent = MDIParent1

        Dim dat() As String = {Now.Year - 2, Now.Year - 1, Now.Year}
        ComboBox1.Items.Clear()
        ComboBox1.Items.AddRange(dat)
        ComboBox2.Items.Clear()

        For x As Integer = 1 To 12

            ComboBox2.Items.Add(New MonthItem With {.Id = x, .Text = MonthName(x)})
        Next

        ComboBox2.DisplayMember = "Text"
        ComboBox2.ValueMember = "Id"

        'ListBox1.Items.Clear()

        Using db As New dbAllDataContext(_cn3)
            Dim f = (From x In db.ОтчетРаботыСотрудника
                     Order By x.Год
                     Select x).ToList()
            If f.Count > 0 Then
                For Each b In f
                    Dim idname As String
                    If b.ID.ToString.Length = 1 Then
                        idname = "00" & b.ID
                    ElseIf b.ID.ToString.Length = 2 Then
                        idname = "0" & b.ID
                    Else
                        idname = b.ID
                    End If
                    LS2.Add(idname & " - " & b.Год & ", " & b.Месяц)
                Next

            End If
            bs.DataSource = LS2


            ListBox1.DataSource = bs
        End Using

        bsgrd1.DataSource = grd1
        Grid1.DataSource = bsgrd1
        GridView(Grid1)
        Grid1.Columns(0).Width = 80
        Grid1.Columns(1).Width = 90


    End Sub
    Private Sub UpdList1()
        LS2.Clear()

        Using db As New dbAllDataContext(_cn3)
            Dim f = (From x In db.ОтчетРаботыСотрудника
                     Order By x.Год
                     Select x).ToList()
            If f.Count > 0 Then
                For Each b In f
                    Dim idname As String
                    If b.ID.ToString.Length = 1 Then
                        idname = "00" & b.ID
                    ElseIf b.ID.ToString.Length = 2 Then
                        idname = "0" & b.ID
                    Else
                        idname = b.ID
                    End If
                    LS2.Add(idname & " - " & b.Год & ", " & b.Месяц)
                Next
                bs.ResetBindings(True)
            End If
        End Using

    End Sub
    Public Class Отчет
        Inherits ОтчетРаботыСотрудникаСводная
        Public Property ID1 As Integer

        Public Property Год As System.Nullable(Of Integer)

        Public Property Месяц As String

        Public Property Экспедитор1 As String
    End Class


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If ComboBox1.Text.Length = 0 Or ComboBox2.Text.Length = 0 Then
            Return
        End If
        Dim mo As New AllUpd
        Do While AllClass.ОтчетРаботыСотрудника Is Nothing
            mo.ОтчетРаботыСотрудникаAll()
        Loop
        Do While AllClass.ОтчетРаботыСотрудникаСводная Is Nothing
            mo.ОтчетРаботыСотрудникаСводнаяAll()
        Loop

        Do While AllClass.РейсыКлиента Is Nothing
            mo.РейсыКлиентаAll()
        Loop

        Do While AllClass.ОплатыКлиент Is Nothing
            mo.ОплатыКлиентAll()
        Loop
        Do While AllClass.РейсыПеревозчика Is Nothing
            mo.РейсыПеревозчикаAll()
        Loop

        Do While AllClass.ОплатыПер Is Nothing
            mo.ОплатыПерAll()
        Loop



        Me.Cursor = Cursors.WaitCursor

        'проверяем, надо ли пересобирать сборку
        If CheckBox1.Checked = False Then
            'проверяем,есть ли отчет в базе?

            Dim fb As New List(Of Отчет)
            If Экспедитор = "Олег" Then
                fb = (From x In AllClass.ОтчетРаботыСотрудника
                      Join y In AllClass.ОтчетРаботыСотрудникаСводная On x.ID Equals y.IDОтчетРабСотрудн
                      Where x.Год = ComboBox1.Text And x.Месяц = ComboBox2.Text
                      Select New Отчет With {.ID = y.ID, .IDОтчетРабСотрудн = y.IDОтчетРабСотрудн, .ВалютаЗак = y.ВалютаЗак, .ВалютаПер = y.ВалютаПер,
                          .Выгрузка = y.Выгрузка, .ДатаВыгрузки = y?.ДатаВыгрузки, .ДатаЗагрузки = y.ДатаЗагрузки, .ДатаОплатыЗак = y.ДатаОплатыЗак,
                          .ДатаОплатыПер = y.ДатаОплатыПер, .ДатаСоздания = y?.ДатаСоздания, .Дельта = y.Дельта, .Загрузка = y.Загрузка, .заказчик = y.заказчик,
                          .ИтогоЗакБелРуб = y.ИтогоЗакБелРуб, .ИтогоОбщая = y.ИтогоОбщая, .ИтогоПерБелРуб = y.ИтогоПерБелРуб, .ИтогоСотрудник = y.ИтогоСотрудник,
                          .КомиссияЗаПеревод = y.КомиссияЗаПеревод, .КурсЗак = y.КурсЗак, .КурсПер = y.КурсПер, .Перевозчик = y.Перевозчик, .СтрахованиеГруза = y.СтрахованиеГруза,
                          .счет = y.счет, .Экспедитор = y.Экспедитор, .ID1 = x.ID, .Год = x.Год, .Месяц = x.Месяц, .Экспедитор1 = x.Экспедитор}).ToList()
            Else
                fb = (From x In AllClass.ОтчетРаботыСотрудника
                      Join y In AllClass.ОтчетРаботыСотрудникаСводная On x.ID Equals y.IDОтчетРабСотрудн
                      Where x.Год = ComboBox1.Text And x.Месяц = ComboBox2.Text And x.Экспедитор = Экспедитор
                      Select New Отчет With {.ID = y.ID, .IDОтчетРабСотрудн = y.IDОтчетРабСотрудн, .ВалютаЗак = y.ВалютаЗак, .ВалютаПер = y.ВалютаПер,
                          .Выгрузка = y.Выгрузка, .ДатаВыгрузки = y?.ДатаВыгрузки, .ДатаЗагрузки = y.ДатаЗагрузки, .ДатаОплатыЗак = y.ДатаОплатыЗак,
                          .ДатаОплатыПер = y.ДатаОплатыПер, .ДатаСоздания = y?.ДатаСоздания, .Дельта = y.Дельта, .Загрузка = y.Загрузка, .заказчик = y.заказчик,
                          .ИтогоЗакБелРуб = y.ИтогоЗакБелРуб, .ИтогоОбщая = y.ИтогоОбщая, .ИтогоПерБелРуб = y.ИтогоПерБелРуб, .ИтогоСотрудник = y.ИтогоСотрудник,
                          .КомиссияЗаПеревод = y.КомиссияЗаПеревод, .КурсЗак = y.КурсЗак, .КурсПер = y.КурсПер, .Перевозчик = y.Перевозчик, .СтрахованиеГруза = y.СтрахованиеГруза,
                          .счет = y.счет, .Экспедитор = y.Экспедитор, .ID1 = x.ID, .Год = x.Год, .Месяц = x.Месяц, .Экспедитор1 = x.Экспедитор}).ToList()

            End If
            'Dim fb = (From x In AllClass.ОтчетРаботыСотрудника
            '          Join y In AllClass.ОтчетРаботыСотрудникаСводная On x.ID Equals y.IDОтчетРабСотрудн
            '          Where x.Год = ComboBox1.Text And x.Месяц = ComboBox2.Text And x.Экспедитор = Экспедитор
            '          Select New ).ToList()

            Dim collection1 As New List(Of GridReport)
            If fb IsNot Nothing Then
                If fb.Count > 0 Then
                    For Each client In fb
                        Dim g = New GridReport With {.Счет = client.счет, .Заказчик = client.заказчик, .Загрузка = client.Загрузка, .ДатаЗагрузки = client.ДатаЗагрузки,
                        .ДатаВыгрузки = client.ДатаВыгрузки, .BYR = client.ИтогоЗакБелРуб, .Валюта = client.ВалютаЗак, .Курс = client.КурсЗак,
                        .ДатаОплаты = client.ДатаОплатыЗак, .Перевозчик = client.Перевозчик, .BYRПер = client.ИтогоПерБелРуб,
                        .ВалютаПер = client.ВалютаПер, .КурсПер = client.КурсПер, .Дельта = client.Дельта}
                        collection1.Add(g)
                    Next

                    ExcelTableAsync(collection1, ComboBox1.Text & " _ " & ComboBox2.Text)
                    Me.Cursor = Cursors.Default
                    Return
                End If
            End If



        End If






        Dim startDate = New DateTime(ComboBox1.Text, ComboBox2.SelectedIndex + 1, 1)
        Dim enddate = startDate.AddMonths(1).AddDays(-1)










        'выгружаем данные всех рейсов клиентов в коллекцию
        Dim obser As New List(Of РейсыКлиента)
        obser.AddRange(AllClass.РейсыКлиента)

        Dim ObserSelect As New List(Of РейсыКлиента)






        Dim AllPayClient As New List(Of ОплатыКлиент)
        Dim list8 As New List(Of NumberReysSort) 'отобранные фрахты перевозчика, если валютные оплачены полностью то попадают в белках в отчет, иначе в валюте
        Dim list5 As New List(Of NumberReysSort) 'отобранные рейсы которые были полностью оплачены в выбранном месяце и попали в отчет
        Dim list2 As New List(Of NumberReysSort)
        Dim ObserSelectПер1 As New List(Of РейсыПеревозчика)


        'выбираем олпаты от клиентов за выбранный месяц и год


        AllPayClient.AddRange(AllClass.ОплатыКлиент)

        Dim f = (From x In (From y In AllPayClient
                            Where (CType(y.ДатаОплаты, Date)).Year = ComboBox1.Text
                            Select y).ToList()
                 Where MonthName(Month(x.ДатаОплаты)) = ComboBox2.Text
                 Select x).ToList()
        If f.Count > 0 Then
            'выбираем рейсы безповтора
            Dim f1 = f.Select(Function(x) x.Рейс).Distinct().ToList()


            'все оплаты клиентов до выбранной даты
            Dim f2 = AllPayClient.Where(Function(c) c.ДатаОплаты <= enddate).Select(Function(c) c).ToList()

            'группируем и сумируем оплаты с нарастающим итогом до выбранной даты
            Dim f3 = (From x In f2
                      Group x By Keys = New With {Key x.Рейс}
                             Into Group
                      Select New With {.Рейс = Keys.Рейс, .Сумма = Group.Sum(Function(t) t.Сумма), .Дата = Group}).OrderBy(Function(x) x.Рейс).ToList()




            'отбираем только рейсы выбранного месяца в коллекцию
            For Each u In f1
                For Each b In f3
                    If b.Рейс.ToString.Contains(u) Then
                        Dim dataOpaty As String = Nothing
                        If b.Дата IsNot Nothing Then
                            dataOpaty = b.Дата.OrderBy(Function(x) x.ДатаОплаты).Select(Function(x) x.ДатаОплаты).LastOrDefault()
                        End If

                        list2.Add(New NumberReysSort With {.НомерРейса = b.Рейс, .СуммаПоступления = b.Сумма, .ДатаОплаты = dataOpaty})

                    End If
                Next
            Next



            'сравнивает стоимость оплаты с фрахтом по каждому рейсу

            For Each z In list2
                Dim mk = obser.Where(Function(z5) CType(z5?.НомерРейса?.ToString, Integer) = z?.НомерРейса).Select(Function(z5) z5).FirstOrDefault()
                Dim newCl As New List(Of DannZakaz) 'специальный класс что бы создавать коллекции
                If mk IsNot Nothing Then

                    'если оплата в белках
                    If mk.СтоимостьФрахта = z.СуммаПоступления And (mk.Валюта = "Рубль") Then
                        Dim msd = AllPayClient.Where(Function(v) v.Рейс?.ToString = z.НомерРейса).ToList()
                        If msd IsNot Nothing Then
                            For Each b In msd
                                Dim m As New DannZakaz With {.ДатыОплатЗак = b.ДатаОплаты, .КурсВалЗак = "1", .СуммаПоступленияЗак = b.Сумма}
                                newCl.Add(m)
                            Next
                        End If

                    ElseIf Not mk.Валюта = "Рубль" Then 'если была оплата  в белках по курсу

                        'если оплата в белках по курсу
                        If mk.ВалютаПлатежа = "BYN" Then
                            Dim Obsr2 = f2.Where(Function(x) x.Рейс = z.НомерРейса).Select(Function(x) x).ToList()
                            Dim валюта As String = mk.Валюта
                            Dim summa As Double = 0
                            Dim курс As Double

                            Dim lst As New List(Of String)
                            For Each x In Obsr2
                                ' получаем курс валюты
                                Dim m As New NbRbClassNew()
                                курс = Replace(m.Курс(x.ДатаОплаты, валюта), ".", ",")

                                'делим сумму прихода на курс получаем в валюте
                                lst.Add(Math.Round(x.Сумма / курс, 2))
                                summa += x.Сумма / курс

                                Dim m2 As New DannZakaz With {.ДатыОплатЗак = x.ДатаОплаты, .КурсВалЗак = CType(курс, String), .СуммаПоступленияЗак = x.Сумма}
                                newCl.Add(m2)
                            Next
                            'если полученная стоимость равна фрахту в рейсе то попадает в список
                            If Math.Round(summa, 0) = Not CDbl(mk.СтоимостьФрахта) Then
                                newCl = Nothing
                            End If

                        Else
                            'если опалта в валюте
                            If mk.СтоимостьФрахта = z.СуммаПоступления Then
                                Dim msd1 = AllPayClient.Where(Function(v) v.Рейс?.ToString = z.НомерРейса).ToList()
                                If msd1 IsNot Nothing Then
                                    For Each b In msd1
                                        Dim m As New DannZakaz With {.ДатыОплатЗак = b.ДатаОплаты, .КурсВалЗак = "1", .СуммаПоступленияЗак = b.Сумма}
                                        newCl.Add(m)
                                    Next
                                End If
                            End If

                        End If
                    End If


                End If
                If newCl IsNot Nothing Then
                    If newCl.Count > 0 Then
                        list5.Add(New NumberReysSort With {.НомерРейса = z.НомерРейса, .lstZakaz = newCl, .ФрахтЗак = mk.СтоимостьФрахта})
                    End If

                End If

            Next



        End If

        'проверяем оплату перевозчика

        Dim f14 = AllClass.РейсыПеревозчика.Select(Function(x) x).ToList()
        If f IsNot Nothing Then
            ObserSelectПер1.AddRange(f14)
        End If

        Dim pl = AllClass.ОплатыПер.Select(Function(c) c).ToList()

        For Each y1 In list2
            Dim m1 = AllClass.ОплатыПер.Where(Function(x) x.Рейс = y1.НомерРейса).ToList() 'выбираем все оплаты по этому рейсу
            Dim m2 = AllClass.РейсыПеревозчика.Where(Function(x) x.НомерРейса = y1.НомерРейса).FirstOrDefault() 'выбираем валюту платежа по этому рейсу
            Dim newCl1 As New List(Of DannPerev)
            'если оплата поступила

            If m1?.Count > 0 Then 'если оплата была
                If m2.Валюта = "Рубль" Then

                    Dim mv As New DannPerev With {.СуммаПоступленияПер = m2.СтоимостьФрахта, .КурсВалПер = "1"}
                    newCl1.Add(mv)


                    'list8.Add(New NumberReysSort With {.НомерРейса = y1.НомерРейса, .СуммаПоступления = m2.СтоимостьФрахта})
                ElseIf m2.ВалютаПлатежа = "BYN" Then


                    For Each b8 In m1
                        Dim mantht1 As String = b8.ДатаОплаты?.ToShortDateString
                        Dim валюта As String = m2.Валюта
                        Dim summa As Double = 0
                        Dim курс As Double
                        Dim m As New NbRbClassNew()
                        курс = Replace(m.Курс(mantht1, валюта), ".", ",")
                        summa = Math.Round(CDbl(b8.Сумма?.ToString) / курс, 2)

                        Dim msx As New DannPerev With {.СуммаПоступленияПер = b8.Сумма, .ДатыОплатПер = b8.ДатаОплаты, .КурсВалПер = курс}
                        newCl1.Add(msx)

                        If summa = Not m2.СтоимостьФрахта Then 'Если оплатили перевозу не всю сумму
                            Dim delta = m2.СтоимостьФрахта - summa 'разница
                            'узнаем курс на последнее число месяца
                            Dim mantht2 = DateTime.DaysInMonth(CDate(y1.ДатаОплаты).Year, CDate(y1.ДатаОплаты).Month) & "." & CDate(y1.ДатаОплаты).Month & "." & CDate(y1.ДатаОплаты).Year 'вычисляем последнее число месяца
                            If mantht2.Length = 0 Then
                                mantht2 = Now.ToShortDateString
                            End If

                            Dim валюта2 As String = m2.Валюта
                            Dim summa2 As Double = 0
                            Dim курс2 As Double
                            Dim m3 As New NbRbClassNew()
                            курс2 = Replace(m3.Курс(mantht2, валюта2), ".", ",")
                            summa2 = Math.Round(delta * курс2, 2)

                            Dim msx1 As New DannPerev With {.СуммаПоступленияПер = summa2, .ДатыОплатПер = mantht2, .КурсВалПер = курс2}
                            newCl1.Add(msx1)

                        End If

                    Next



                End If

            Else 'если оплата не произведена и валюта плтаежа по курсу то берем курс на последнее число месяца

                If m2 IsNot Nothing Then
                    If m2.ВалютаПлатежа = "BYN" Or m2.ВалютаПлатежа = "€" Then
                        'Dim clnt = AllClass.ОплатыКлиент
                        Dim mantht As String
                        If y1.ДатаОплаты IsNot Nothing Then 'если есть дата последней оплаты клиенту

                            mantht = DateTime.DaysInMonth(CDate(y1.ДатаОплаты).Year, CDate(y1.ДатаОплаты).Month) & "." & CDate(y1.ДатаОплаты).Month & "." & CDate(y1.ДатаОплаты).Year 'вычисляем последнее число месяца

                            If mantht.Length = 0 Then
                                mantht = Now.ToShortDateString
                            End If

                            Dim валюта As String = m2.Валюта
                            Dim summa As Double = 0
                            Dim курс As Double
                            Dim m As New NbRbClassNew()
                            курс = Replace(m.Курс(mantht, валюта), ".", ",")
                            summa = m2.СтоимостьФрахта * курс

                            Dim msx2 As New DannPerev With {.СуммаПоступленияПер = summa, .ДатыОплатПер = mantht, .КурсВалПер = курс}
                            newCl1.Add(msx2)

                        End If
                    Else




                    End If
                ElseIf m2.Валюта = "Рубль" Then
                    Dim msx2 As New DannPerev With {.СуммаПоступленияПер = m2.СтоимостьФрахта, .КурсВалПер = "1"}
                    newCl1.Add(msx2)

                End If
                If m1.Count = 0 And m2.ВалютаПлатежа = "" Then
                    Dim msx2 As New DannPerev With {.СуммаПоступленияПер = m2.СтоимостьФрахта, .КурсВалПер = "1"}
                    newCl1.Add(msx2)
                End If

            End If
            list8.Add(New NumberReysSort With {.НомерРейса = y1.НомерРейса, .ФрахтПер = m2.СтоимостьФрахта, .lstPerev = newCl1})
        Next







        'Отбираем все данные в конечную коллекцию для вставки в эксель
        Dim КоллекцияГотовая As New List(Of GridReport)
        Dim Готовый As GridReport
        For Each x In list5

            Dim Driver = ObserSelectПер1.Where(Function(c) c.НомерРейса?.ToString = x.НомерРейса).Select(Function(c) c).FirstOrDefault()
            Dim client = obser.Where(Function(c) c.НомерРейса?.ToString = x.НомерРейса).Select(Function(c) c).FirstOrDefault()
            Dim dltZak As Double = 0
            Dim dltPer As Double = 0

            Dim фрахт As String = Nothing
            Dim курс As String = Nothing
            Dim белки As String = Nothing

            If client Is Nothing Then Continue For

            If client.Валюта = "Рубль" Then 'оплата в белках

                For Each b2 In x.lstZakaz
                    If белки Is Nothing Then
                        курс = Replace(b2.КурсВалЗак, ",", ".")
                        белки = b2.СуммаПоступленияЗак
                        dltZak = CDbl(Replace(b2.СуммаПоступленияЗак, ".", ","))
                    Else
                        курс = курс & vbCrLf & b2.КурсВалЗак
                        белки = белки & vbCrLf & b2.СуммаПоступленияЗак
                        dltZak += CDbl(Replace(b2.СуммаПоступленияЗак, ".", ","))
                    End If

                Next
                фрахт = x.ФрахтЗак
            Else

                For Each b7 In x.lstZakaz
                    If белки Is Nothing Then
                        курс = Replace(b7.КурсВалЗак, ",", ".")
                        белки = b7.СуммаПоступленияЗак
                        dltZak = CDbl(Replace(b7.СуммаПоступленияЗак, ".", ","))
                    Else
                        курс = курс & vbCrLf & b7.КурсВалЗак
                        белки = белки & vbCrLf & b7.СуммаПоступленияЗак
                        dltZak += CDbl(Replace(b7.СуммаПоступленияЗак, ".", ","))
                    End If
                Next
                фрахт = x.ФрахтЗак

            End If

            'выбираем даты оплат

            Dim ДатаОплаты As String = Nothing
            For Each b4 In x.lstZakaz
                If ДатаОплаты Is Nothing Then
                    ДатаОплаты = b4.ДатыОплатЗак
                Else
                    ДатаОплаты = ДатаОплаты & vbCrLf & b4.ДатыОплатЗак
                End If

            Next


            'выбираем оплаты перевозчику и фрахт
            Dim f9 = list8.Where(Function(c) c.НомерРейса = x.НомерРейса).Select(Function(c) c).FirstOrDefault()

            Dim фрахтПер As String = Nothing
            Dim курсПер As String = Nothing
            Dim белкиПер As String = Nothing
            If f9 Is Nothing Then Continue For

            If x.НомерРейса = 547 Then
                Dim mf As String = ""
            End If

            For Each b2 In f9.lstPerev
                If белкиПер Is Nothing Then
                    курсПер = Replace(b2.КурсВалПер, ",", ".")
                    белкиПер = Replace(b2.СуммаПоступленияПер, ",", ".")
                    dltPer = CDbl(Replace(b2.СуммаПоступленияПер, ".", ","))
                Else
                    курсПер = курсПер & vbCrLf & Replace(b2.КурсВалПер, ",", ".")
                    белкиПер = белкиПер & vbCrLf & Replace(b2.СуммаПоступленияПер, ",", ".")
                    dltPer += CDbl(Replace(b2.СуммаПоступленияПер, ".", ","))
                End If

            Next


            Dim Расч As String = Nothing
            If Driver.Валюта = "Рубль" Then
                фрахтПер = f9.ФрахтПер
                Расч = f9.ФрахтПер
            Else
                фрахтПер = f9.ФрахтПер
                Расч = dltPer
            End If







            Dim Дельта As Double = Math.Round(dltZak - CDbl(Replace(Расч, ".", ",")))




            If client IsNot Nothing Then
                Готовый = New GridReport With {.Счет = client.НомерРейса, .Заказчик = client.НазвОрганизации, .Загрузка = client.Маршрут, .ДатаЗагрузки = client.ДатаПодачиПодЗагрузку,
                    .ДатаВыгрузки = client.ДатаПодачиПодРастаможку, .BYR = белки, .Валюта = фрахт, .Курс = курс, .ДатаОплаты = ДатаОплаты, .Перевозчик = Driver.НазвОрганизации,
                    .BYRПер = белкиПер, .ВалютаПер = фрахтПер, .КурсПер = курсПер, .Дельта = Дельта}
                КоллекцияГотовая.Add(Готовый)
            End If


        Next

        'Добавляем отчет в базу или обновляем
        Using db As New dbAllDataContext(_cn3)
            Dim fb = (From x In db.ОтчетРаботыСотрудника
                      Where x.Год = ComboBox1.Text And x.Месяц = ComboBox2.Text
                      Select x).FirstOrDefault()
            ' удаляем старые данные добавляем новые
            If fb IsNot Nothing Then
                db.ОтчетРаботыСотрудника.DeleteOnSubmit(fb)
                db.SubmitChanges()
            End If


            Dim f17 As New ОтчетРаботыСотрудника
            f17.Год = ComboBox1.Text
            f17.Месяц = ComboBox2.Text
            db.ОтчетРаботыСотрудника.InsertOnSubmit(f17)
            db.SubmitChanges()
            mo.ОтчетРаботыСотрудникаAllAsync()
            AllClass.ОтчетРаботыСотрудника.Add(f17)

            Dim ID As Integer = f17.ID

            Dim f3 As New List(Of ОтчетРаботыСотрудникаСводная)
            For Each cv In КоллекцияГотовая
                Dim f2 As New ОтчетРаботыСотрудникаСводная
                With f2
                    .IDОтчетРабСотрудн = ID
                    .счет = cv.Счет
                    .заказчик = cv.Заказчик
                    .Загрузка = cv.Загрузка
                    .ДатаЗагрузки = cv.ДатаЗагрузки
                    .ДатаВыгрузки = cv.ДатаВыгрузки
                    .ИтогоЗакБелРуб = cv.BYR
                    .ВалютаЗак = cv.Валюта
                    .КурсЗак = cv.Курс
                    .ДатаОплатыЗак = cv.ДатаОплаты
                    .Перевозчик = cv.Перевозчик
                    .ИтогоПерБелРуб = cv.BYRПер
                    .ВалютаПер = cv.ВалютаПер
                    .КурсПер = cv.КурсПер
                    .Дельта = cv.Дельта
                    .ДатаСоздания = Now
                    .Экспедитор = Экспедитор
                End With
                f3.Add(f2)

            Next
            db.ОтчетРаботыСотрудникаСводная.InsertAllOnSubmit(f3)
            db.SubmitChanges()
            mo.ОтчетРаботыСотрудникаСводнаяAllAsync()
            AllClass.ОтчетРаботыСотрудникаСводная.AddRange(f3)
        End Using





        ExcelTableAsync(КоллекцияГотовая, ComboBox1.Text & " _ " & ComboBox2.Text)
        UpdList1()
        Me.Cursor = Cursors.Default
        CheckBox1.Checked = False
        ComboBox1.Text = String.Empty
        ComboBox2.Text = String.Empty






    End Sub
    Private Property mosR As New Stopwatch
    Private Sub Prgsbr(d As Boolean)
        If d = True Then
            mosR.Start()
            Do While mosR.IsRunning = True
                If mosR.Elapsed.TotalSeconds >= 60 Then
                    mosR.Stop()
                End If
                ProgressBar1.Value = Math.Round(mosR.Elapsed.TotalSeconds)
            Loop
        Else
            mosR.Stop()
        End If

    End Sub
    Private Async Sub ExcelTableAsync(ByVal ListClient As List(Of GridReport), ByVal NameSheet As String)

        ProgressBar1.Value = 0
        ProgressBar1.Maximum = 60
        ProgressBar1.Step = 1

        'Dim progress As New Progress(Of Integer)(Function(percent) ProgressBar1.Value = percent)

        Dim m As String = Await Task.Run(Function() ExcelTable(ListClient, NameSheet))
        If m IsNot Nothing Then
            ProgressBar1.Value = 60
            Process.Start(m)
            ProgressBar1.Value = 0
        End If

    End Sub

    Private Function ExcelTable(ListClient As List(Of GridReport), NamSheet As String)  'progress As IProgress(Of Integer)
        'Dim stopw As New Stopwatch
        'stopw.Start()
        'Me.Cursor = Cursors.WaitCursor
        ProgressBar1.BeginInvoke(New MethodInvoker(Sub() Prgsbr(True)))
        Dim op As New ПутьЗапускаПрограммы
        Dim Pathвремянка2 = op.ВремянкаДляРесурса()

        If IO.File.Exists(Pathвремянка2 & "\" & "Отчет1.xlsx") = False Then
            'копируем из ресурсов в папку general
            IO.File.WriteAllBytes(Pathвремянка2 & "\" & "Отчет1.xlsx", My.Resources.Report)
        Else
            IO.File.Delete(Pathвремянка2 & "\" & "Отчет1.xlsx")
            IO.File.WriteAllBytes(Pathвремянка2 & "\" & "Отчет1.xlsx", My.Resources.Report)
        End If
        Dim FileВремянка2 = Pathвремянка2 & "\" & "Отчет1.xlsx"



        Dim ecx As New ExcelDounloaded
        ecx.Excel2 = New Excel.Application()

        ecx.wb = ecx.Excel2.Workbooks.Add(FileВремянка2)
        ecx.ws = CType(ecx.wb.Sheets.Add(After:=ecx.wb.ActiveSheet), Excel.Worksheet)
        ecx.ws.Name = NamSheet

        'ecx.ws = ecx.wb.Worksheets.Add(ComboBox1.Text & " _ " & ComboBox2.Text)
        'ecx.ws.Activate()
        ecx.Excel2.Visible = False

        'Progress.Report(10)

        'ProgressBar1.Invoke(New Action(Sub() Prgsbr(stopw.)))

        With ecx.ws

            .Range(Cell1:="A1", Cell2:="A2").Merge()
            With .Range("a1")
                .Value = "счет №"

                .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                .VerticalAlignment = Excel.XlHAlign.xlHAlignCenter
            End With


            .Range(Cell1:="b1", Cell2:="b2").Merge()
            With .Range("b1")
                .Value = "Заказчик"
                .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                .VerticalAlignment = Excel.XlHAlign.xlHAlignCenter

            End With


            .Range(Cell1:="c1", Cell2:="c2").Merge()
            With .Range("c1")
                .Value = "Загр."
                .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                .VerticalAlignment = Excel.XlHAlign.xlHAlignCenter
            End With


            .Range(Cell1:="d1", Cell2:="d2").Merge()
            .Range("d1").Value = "Выгр."
            .Range("d1").HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
            .Range("d1").VerticalAlignment = Excel.XlHAlign.xlHAlignCenter

            .Range(Cell1:="e1", Cell2:="e2").Merge()
            .Range("e1").Value = "Дата загрузки"
            .Range("e1").HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
            .Range("e1").VerticalAlignment = Excel.XlHAlign.xlHAlignCenter

            .Range(Cell1:="f1", Cell2:="f2").Merge()
            .Range("f1").Value = "Дата выгрузки"
            .Range("f1").HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
            .Range("f1").VerticalAlignment = Excel.XlHAlign.xlHAlignCenter

            .Range(Cell1:="g1", Cell2:="j1").Merge()
            .Range("g1").Value = "Фрахт заказчика"
            .Range("g1").HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
            .Range("g1").VerticalAlignment = Excel.XlHAlign.xlHAlignCenter


            .Cells(2, 7) = "Валюта"
            .Range("g2").HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
            .Range("g2").VerticalAlignment = Excel.XlHAlign.xlHAlignCenter

            'Progress.Report(20)
            'ProgressBar1.Invoke(New Action(Sub() Prgsbr(20)))

            .Cells(2, 9) = "Курс"
            .Range("i2").HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
            .Range("i2").VerticalAlignment = Excel.XlHAlign.xlHAlignCenter

            .Cells(2, 10) = "BYR"
            .Range("j2").HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
            .Range("j2").VerticalAlignment = Excel.XlHAlign.xlHAlignCenter

            .Range(Cell1:="k1", Cell2:="k2").Merge()
            .Range("k1").Value = "Дата оплаты"
            .Range("k1").HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
            .Range("k1").VerticalAlignment = Excel.XlHAlign.xlHAlignCenter


            .Range(Cell1:="l1", Cell2:="l2").Merge()
            .Range("l1").Value = "Перевозчик"
            .Range("l1").HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
            .Range("l1").VerticalAlignment = Excel.XlHAlign.xlHAlignCenter

            .Cells(2, 13) = "Валюта"
            .Range("m2").HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
            .Range("m2").VerticalAlignment = Excel.XlHAlign.xlHAlignCenter

            .Cells(2, 15) = "Курс"
            .Range("o2").HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
            .Range("o2").VerticalAlignment = Excel.XlHAlign.xlHAlignCenter

            .Cells(2, 16) = "BYR"
            .Range("p2").HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
            .Range("p2").VerticalAlignment = Excel.XlHAlign.xlHAlignCenter



            .Range(Cell1:="m1", Cell2:="p1").Merge()
            .Range("m1").Value = "Фрахт перевозчика"
            .Range("m1").HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
            .Range("m1").VerticalAlignment = Excel.XlHAlign.xlHAlignCenter

            .Range(Cell1:="q1", Cell2:="q2").Merge()

            'ProgressBar1.BeginInvoke(New MethodInvoker(Sub() Prgsbr(30)))

            With .Range("q1")
                .Value = "комиссия за перевод"
                .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                .VerticalAlignment = Excel.XlHAlign.xlHAlignCenter
                .WrapText = True

            End With


            .Range(Cell1:="r1", Cell2:="r2").Merge()
            With .Range("r1")
                .Value = "Страхование груза"
                .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                .VerticalAlignment = Excel.XlHAlign.xlHAlignCenter
                .WrapText = True
            End With


            .Range(Cell1:="s1", Cell2:="s2").Merge()
            With .Range("s1")
                .Value = "Дата оплаты"
                .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                .VerticalAlignment = Excel.XlHAlign.xlHAlignCenter
                .WrapText = True
            End With


            .Range(Cell1:="t1", Cell2:="t2").Merge()
            With .Range("t1")
                .Value = "Войтик Екатерина"
                .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                .VerticalAlignment = Excel.XlHAlign.xlHAlignCenter
            End With



            .Range(Cell1:="A1", Cell2:="V" & ListClient.Count + 3).Cells.Borders.LineStyle = 1 'рисуем границы
            .Range(Cell1:="A1", Cell2:="V2").Cells.Font.Bold = True 'жирный шрифт
            '.Cells(dt.Rows.Count + 4, 2).Font.Color = Color.FromRgb(238, 245, 95)
            '.Range(Cell1:="A" & (dt.Rows.Count + 4), Cell2:="E" & (dt.Rows.Count + 4)).Merge()
            '.Range(Cell1:="A" & (dt.Rows.Count + 4), Cell2:="J" & (dt.Rows.Count + 4)).Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
            '.Range(Cell1:="A" & (dt.Rows.Count + 4), Cell2:="J" & (dt.Rows.Count + 4)).Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThick
            ''.Range(Cell1:="A" & (4), Cell2:="J" & (4)).Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThick
        End With
        ' ProgressBar1.Invoke(New Action(Sub() Prgsbr(40)))
        Dim AllSum As Double


        'ecx.ws.Range("A1:" & "V" & ListClient.Count + 5).Columns.AutoFit()

        For n As Integer = 0 To ListClient.Count - 1

            With ecx.ws.Range(Cell1:="C" & 4 + n, Cell2:="D" & 4 + n)
                .Merge()
                .WrapText = True
                .ColumnWidth = 30
                .RowHeight = 32
                .HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft 'выравнивает по горизонтали
                .VerticalAlignment = Excel.XlHAlign.xlHAlignCenter 'выравнивает по вертикали
            End With


            With ecx.ws.Range("A" & 4 + n)
                .Value2 = ListClient(n).Счет.ToString
                .Borders.LineStyle = Excel.XlLineStyle.xlContinuous
                .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                .VerticalAlignment = Excel.XlHAlign.xlHAlignCenter
            End With

            'ProgressBar1.Invoke(New Action(Sub() Prgsbr(50)))

            With ecx.ws.Range("b" & 4 + n)
                .Value = ListClient(n).Заказчик.ToString
                .ColumnWidth = 25
                .HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
                .VerticalAlignment = Excel.XlHAlign.xlHAlignCenter
                .WrapText = True
            End With



            With ecx.ws.Range("c" & 4 + n)
                .Value = ListClient(n).Загрузка.ToString
                .HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
                .VerticalAlignment = Excel.XlHAlign.xlHAlignCenter
            End With

            With ecx.ws.Range("e" & 4 + n)
                .Value = ListClient(n).ДатаЗагрузки.ToString
                .Borders.LineStyle = Excel.XlLineStyle.xlContinuous
                .ColumnWidth = 14
                .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                .VerticalAlignment = Excel.XlHAlign.xlHAlignCenter
            End With

            'ProgressBar1.Invoke(New Action(Sub() Prgsbr(60)))

            With ecx.ws.Range("f" & 4 + n)
                .Value = ListClient(n).ДатаВыгрузки.ToString
                .ColumnWidth = 14
                .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                .VerticalAlignment = Excel.XlHAlign.xlHAlignCenter
            End With

            With ecx.ws.Range("g" & 4 + n)
                .Value = ListClient(n).Валюта.ToString
                .ColumnWidth = 8
                .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                .VerticalAlignment = Excel.XlHAlign.xlHAlignCenter
            End With




            ecx.ws.Range("h" & 4 + n).ColumnWidth = 2


            With ecx.ws.Range("i" & 4 + n)
                .Value = ListClient(n).Курс
                .ColumnWidth = 8
                .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                .VerticalAlignment = Excel.XlHAlign.xlHAlignCenter
            End With


            With ecx.ws.Range("j" & 4 + n)
                .Value = ListClient(n).BYR.ToString
                .ColumnWidth = 8
                .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                .VerticalAlignment = Excel.XlHAlign.xlHAlignCenter
            End With

            With ecx.ws.Range("k" & 4 + n)
                .Value = Strings.Left(ListClient(n).ДатаОплаты.ToString, 10)
                .ColumnWidth = 14
                .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                .VerticalAlignment = Excel.XlHAlign.xlHAlignCenter
            End With

            'ProgressBar1.Invoke(New Action(Sub() Prgsbr(70)))

            With ecx.ws.Range("l" & 4 + n)
                .Value = ListClient(n).Перевозчик.ToString
                .ColumnWidth = 25
                .HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
                .VerticalAlignment = Excel.XlHAlign.xlHAlignCenter
                .WrapText = True
            End With


            With ecx.ws.Range("m" & 4 + n)
                .Value = ListClient(n).ВалютаПер.ToString
                .ColumnWidth = 8
                .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                .VerticalAlignment = Excel.XlHAlign.xlHAlignCenter
            End With

            ecx.ws.Range("n" & 4 + n).ColumnWidth = 2

            With ecx.ws.Range("o" & 4 + n)
                .Value = ListClient(n).КурсПер
                .ColumnWidth = 8
                .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                .VerticalAlignment = Excel.XlHAlign.xlHAlignCenter
            End With

            With ecx.ws.Range("p" & 4 + n)
                .Value = ListClient(n).BYRПер?.ToString
                .ColumnWidth = 8
                .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                .VerticalAlignment = Excel.XlHAlign.xlHAlignCenter
            End With

            'ProgressBar1.Invoke(New Action(Sub() Prgsbr(80)))


            ecx.ws.Range("q" & 4 + n).Value = ListClient(n).Комиссия
            ecx.ws.Range("q" & 4 + n).ColumnWidth = 9


            ecx.ws.Range("r" & 4 + n).Value = ListClient(n).Страхование
            ecx.ws.Range("r" & 4 + n).ColumnWidth = 9

            With ecx.ws.Range("t" & 4 + n)
                .Value = ListClient(n).Дельта.ToString
                .ColumnWidth = 17
                .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                .VerticalAlignment = Excel.XlHAlign.xlHAlignCenter
                '.Font.Color = Color.Yellow
                .Interior.Color = Color.LightGreen
            End With

            'ecx.ws.Range("s4" + n).Value = ListClient(n).Счет.ToString

            AllSum += ListClient(n).Дельта
        Next

        'ProgressBar1.Invoke(New Action(Sub() Prgsbr(90)))

        With ecx.ws
            .Range("V5").Value = AllSum
            .Range("V9").Value = AllSum * 40 / 100
            .Range("U9").Value = "ЗП"

            .Range("V" & ListClient.Count + 5).Value = AllSum * 40 / 100
            .Range("U" & ListClient.Count + 5).Value = "Долг"

            .Range("B" & ListClient.Count + 5).Value = "Количество оплаченных рейсов - " & ListClient.Count
            Dim kj = ListClient.GroupBy(Function(c) c.Заказчик, Function(c) c.Дельта,
                                Function(off, sal) New With {.key = off, .sumSal = sal.Count(), .ап = sal.Sum()})


            Dim ip As Integer = 1
            For Each bu In kj
                .Range("B" & ListClient.Count + 5 + ip).Value = bu.key & " - " & bu.sumSal & ". Средний заработок - " & Math.Round((CDbl(bu.ап) / CDbl(bu.sumSal)), 0)
                ip += 1
            Next
        End With



        'альбомная ориентация страницы
        ecx.ws.PageSetup.Orientation = Excel.XlPageOrientation.xlPortrait




        ''страничный режим
        'ActiveWindow.View = xlPageBreakPreview


        Dim Времянка As New ПутьЗапускаПрограммы(False)
        Dim Путь = Времянка._ПутьЗапускаПрограммы & "\New(" & Replace(Now.ToShortTimeString, ":", "-") & ").xlsx"
        ecx.wb.SaveAs(Путь)

        ecx.wb.Close()
        ecx.Excel2.Quit()



        releaseobject(ecx.ws)
        releaseobject(ecx.wb)
        releaseobject(ecx.Excel2)

        'Me.Cursor = Cursors.Default

        'Process.Start(Путь)

        'Dim mhj = ProgressBar1.EndInvoke(Nothing)
        If mosR.IsRunning = True Then
            mosR.Stop()
        End If


        Return Путь


    End Function
    Private Sub releaseobject(ByVal obj As Object)
        Try
            Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim myStream As Stream = Nothing
        Dim f As New OpenFileDialog()
        f.InitialDirectory = "c:\"
        f.Filter = "xml files (*.xml)|*.xml|All files (*.*)|*.*"
        f.FilterIndex = 0
        f.DefaultExt = "xml"
        f.RestoreDirectory = True

        Dim obser As New List(Of XmlVypiskaBankClass)

        If f.ShowDialog() = DialogResult.OK Then



            '
            If Not String.Equals(Path.GetExtension(f.FileName), ".xml", StringComparison.OrdinalIgnoreCase) Then

                MessageBox.Show("Невозможно прочитать файл. Не тот формат!", Рик)
                Return
            Else
                Try
                    myStream = f.OpenFile()
                    If myStream IsNot Nothing Then
                        Dim xdoc As XDocument = XDocument.Load(myStream)
                        Label4.Text = xdoc.Element("ROOT").Elements("QUERY").Elements("OUTPUT").Elements("Header1").Value
                        Dim i As Integer = 1
                        For Each x As XElement In xdoc.Element("ROOT").Elements("QUERY").Elements("OUTPUT").Elements("DOC")

                            'Dim fg = x.Attribute("Db").Value
                            'Dim hj = x.Attribute("Credit").Value
                            'Dim hj1 = x.Attribute("Docdate").Value
                            'Dim hj = x.Attribute("Credit").Value

                            Dim nb As New XmlVypiskaBankClass With {.Номер = i, .Организация = x.Element("KorName").Value,
                                .Приход = x.Attribute("Credit").Value,
                                .Расход = x.Attribute("Db").Value,
                                .Дата = x.Attribute("DocDate").Value,
                                .Назначение = x.Element("Nazn").Value,
                                .Назначение2 = x.Element("Nazn2").Value}

                            obser.Add(nb)
                            i += 1
                            'Dim attr As XAttribute = x.Attribute("Id")
                            'Dim NumCode As XElement = x.Element("NumCode")
                            'Dim CharCode As XElement = x.Element("CharCode")
                            'Dim Scale As XElement = x.Element("Scale")
                            'Dim Name As XElement = x.Element("Name")
                            'Dim Rate As XElement = x.Element("Rate")

                            'Dim list As New List(Of String) From {attr, NumCode, CharCode, Scale, Name, Rate}



                        Next

                        Grid1.DataSource = obser
                        GridView(Grid1)
                        Grid1.Columns(5).Width = 250
                        Grid1.Columns(0).Width = 30
                        Grid1.Columns(1).Width = 40
                        Grid1.Columns(2).Width = 40
                        Grid1.Columns(3).Width = 50
                        Grid1.Columns(6).Width = 80
                    End If


                Catch Ex As Exception
                    MessageBox.Show("Невозможно прочитать файл. Ошибка: " & Ex.Message)
                Finally
                    ' Check this again, since we need to make sure we didn't throw an exception on open.
                    If (myStream IsNot Nothing) Then
                        myStream.Close()
                    End If
                End Try
            End If             '
        End If




    End Sub
    Public Property ListSelect As String
    Private Sub ListBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox1.SelectedIndexChanged
        ListSelect = ListBox1.SelectedItem

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If ListSelect.Length > 0 Then
            Me.Cursor = Cursors.WaitCursor
            ComboBox1.Text = String.Empty
            ComboBox2.Text = String.Empty

            Using db As New dbAllDataContext(_cn3)
                Dim fb = (From x In db.ОтчетРаботыСотрудника
                          Join y In db.ОтчетРаботыСотрудникаСводная On x.ID Equals y.IDОтчетРабСотрудн
                          Where x.ID = CDbl(Strings.Left(ListSelect, 3))
                          Select x, y).ToList()
                Dim collection1 As New List(Of GridReport)
                If fb.Count > 0 Then
                    For Each client In fb
                        Dim g = New GridReport With {.Счет = client.y.счет, .Заказчик = client.y.заказчик, .Загрузка = client.y.Загрузка, .ДатаЗагрузки = client.y.ДатаЗагрузки,
                        .ДатаВыгрузки = client.y.ДатаВыгрузки, .BYR = client.y.ИтогоЗакБелРуб, .Валюта = client.y.ВалютаЗак, .Курс = client.y.КурсЗак,
                        .ДатаОплаты = client.y.ДатаОплатыЗак, .Перевозчик = client.y.Перевозчик, .BYRПер = client.y.ИтогоПерБелРуб,
                        .ВалютаПер = client.y.ВалютаПер, .КурсПер = client.y.КурсПер, .Дельта = client.y.Дельта}
                        collection1.Add(g)
                    Next

                    ExcelTableAsync(collection1, ListSelect)
                    Me.Cursor = Cursors.Default

                End If
            End Using
        End If
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        If ComboBox1.Text.Length = 0 Or ComboBox2.Text.Length = 0 Then
            Return
        End If

        Do While AllClass.РейсыПеревозчика Is Nothing
            Dim mo As New AllUpd
            mo.РейсыПеревозчикаAll()
        Loop

        Dim comb = ComboBox2.SelectedIndex + 1
        Dim sum As Double = 0

        Dim m = CDate("01." & comb & "." & ComboBox1.Text)
        Dim m2 = CDate(DateTime.DaysInMonth(ComboBox1.Text, comb) & "." & comb & "." & ComboBox1.Text)


        If grd1.Count > 0 Then
            grd1.Clear()
        End If
        Dim f3 As New List(Of Grid1Class)


        Using db As New DbAllDataContext(_cn3)
            Dim f = (From x In db.РейсыКлиента
                     Order By x.НомерРейса
                     Where x.ДатаСоздания >= m And x.ДатаСоздания <= m2
                     Select x).ToList()

            For Each b In f
                Try
                    Dim f2 = CDate(b.ДатаПодачиПодЗагрузку)

                    If m <= f2 And m2 >= f2 Then
                        f3.Add(New Grid1Class With {.дата_Загрузки = b.ДатаПодачиПодЗагрузку, .Клиент = b.НазвОрганизации,
                               .маршрут = b.Маршрут, .номер_Рейса = b.НомерРейса, .ставка_Клиента = b.СтоимостьФрахта, .Перевозчик = (From x In AllClass.РейсыПеревозчика
                                                                                                                                      Where x.НомерРейса = b.НомерРейса
                                                                                                                                      Select x.НазвОрганизации).FirstOrDefault(), .ставка_Перевозчика = (From x In AllClass.РейсыПеревозчика
                                                                                                                                                                                                         Where x.НомерРейса = b.НомерРейса
                                                                                                                                                                                                         Select x.СтоимостьФрахта).FirstOrDefault(), .дельта = Math.Round(CType(Replace(.ставка_Клиента, ".", ",") - Replace(.ставка_Перевозчика, ".", ","), Double), 2)})
                    End If
                Catch ex As Exception
                    f3.Add(New Grid1Class With {.дата_Загрузки = b.ДатаПодачиПодЗагрузку, .Клиент = b.НазвОрганизации,
                               .маршрут = b.Маршрут, .номер_Рейса = b.НомерРейса, .ставка_Клиента = b.СтоимостьФрахта, .Перевозчик = (From x In AllClass.РейсыПеревозчика
                                                                                                                                      Where x.НомерРейса = b.НомерРейса
                                                                                                                                      Select x.НазвОрганизации).FirstOrDefault(), .ставка_Перевозчика = (From x In AllClass.РейсыПеревозчика
                                                                                                                                                                                                         Where x.НомерРейса = b.НомерРейса
                                                                                                                                                                                                         Select x.СтоимостьФрахта).FirstOrDefault(), .дельта = Math.Round(CType(Replace(.ставка_Клиента, ".", ",") - Replace(.ставка_Перевозчика, ".", ","), Double), 2)})
                End Try
            Next


            Dim i As Integer = 1
            For Each b In f3
                b.Номер = i
                i += 1
                sum += CDbl(b.дельта)
                grd1.Add(b)
            Next
        End Using

        Grid1.Columns(5).HeaderText = "Дата загрузки"

        Dim mk As New Grid1Class With {.Клиент = "Итого", .дельта = sum}
        grd1.Add(mk)
        bsgrd1.ResetBindings(False)
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click

        If ComboBox1.Text.Length = 0 Or ComboBox2.Text.Length = 0 Then
            Return
        End If


        Dim mo As New AllUpd
        mo.РейсыПеревозчикаAll()


        Dim comb = ComboBox2.SelectedIndex + 1
        Dim sum As Double = 0

        Dim m = CDate("01." & comb & "." & ComboBox1.Text)
        Dim m2 = CDate(DateTime.DaysInMonth(ComboBox1.Text, comb) & "." & comb & "." & ComboBox1.Text)


        If grd1.Count > 0 Then
            grd1.Clear()
        End If
        Dim f3 As New List(Of Grid1Class)


        Using db As New DbAllDataContext(_cn3)
            Dim f = (From x In db.РейсыКлиента
                     Order By x.НомерРейса
                     Where x.ДатаОплаты >= m And x.ДатаОплаты <= m2
                     Select x).ToList()

            For Each b In f

                f3.Add(New Grid1Class With {.дата_Загрузки = b.ДатаОплаты, .Клиент = b.НазвОрганизации,
                               .маршрут = b.Маршрут, .номер_Рейса = b.НомерРейса, .ставка_Клиента = b.СтоимостьФрахта, .Перевозчик = (From x In AllClass.РейсыПеревозчика
                                                                                                                                      Where x.НомерРейса = b.НомерРейса
                                                                                                                                      Select x.НазвОрганизации).FirstOrDefault(), .ставка_Перевозчика = (From x In AllClass.РейсыПеревозчика
                                                                                                                                                                                                         Where x.НомерРейса = b.НомерРейса
                                                                                                                                                                                                         Select x.СтоимостьФрахта).FirstOrDefault(), .дельта = Math.Round(CType(Replace(.ставка_Клиента, ".", ",") - Replace(.ставка_Перевозчика, ".", ","), Double), 2)})

            Next


            Dim i As Integer = 1
            For Each b In f3
                b.Номер = i
                i += 1
                sum += CDbl(b.дельта)
                grd1.Add(b)
            Next
        End Using


        Grid1.Columns(5).HeaderText = "Дата оплаты"
        Dim mk As New Grid1Class With {.Клиент = "Итого", .дельта = sum}
        grd1.Add(mk)
        bsgrd1.ResetBindings(False)
    End Sub
End Class
Public Class MonthItem
    Public Property Id As Integer
    Public Property Text As String
End Class
Public Class DannPerev
    Public Property КурсВалПер As String
    Public Property СуммаПоступленияПер As String
    Public Property ДатыОплатПер As String
End Class
Public Class DannZakaz
    Public Property КурсВалЗак As String
    Public Property СуммаПоступленияЗак As String
    Public Property ДатыОплатЗак As String
End Class
Public Class NumberReysSort
    Public Property НомерРейса As Integer
    Public Property СуммаПоступления As String
    Public Property ДатаОплаты As String
    Public Property ФрахтПер As String
    Public Property ФрахтЗак As String
    Public Property lstZakaz As List(Of DannZakaz)
    Public Property lstPerev As List(Of DannPerev)

End Class
Public Class XmlVypiskaBankClass
    Public Property Номер As Integer
    Public Property Приход As String
    Public Property Расход As String
    Public Property Дата As String
    Public Property Организация As String
    Public Property Назначение As String
    Public Property Назначение2 As String




End Class
Public Class GridReport
    Public Property Счет As Integer
    Public Property Заказчик As String
    Public Property Загрузка As String
    Public Property Выгрузка As String
    Public Property ДатаЗагрузки As String
    Public Property ДатаВыгрузки As String

    Public Property Валюта As String
    Public Property Курс As String
    Public Property BYR As String
    Public Property ДатаОплаты As String

    Public Property Перевозчик As String

    Public Property ВалютаПер As String
    Public Property КурсПер As String
    Public Property BYRПер As String

    Public Property Комиссия As String
    Public Property Страхование As String
    Public Property Дельта As Double

    Public Property Итого As String
    Public Property ЗП As String

    'Public Property Страхование As String

    'Public Property Страхование As String
End Class
Public Class ExcelDounloaded
    Public Property Excel2 As Excel.Application

    Public Property wb As Excel.Workbook
    Public Property missing As Object = Type.Missing
    Public Property ws As Excel.Worksheet
    Public Property rng As Excel.Range
End Class
