Imports System.Collections.ObjectModel
Imports System.ComponentModel
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

        'ListBox1.Items.Clear()

        Using db As New dbAllDataContext()
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





    End Sub
    Private Sub UpdList1()
        LS2.Clear()

        Using db As New dbAllDataContext()
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

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If ComboBox1.Text.Length = 0 Or ComboBox2.Text.Length = 0 Then
            Return
        End If

        Me.Cursor = Cursors.WaitCursor

        'проверяем, надо ли пересобирать сборку
        If CheckBox1.Checked = False Then
            'проверяем,есть ли отчет в базе?
            Using db As New dbAllDataContext()
                Dim fb = (From x In db.ОтчетРаботыСотрудника
                          Join y In db.ОтчетРаботыСотрудникаСводная On x.ID Equals y.IDОтчетРабСотрудн
                          Where x.Год = ComboBox1.Text And x.Месяц = ComboBox2.Text
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

                    ExcelTable(collection1)
                    Me.Cursor = Cursors.Default
                    Exit Sub
                End If
            End Using

        End If






        Dim startDate = New DateTime(ComboBox1.Text, ComboBox2.SelectedIndex + 1, 1)
        Dim enddate = startDate.AddMonths(1).AddDays(-1)










        'выгружаем данные всех рейсов клиентов в коллекцию
        Dim obser As New List(Of РейсыКлиента)
        Dim ObserSelect As New List(Of РейсыКлиента)

        Using db As New dbAllDataContext()
            Dim f = db.РейсыКлиента.Select(Function(x) x)
            If f IsNot Nothing Then
                For Each x In f
                    obser.Add(x)
                Next
            End If
        End Using

        Dim AllPayClient As List(Of ОплатыКлиент)
        Dim list8 As New List(Of NumberReysSort) 'отобранные фрахты перевозчика, если валютные оплачены полностью то попадают в белках в отчет, иначе в валюте
        Dim list5 As New List(Of NumberReysSort) 'отобранные рейсы которые были полностью оплачены в выбранном месяце и попали в отчет
        Dim list2 As New List(Of NumberReysSort)
        Dim ObserSelectПер1 As New List(Of РейсыПеревозчика)


        'выбираем олпаты от клиентов за выбранный месяц и год
        Using db As New dbAllDataContext()

            AllPayClient = db.ОплатыКлиент.Select(Function(x) x).ToList()

            Dim f = (From x In (From y In db.ОплатыКлиент
                                Where (CType(y.ДатаОплаты, Date)).Year = ComboBox1.Text
                                Select y).ToList()
                     Where MonthName(Month(x.ДатаОплаты)) = ComboBox2.Text
                     Select x).ToList()
            If f.Count > 0 Then
                'выбираем рейсы безповтора
                Dim f1 = f.Select(Function(x) x.Рейс).Distinct().ToList()
                ' Dim f2 = f.GroupBy(Function(x) x.Рейс, Function(x) x.Сумма, Function(_nom, suma) (Nm:=_nom, salsum:=suma.Sum()))

                'все оплаты клиентов до выбранной даты
                Dim f2 = db.ОплатыКлиент.Where(Function(c) c.ДатаОплаты <= enddate).Select(Function(c) c).ToList()

                'группируем и сумируем оплаты с нарастающим итогом до выбранной даты
                Dim f3 = (From x In f2
                          Group x By Keys = New With {Key x.Рейс}
                             Into Group
                          Select New With {.Рейс = Keys.Рейс, .Сумма = Group.Sum(Function(t) t.Сумма)}).OrderBy(Function(x) x.Рейс).ToList()




                'отбираем только рейсы выбранного месяца в коллекцию
                For Each u In f1
                    For Each b In f3
                        If b.Рейс.ToString.Contains(u) Then
                            list2.Add(New NumberReysSort With {.НомерРейса = b.Рейс, .СуммаПоступления = b.Сумма})

                        End If
                    Next
                Next



                'сравнивает стоимость оплаты с фрахтом по каждому рейсу
                For Each r In obser
                    For Each z In list2
                        'если оплата в белках
                        If r.НомерРейса = z.НомерРейса And r.СтоимостьФрахта = z.СуммаПоступления And (r.Валюта = "Рубль") Then
                            list5.Add(New NumberReysSort With {.НомерРейса = z.НомерРейса, .СуммаПоступления = z.СуммаПоступления})
                        ElseIf r.НомерРейса = z.НомерРейса And Not r.Валюта = "Рубль" Then 'если была оплата  в белках по курсу

                            'если оплата в белках по курсу
                            If r.ВалютаПлатежа = "BYN" Then
                                Dim Obsr2 = f2.Where(Function(x) x.Рейс = z.НомерРейса).Select(Function(x) x).ToList()
                                Dim валюта As String = r.Валюта
                                Dim summa As Double = 0
                                Dim курс As Double
                                For Each x In Obsr2
                                    ' получаем курс валюты
                                    Dim m As New NbRbClassNew()
                                    курс = Replace(m.Курс(x.ДатаОплаты, валюта), ".", ",")
                                    'делим сумму прихода на курс получаем в валюте
                                    summa += x.Сумма / курс

                                Next
                                'если полученная стоимость равна фрахту в рейсе то попадает в список
                                If Math.Round(summa, 0) = CDbl(r.СтоимостьФрахта) Then
                                    list5.Add(New NumberReysSort With {.НомерРейса = z.НомерРейса, .СуммаПоступления = summa})
                                End If

                            Else
                                'если опалта в валюте
                                If r.СтоимостьФрахта = z.СуммаПоступления Then
                                    list5.Add(New NumberReysSort With {.НомерРейса = z.НомерРейса, .СуммаПоступления = z.СуммаПоступления})
                                End If

                            End If
                        End If
                    Next
                Next


            End If

            'проверяем оплату перевозчика

            'все рейсы перевозчика

            Dim f7 = db.РейсыПеревозчика.Select(Function(x) x)
            If f IsNot Nothing Then
                For Each x In f7
                    ObserSelectПер1.Add(x)
                Next
            End If

            Dim pl = db.ОплатыПер.Select(Function(c) c).ToList()

            For Each x In ObserSelectПер1
                For Each y In list2
                    'если номер рейса нашли и оплата белки
                    If x.НомерРейса = y.НомерРейса And x.Валюта = "Рубль" Then
                        list8.Add(New NumberReysSort With {.НомерРейса = x.НомерРейса, .СуммаПоступления = x.СтоимостьФрахта})
                    Else

                        For Each n In pl

                            'если оплата по этому рейсу произведена, то
                            If x.НомерРейса = n.Рейс Then

                                'выбираем все оплаты по этому рейсу перевозчика
                                Dim f9 = db.ОплатыПер.Where(Function(b) b.Рейс = x.НомерРейса).Select(Function(b) b).ToList()
                                'если валюта платежа белки по курсу
                                If x.ВалютаПлатежа = "BYN" Then

                                    Dim валюта As String = x.Валюта
                                    Dim summa As Double = 0
                                    Dim курс As Double
                                    For Each x1 In f9
                                        Dim m As New NbRbClassNew()
                                        курс = Replace(m.Курс(x1.ДатаОплаты, валюта), ".", ",")

                                        summa += x1.Сумма / курс

                                    Next

                                    'если оплата была полной, то добалвяем в список
                                    If Math.Round(summa, 0) = CDbl(x.СтоимостьФрахта) Then
                                        list8.Add(New NumberReysSort With {.НомерРейса = x.НомерРейса, .СуммаПоступления = summa})
                                    End If

                                End If
                            End If
                            Exit For
                        Next
                    End If
                Next
            Next


        End Using

        'Отбираем все данные в конечную коллекцию для вставки в эксель
        Dim КоллекцияГотовая As New List(Of GridReport)
        Dim Готовый As GridReport
        For Each x In list5

            Dim Driver = ObserSelectПер1.Where(Function(c) c.НомерРейса = x.НомерРейса).Select(Function(c) c).FirstOrDefault()

            Dim client As РейсыКлиента
            Using db As New dbAllDataContext
                client = db.РейсыКлиента.Where(Function(c) c.НомерРейса = x.НомерРейса).Select(Function(c) c).FirstOrDefault()
            End Using


            Dim фрахт As Double
            Dim курс As Double
            Dim белки As Double

            If client Is Nothing Then Continue For

            If client.Валюта = "Рубль" Then
                фрахт = x.СуммаПоступления
                курс = 1
                белки = x.СуммаПоступления
            Else
                'выбираею даты поступления денег от заказчика при оплате в белках по курсу
                Dim ku As List(Of ОплатыКлиент)
                If client.ВалютаПлатежа = "BYN" Then
                    Using db As New dbAllDataContext()
                        ku = db.ОплатыКлиент.Where(Function(c) c.Рейс = x.НомерРейса).Select(Function(c) c).ToList()
                    End Using
                    If ku.Count > 0 Then
                        For Each b In ku
                            Dim fk = CDbl(b.Сумма)
                            белки += fk
                        Next

                    End If
                End If




                'сумма оплаты делим на сумму в валюте

                фрахт = client.СтоимостьФрахта
                курс = Math.Round(белки / фрахт, 4)

            End If

            'выбираем даты оплат
            Dim ДатаОплаты As String = Nothing
            If AllPayClient.Count > 0 Then
                Dim fg = AllPayClient.Where(Function(b) b.Рейс = client.НомерРейса).Select(Function(b) b).ToList()
                If fg.Count > 0 Then
                    For Each l In fg
                        ДатаОплаты &= Strings.Left(l.ДатаОплаты & vbCrLf, 10)
                    Next

                End If
            End If

            'выбираем оплаты перевозчику и фрахт
            Dim фрахтПер As Double
            Dim курсПер As Double
            Dim белкиПер As Double

            If Driver.Валюта = "Рубль" Then
                фрахтПер = Driver.СтоимостьФрахта
                курсПер = 1
                белкиПер = Driver.СтоимостьФрахта
            Else
                If Driver.ВалютаПлатежа = "BYN" Then
                    'сумма оплаты делим на сумму в валюте, если оплата в белках
                    Dim nm As List(Of ОплатыПер)
                    Using db As New dbAllDataContext()
                        nm = db.ОплатыПер.Where(Function(c) c.Рейс = x.НомерРейса).Select(Function(c) c).ToList()
                    End Using
                    If nm.Count > 0 Then
                        For Each vx In nm
                            белкиПер += vx.Сумма
                        Next

                    End If
                    фрахтПер = Driver.СтоимостьФрахта
                    курсПер = Math.Round(белкиПер / фрахтПер, 4)
                Else
                    фрахтПер = Driver.СтоимостьФрахта
                    курсПер = 1
                    белкиПер = Driver.СтоимостьФрахта
                End If


            End If

            Dim Дельта As Double = Math.Round(белки - белкиПер)




            If client IsNot Nothing Then
                Готовый = New GridReport With {.Счет = client.НомерРейса, .Заказчик = client.НазвОрганизации, .Загрузка = client.Маршрут, .ДатаЗагрузки = client.ДатаПодачиПодЗагрузку,
                    .ДатаВыгрузки = client.ДатаПодачиПодРастаможку, .BYR = белки, .Валюта = фрахт, .Курс = курс, .ДатаОплаты = ДатаОплаты, .Перевозчик = Driver.НазвОрганизации,
                    .BYRПер = белкиПер, .ВалютаПер = фрахтПер, .КурсПер = курсПер, .Дельта = Дельта}
                КоллекцияГотовая.Add(Готовый)
            End If


        Next

        'Добавляем отчет в базу или обновляем
        Using db As New dbAllDataContext()
            Dim fb = (From x In db.ОтчетРаботыСотрудника
                      Where x.Год = ComboBox1.Text And x.Месяц = ComboBox2.Text
                      Select x).FirstOrDefault()
            ' удаляем старые данные добавляем новые
            If fb IsNot Nothing Then
                db.ОтчетРаботыСотрудника.DeleteOnSubmit(fb)
                db.SubmitChanges()
            End If


            Dim f As New ОтчетРаботыСотрудника
            f.Год = ComboBox1.Text
            f.Месяц = ComboBox2.Text
            db.ОтчетРаботыСотрудника.InsertOnSubmit(f)
            db.SubmitChanges()

            Dim ID As Integer = f.ID

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
                End With
                f3.Add(f2)

            Next
            db.ОтчетРаботыСотрудникаСводная.InsertAllOnSubmit(f3)
            db.SubmitChanges()

        End Using





        ExcelTable(КоллекцияГотовая)
        UpdList1()
        Me.Cursor = Cursors.Default
        CheckBox1.Checked = False
        ComboBox1.Text = String.Empty
        ComboBox2.Text = String.Empty
        Exit Sub






        '------------------------------------------------------------------   раздел перевозчика

        Dim list3 As New List(Of NumberReysSort) 'сгрупированные и сумированные рейсы по выбранному месяцу но без проверки с фрахтом рейса
        Dim list4 As New List(Of NumberReysSort) 'отобранные рейсы которые были полностью оплачены в выбранном месяце и попали в отчет
        'выгружаем данные всех рейсов перевозчиков в коллекцию
        Dim obserПер As New List(Of РейсыПеревозчика)
        Dim ObserSelectПер As New List(Of РейсыПеревозчика)
        Using db As New dbAllDataContext()
            Dim f = db.РейсыПеревозчика.Select(Function(x) x)
            If f IsNot Nothing Then
                For Each x In f
                    obserПер.Add(x)
                Next
            End If
        End Using
        'выбираем олпаты перевозчикам за выбранный месяц и год
        Using db As New dbAllDataContext()
            Dim f = (From x In (From y In db.ОплатыПер
                                Where (CType(y.ДатаОплаты, Date)).Year = ComboBox1.Text
                                Select y).ToList()
                     Where MonthName(Month(x.ДатаОплаты)) = ComboBox2.Text
                     Select x).ToList()
            If f.Count > 0 Then
                'выбираем рейсы безповтора
                Dim f1 = f.Select(Function(x) x.Рейс).Distinct().ToList()
                ' Dim f2 = f.GroupBy(Function(x) x.Рейс, Function(x) x.Сумма, Function(_nom, suma) (Nm:=_nom, salsum:=suma.Sum()))


                ' выбираем все оплаты до конца даты выбранного месяца
                Dim f2 = db.ОплатыПер.Where(Function(c) c.ДатаОплаты <= enddate).Select(Function(c) c).ToList()


                'групируем и складываем суммы оплат всех рейсов
                Dim f3 = (From x In f2
                          Group x By Keys = New With {Key x.Рейс}
                             Into Group
                          Select New With {.Рейс = Keys.Рейс, .Сумма = Group.Sum(Function(t) t.Сумма)}).OrderBy(Function(x) x.Рейс).ToList()


                'отбираем только рейсы выбранного месяца в коллекцию
                For Each u In f1
                    For Each b In f3
                        If b.Рейс.ToString.Contains(u) Then
                            list3.Add(New NumberReysSort With {.НомерРейса = b.Рейс, .СуммаПоступления = b.Сумма})

                        End If
                    Next
                Next


                ' проаеряем вся-ли стоимость оплачена перевозчику (сверяем со cтоимостью указанной в рейсе  - фрахт)
                For Each r In obserПер
                    For Each z In list3
                        'если оплата в белках
                        If r.НомерРейса = z.НомерРейса And r.СтоимостьФрахта = z.СуммаПоступления And (r.Валюта = "Рубль") Then
                            list4.Add(New NumberReysSort With {.НомерРейса = z.НомерРейса, .СуммаПоступления = z.СуммаПоступления})
                        ElseIf r.НомерРейса = z.НомерРейса And Not r.Валюта = "Рубль" Then 'если была оплата перевозу в белках по валютному рейсу

                            'если оплата в белках по курсу
                            If r.ВалютаПлатежа = "BYN" Then
                                Dim Obsr2 = f2.Where(Function(x) x.Рейс = z.НомерРейса).Select(Function(x) x).ToList()
                                Dim валюта As String = r.Валюта
                                Dim summa As Double = 0
                                Dim курс As Double
                                For Each x In Obsr2
                                    Dim m As New NbRbClassNew()
                                    курс = Replace(m.Курс(x.ДатаОплаты, валюта), ".", ",")

                                    summa += x.Сумма / курс

                                Next

                                If Math.Round(summa, 0) = CDbl(r.СтоимостьФрахта) Then
                                    list4.Add(New NumberReysSort With {.НомерРейса = z.НомерРейса, .СуммаПоступления = summa})
                                End If

                            Else
                                'если опалта в валюте
                                If r.СтоимостьФрахта = z.СуммаПоступления Then
                                    list4.Add(New NumberReysSort With {.НомерРейса = z.НомерРейса, .СуммаПоступления = z.СуммаПоступления})
                                End If

                            End If
                        End If
                    Next
                Next

            End If
        End Using



    End Sub

    Private Sub ExcelTable(ByVal ListClient As List(Of GridReport))

        Me.Cursor = Cursors.WaitCursor

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
        ecx.ws.Name = ComboBox1.Text & " _ " & ComboBox2.Text

        'ecx.ws = ecx.wb.Worksheets.Add(ComboBox1.Text & " _ " & ComboBox2.Text)
        'ecx.ws.Activate()
        ecx.Excel2.Visible = False

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
        Dim AllSum As Double
        Dim Zp As Double

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
                .Value = ListClient(n).BYRПер.ToString
                .ColumnWidth = 8
                .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                .VerticalAlignment = Excel.XlHAlign.xlHAlignCenter
            End With




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

        Me.Cursor = Cursors.Default

        Process.Start(Путь)






    End Sub
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

            Using db As New dbAllDataContext()
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

                    ExcelTable(collection1)
                    Me.Cursor = Cursors.Default

                End If
            End Using
        End If
    End Sub
End Class
Public Class MonthItem
    Public Property Id As Integer
    Public Property Text As String
End Class
Public Class NumberReysSort
    Public Property НомерРейса As Integer
    Public Property СуммаПоступления As String
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

    Public Property Валюта As Double
    Public Property Курс As Double
    Public Property BYR As Double
    Public Property ДатаОплаты As String

    Public Property Перевозчик As String

    Public Property ВалютаПер As Double
    Public Property КурсПер As Double
    Public Property BYRПер As Double

    Public Property Комиссия As String
    Public Property Страхование As String
    Public Property Дельта As Double

    Public Property Итого As Double
    Public Property ЗП As Double
End Class
Public Class ExcelDounloaded
    Public Property Excel2 As Excel.Application

    Public Property wb As Excel.Workbook
    Public Property missing As Object = Type.Missing
    Public Property ws As Excel.Worksheet
    Public Property rng As Excel.Range
End Class
