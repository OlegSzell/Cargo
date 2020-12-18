Imports System.ComponentModel

Public Class ЖурналВсплывДанныеОбщиеДляГрид2
    Private grid1All As List(Of Grid1ЖурналClass)
    Private bsgrid1 As BindingSource
    Private lst1 As List(Of String)
    Private bslst1 As BindingSource
    Private lst2 As BindingList(Of String)
    Private bslst2 As BindingSource
    Dim D As Grid2ЖурналClass
    Sub New(ByVal _D As Grid2ЖурналClass)
        D = _D
        ' Этот вызов является обязательным для конструктора.
        InitializeComponent()

        ' Добавить код инициализации после вызова InitializeComponent().

    End Sub
    Private Sub ЖурналВсплывДанныеОбщиеДляГрид2_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        grid1All = New List(Of Grid1ЖурналClass)
        bsgrid1 = New BindingSource
        bsgrid1.DataSource = grid1All
        Grid1.DataSource = bsgrid1
        GridView(Grid1)
        Grid1.Columns(0).Width = 50
        Grid1.Columns(6).HeaderText = "Доп.инфо"
        Grid1.Columns(7).Visible = False
        Grid1.Columns(8).Visible = False
        Grid1.Columns(9).Visible = False
        Grid1.Columns(10).Visible = False
        Grid1.DefaultCellStyle.Font = New Font("Calibri", 9)

        Dim f1 = (From x In AllClass.ЖурналКлиентГруз
                  Join z In AllClass.ЖурналКлиентМаршрут On x.Код Equals z.КодЖурналКлиентГруз
                  Join y In AllClass.ЖурналПеревозчик On x.Код Equals y.КодЖурналКлиентГруз
                  Where x.Код = D.КодЖурналГруз
                  Select x, y, z).ToList
        Dim f2 = f1.Select(Function(x) x.y).ToList
        If f2 IsNot Nothing Then
            Dim i As Integer = 1
            For Each b In f2
                Dim f3 As New Grid1ЖурналClass With {.Дата = b.Дата, .ДопИнформация = b.ДопИнформация, .КодЖурналПеревозчик = b.Код, .КодПеревозчика = b.Кодперевозчик,
                    .Перевозчик = b.Организация, .Состояние = b.Состояние, .Ставка = b.Ставкапервозчика, .Номер = i, .Контакт = b.КонтДанные, .Skype = b?.Skype?.ToString, .SkypeDate = b?.SkypeDate?.ToString}
                grid1All.Add(f3)
                i += 1
            Next

            For Each b In grid1All
                If b.Skype?.Length > 0 Then
                    b.ДопИнформация = b.ДопИнформация & vbCrLf & "Skype +"
                End If
            Next

            bsgrid1.ResetBindings(False)
        End If



        Dim f4 As ЖурналКлиентГруз = Nothing
        Dim f5 As New List(Of ЖурналКлиентМаршрут)
        If f1 IsNot Nothing Then
            If f1.Count = 0 Then
                Dim f8 = (From x In AllClass.ЖурналКлиентГруз
                          Join z In AllClass.ЖурналКлиентМаршрут On x.Код Equals z.КодЖурналКлиентГруз
                          Where x.Код = D.КодЖурналГруз
                          Select x, z).ToList
                f4 = f8.Select(Function(x) x.x).FirstOrDefault()
                f5 = f8.Select(Function(x) x.z).ToList()
            Else

                f4 = f1.Select(Function(x) x.x).FirstOrDefault()
                f5 = f1.Select(Function(x) x.z).Distinct.ToList()
            End If
        End If
        'листбокс

        lst1 = New List(Of String)
        bslst1 = New BindingSource
        bslst1.DataSource = lst1
        ListBox1.DataSource = bslst1

        Dim mo As New AllUpd
        Do While AllClass.Клиент Is Nothing
            mo.КлиентAll()
        Loop
        Dim f7 = AllClass.Клиент.Where(Function(x) x.НазваниеОрганизации = f4?.Клиент).FirstOrDefault()

        ListBox1.BeginUpdate()
        'ListBox1.DrawMode = DrawMode.OwnerDrawFixed
        If f4 IsNot Nothing Then

            With lst1

                .Add("- - ДАННЫЕ ПО ГРУЗУ - -")
                .Add(" - Наименование груза: " & vbCrLf & f4.НаименованиеГруза)
                .Add(" - Размеры груза:")
                .Add(" Длина: " & f4.Длина)
                .Add(" Ширина: " & f4.Ширина)
                .Add(" Высота: " & f4.Высота)
                .Add(" - Дата загрузки: " & f4.ДатаЗагрузки)
                .Add(" - Дата выгрузки: " & f4.ДатаВыгрузки)
                .Add(" - Вес: " & f4.Вес & " кг.")
                .Add(" - Обем: " & f4.Обьем & " м3.")
                .Add(" - Кол.пал: " & f4.ПаллетыШтук & " шт.")
                .Add(" - Размер.пал: " & f4.РазмерПаллет & " м.")
                If f5 IsNot Nothing Then
                    .Add(" - СТАВКА: " & f5(0).Ставка)
                Else
                    .Add(" - СТАВКА: " & "Нет данных")
                End If

                .Add(vbCrLf & "___________________________________" & vbCrLf)
                .Add("- - ДАННЫЕ ПО МЕСТУ ЗАГРУЗКИ - -")
                If f5 IsNot Nothing Then
                    For Each b In f5
                        .Add(" - Место загрузки: " & "(" & b.СтранаПогрузки & ") " & b.ГородПогрузки & " (" & b.КвадратПогрузки & ")")
                        .Add(" - Место выгрузки: " & "(" & b.СтранаВыгрузки & ") " & b.ГородВыгрузки & " (" & b.КвадратВыгрузки & ")")
                        .Add(" - Там.отпр: " & vbCrLf & b.ТаможняОтправления)
                        .Add(" - Там.Назнач: " & vbCrLf & b.ТаможняНазначения)
                        .Add(" - Доп.информация: " & b.ДополнитИнформация)
                        .Add("/______________ . ________________/")
                    Next
                End If

                .Add("- - ДАННЫЕ ПО КЛИЕНТУ - -" & vbCrLf)
                .Add(" - РЕЗУЛЬТАТ: " & f4.РезультатРаботы & " - ДАТА: " & f4?.ДатаРезультата?.ToString)
                .Add(" - Конт.тел: " & f7?.Телефон?.ToString & ", " & " - Конт.лицо: " & f7?.Контактное_лицо?.ToString)

            End With

            bslst1.ResetBindings(False)

            ListBox1.EndUpdate()

        End If
        'lst2 = New BindingList(Of String)
        'bslst2 = New BindingSource
        'bslst2.DataSource = lst2
        'RichTextBox1.Text = bslst2.DataMember

    End Sub

    Private Sub Grid1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles Grid1.CellClick
        Dim f = grid1All.ElementAt(e.RowIndex)
        'If lst2 IsNot Nothing Then
        '    lst2.Clear()
        'End If
        'lst2.Add(f.Skype)
        RichTextBox1.Text = f.Skype
        Label5.Text = f.SkypeDate
    End Sub
End Class