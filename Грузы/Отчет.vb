Option Explicit On
Imports System.Data.OleDb
Imports System.IO

Public Class Отчет
    Public datatbl As DataTable
    Public Clickоплата As String = ""

    Public назворг As String = ""
    Dim ds3, ds4 As DataTable
    Public ID As Integer = 0
    Private Nom As Integer
    Private grid1all As List(Of Grid1Class)
    Private bsgrid1 As BindingSource

    Sub New(Optional _nom As Integer = 0)
        Nom = _nom
        ' Этот вызов является обязательным для конструктора.
        InitializeComponent()

        ' Добавить код инициализации после вызова InitializeComponent().

    End Sub
    Public Class Grid1Class
        Inherits Grid2Class
        Public Property Организац As String
        Public Property Рейс As String
        Public Property Маршрут As String
        Public Property Груз As String
        Public Property Фрахт As String
        Public Property СрокОплаты As String
        Public Property УсловияОплаты As String
        Public Property ДатаОтправкиДоков As String
        Public Property ДатаОплаты As String
    End Class
    Public Class Grid2Class
        Public Property ДатаПолученияДоков As String

    End Class
    Private Sub Отчет_FormClosed(sender As Object, e As FormClosedEventArgs) Handles MyBase.FormClosed
        MaskedTextBox1.Text = ""
        MaskedTextBox2.Text = ""

        НомРейса = Nothing
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        MaskedTextBox1.Text = ""
        MaskedTextBox2.Text = ""

        НомРейса = Nothing
        Me.Close()
    End Sub

    Private Sub Отчет_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If bl = True Then 'переменная для определия откуда запущена форма
            Label3.Text = НомРейса
        Else
            НомРейса = Nom
            Label3.Text = Nom
        End If

        If Not ds3 Is Nothing Then
            ds3.Clear()
        End If
        If Not ds4 Is Nothing Then
            ds4.Clear()
        End If

        ОсткиОплат()

        UpdateGrid()

    End Sub
    Public Sub ОсткиОплат()
        'рейс клиента
        Dim mo As New AllUpd
        Do While AllClass.РейсыКлиента Is Nothing
            mo.РейсыКлиентаAll()
        Loop
        Do While AllClass.РейсыПеревозчика Is Nothing
            mo.РейсыПеревозчикаAll()
        Loop

        Do While AllClass.ОплатыКлиент Is Nothing
            mo.ОплатыКлиентAll()
        Loop
        Do While AllClass.ОплатыПер Is Nothing
            mo.ОплатыПерAll()
        Loop

        Dim f = (From x In AllClass.РейсыКлиента
                 Join y In AllClass.ОплатыКлиент On x.Код Equals y.IDКлиента
                 Where x.НомерРейса = НомРейса
                 Select x, y).ToList()
        If f IsNot Nothing Then
            If f.Count > 0 Then


                Dim f4 = (From x In f
                          Group x By Keys = New With {Key x.x.СтоимостьФрахта}
                             Into Group
                          Select New With {Keys.СтоимостьФрахта, .Оплачено = (From b In Group
                                                                              Select b.y.Сумма).ToList()}).FirstOrDefault()


                If f4.Оплачено Is Nothing Then
                    Label9.Text = "Рейс не оплачен!"
                    Label9.ForeColor = Color.Red
                Else
                    Dim opl = CDbl(f4.СтоимостьФрахта) - f4.Оплачено.Select(Function(x) CDbl(x)).Sum()
                    If opl = 0 Then
                        Label9.Text = "Рейс оплачен!"
                        Label9.ForeColor = Color.Green
                    Else
                        If opl = CDbl(f4.СтоимостьФрахта) Then
                            Label9.Text = "Рейс не оплачен!"
                            Label9.ForeColor = Color.Red
                        Else
                            Label9.Text = opl
                            Label9.ForeColor = Color.Green
                        End If

                    End If

                End If
            Else
                Label9.Text = "Рейс не оплачен!"
                Label9.ForeColor = Color.Red
            End If
        Else
            Label9.Text = "Рейс не оплачен!"
            Label9.ForeColor = Color.Red
        End If



        'рейс перевозчика


        Dim f5 = (From x In AllClass.РейсыПеревозчика
                  Join y In AllClass.ОплатыПер On x.Код Equals y.IDПер
                  Where x.НомерРейса = НомРейса
                  Select x, y).ToList()
        If f5 IsNot Nothing Then
            If f5.Count > 0 Then
                Dim f7 = (From x In f5
                          Group x By Keys = New With {Key x.x.СтоимостьФрахта}
                             Into Group
                          Select New With {Keys.СтоимостьФрахта, .Оплачено = Group.Select(Function(x) x.y.Сумма).ToList()}).FirstOrDefault()


                If f7.Оплачено Is Nothing Then
                    Label8.Text = "Рейс не оплачен!"
                    Label8.ForeColor = Color.Red
                Else
                    Dim opl = CDbl(f7.СтоимостьФрахта) - f7.Оплачено.Select(Function(x) CDbl(x)).Sum()


                    If opl = 0 Then
                        Label8.Text = "Рейс оплачен!"
                        Label8.ForeColor = Color.Green
                    Else
                        If opl = CDbl(f7.СтоимостьФрахта) Then
                            Label8.Text = "Рейс не оплачен!"
                            Label8.ForeColor = Color.Red
                        Else
                            Label8.Text = opl
                            Label8.ForeColor = Color.Green
                        End If
                    End If
                End If
            Else
                Label8.Text = "Рейс не оплачен!"
                Label8.ForeColor = Color.Red

            End If

        Else
            Label8.Text = "Рейс не оплачен!"
            Label8.ForeColor = Color.Red
        End If

















        'Dim f1 = AllClass.РейсыПеревозчика.Where(Function(x) x.НомерРейса = НомРейса).Select(Function(x) x).FirstOrDefault()

        'If f1 IsNot Nothing Then
        '    If f1.ОстатокОплаты Is Nothing Then
        '        Label8.Text = "Рейс не оплачен!"
        '        Label8.ForeColor = Color.Red
        '    ElseIf f1.ОстатокОплаты = 0 Then
        '        Label8.Text = "Рейс оплачен!"
        '        Label8.ForeColor = Color.Green
        '    End If
        'End If

    End Sub
    Private Sub UpdateGrid()
        grid1all = New List(Of Grid1Class)
        bsgrid1 = New BindingSource
        bsgrid1.DataSource = grid1all
        Grid1.DataSource = bsgrid1
        GridView(Grid1)
        Grid1.Columns(5).HeaderText = "Срок оплаты"
        Grid1.Columns(6).HeaderText = "Условия оплаты"
        Grid1.Columns(7).HeaderText = "Дата отп.док клиенту"
        Grid1.Columns(8).HeaderText = "Дата оплаты"
        Grid1.Columns(9).HeaderText = "Дата получ.док от перевоз"




        Dim mo As New AllUpd
        Do While AllClass.РейсыКлиента Is Nothing
            mo.РейсыКлиентаAll()
        Loop
        Do While AllClass.РейсыПеревозчика Is Nothing
            mo.РейсыПеревозчикаAll()
        Loop
        If grid1all IsNot Nothing Then
            grid1all.Clear()
        End If


        Dim f = (From x In AllClass.РейсыКлиента
                 Where x.НомерРейса = НомРейса
                 Select x).FirstOrDefault()
        If f IsNot Nothing Then
            Dim k = Format(f.ДатаОтправкиДоков, "dd.MM.yyyy")
            If k.Length > 0 Then
                MaskedTextBox2.Text = k
                Dim m As New Grid1Class With {.Груз = IIf(f.НаименованиеГруза.Length > 50, Strings.Left(f.НаименованиеГруза, 50), f.НаименованиеГруза), .ДатаОплаты = f.ДатаОплаты, .ДатаОтправкиДоков = f.ДатаОтправкиДоков,
                    .Маршрут = f.Маршрут, .Организац = f.НазвОрганизации, .Рейс = f.НомерРейса, .СрокОплаты = f.СрокОплаты, .УсловияОплаты = f.УсловияОплаты, .Фрахт = f.СтоимостьФрахта & " " & f.Валюта}
                grid1all.Add(m)

            End If
            Label10.Text = f.НазвОрганизации
        End If

        Dim f2 = (From x In AllClass.РейсыПеревозчика
                  Where x.НомерРейса = НомРейса
                  Select x).FirstOrDefault()
        If f2 IsNot Nothing Then
            Dim k = Format(f2.ДатаПолученияДоков, "dd.MM.yyyy")
            If k.Length > 0 Then
                MaskedTextBox1.Text = k
                Dim m2 As New Grid1Class With {.Груз = IIf(f2.НаименованиеГруза.Length > 50, Strings.Left(f2.НаименованиеГруза, 50), f2.НаименованиеГруза), .ДатаОплаты = f2.ДатаОплаты, .ДатаПолученияДоков = f2.ДатаПолученияДоков,
                    .Маршрут = f2.Маршрут, .Организац = f2.НазвОрганизации, .Рейс = f2.НомерРейса, .СрокОплаты = f2.СрокОплаты, .УсловияОплаты = f2.УсловияОплаты, .Фрахт = f2.СтоимостьФрахта & " " & f2.Валюта}
                grid1all.Add(m2)

            End If
            Label11.Text = f2.НазвОрганизации
        End If

        Dim m3 As New Grid1Class With {.Организац = "ИТОГО", .Фрахт = Math.Round(f.СтоимостьФрахта - f2.СтоимостьФрахта, 2)}
        grid1all.Add(m3)
        bsgrid1.ResetBindings(False)

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        СортПеревоз()
        СортКлиент()
        UpdateGrid()
    End Sub
    Private Sub СортКлиент()

        Dim mo As New AllUpd
        Do While AllClass.РейсыКлиента Is Nothing
            mo.РейсыКлиентаAll()
        Loop


        If MaskedTextBox2.MaskCompleted = True Then

            Dim f = (From x In AllClass.РейсыКлиента
                     Where x.НомерРейса = НомРейса
                     Select x).FirstOrDefault()
            If f Is Nothing Then Return

            Dim str As String = Trim(f.СрокОплаты)
            Dim d As Integer

            If str.Length > 2 Then 'отбираем количество дней оплаты
                If str.Length > 5 Then 'если цифр больше 5
                    MessageBox.Show("Число цифр больше 5, сообщите администратору!")
                    Return
                End If
                If str.Length = 3 Then
                    d = CType(Strings.Right(str, 1), Integer)
                Else
                    d = CType(Strings.Right(str, 2), Integer)
                End If
            Else
                d = CType(str, Integer)
            End If
            If f.УсловияОплаты = "БанкПоДок" Or f.УсловияОплаты = "БанкПоВыг" Then
                Dim dt = AddWorkDays(MaskedTextBox2.Text, d)
                Using db As New dbAllDataContext()
                    Dim f4 = db.РейсыКлиента.Where(Function(x) x.НомерРейса = НомРейса).Select(Function(x) x).FirstOrDefault()
                    If f4 IsNot Nothing Then
                        With f4
                            .ДатаОплаты = dt
                            .ДатаОтправкиДоков = MaskedTextBox2.Text
                            db.SubmitChanges()
                            mo.РейсыКлиентаAllAsync()
                        End With
                    End If
                End Using
            Else
                Dim ml = CDate(MaskedTextBox2.Text).AddDays(d)
                Using db As New dbAllDataContext()
                    Dim f4 = db.РейсыКлиента.Where(Function(x) x.НомерРейса = НомРейса).Select(Function(x) x).FirstOrDefault()
                    If f4 IsNot Nothing Then
                        With f4
                            .ДатаОплаты = ml
                            .ДатаОтправкиДоков = MaskedTextBox2.Text
                            db.SubmitChanges()
                            mo.РейсыКлиентаAllAsync()
                        End With
                    End If
                End Using
            End If




        End If



        'Dim strsql, strsql1, str As String
        'Dim ds As DataTable
        'Dim d As Integer
        'Dim df As Double

        'If MaskedTextBox2.MaskCompleted = True Then
        '    strsql = "SELECT СрокОплаты, УсловияОплаты FROM РейсыКлиента WHERE НомерРейса=" & НомРейса & ""
        '    ds = Selects3(strsql)
        '    str = ds.Rows(0).Item(0).ToString
        '    Dim dat As Date = MaskedTextBox2.Text

        '    If str.Length > 2 Then 'отбираем количество дней оплаты
        '        If str.Length > 5 Then 'если цифр больше 5
        '            MessageBox.Show("Число цифр больше 5, сообщите администратору!")
        '            Exit Sub
        '        End If

        '        If str.Length = 3 Then
        '            d = CType(Strings.Right(str, 1), Integer)
        '        Else
        '            d = CType(Strings.Right(str, 2), Integer)
        '        End If
        '    Else
        '        d = CType(str, Integer)
        '    End If

        '    If ds.Rows(0).Item(1).ToString = "БанкПоДок" Or ds.Rows(0).Item(1).ToString = "БанкПоВыг" Then

        '        If d <= 5 Then 'отсортировываем календарные и рабочие дни
        '            d += 2
        '        Else
        '            df = Math.Round(d / 5)
        '            d += (df * 2)
        '        End If
        '    End If

        '    dat = dat.AddDays(d)
        '    strsql1 = "UPDATE РейсыКлиента SET ДатаОплаты='" & dat & "', ДатаОтправкиДоков='" & MaskedTextBox2.Text & "' WHERE НомерРейса=" & НомРейса & ""
        '    Updates3(strsql1)
        'End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim f = Перевоз()
        If f = "0" Then
            Label8.Text = "Рейс оплачен!"
        Else
            Label8.Text = f

        End If

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim f = Кли()
        If f = "0" Then
            Label9.Text = "Рейс оплачен!"
        Else
            Label9.Text = f

        End If

    End Sub
    Public Function Перевоз() As String
        Dim mo As New AllUpd
        Do While AllClass.РейсыПеревозчика Is Nothing
            mo.РейсыПеревозчикаAll()
        Loop
        Dim nomrs As String = Label3.Text
        Dim f = (From x In AllClass.РейсыПеревозчика
                 Where x.НомерРейса = НомРейса
                 Select New ОтчетДляОплаты With {.Clickоплата = "Перевозчик", .Валюта = x.Валюта, .Код = x.Код, .СтоимостьФрахта = x.СтоимостьФрахта,
                                                                                                .НазвОрганизации = x.НазвОрганизации, .НомерРейса = nomrs}).FirstOrDefault()
        If f IsNot Nothing Then
            Dim f2 As New Оплата(f)
            f2.ShowDialog()
            Return f2.ostt
        Else
            Return "Рейс не найден!"
        End If

    End Function
    Public Function Кли() As String
        Dim mo As New AllUpd
        Do While AllClass.РейсыКлиента Is Nothing
            mo.РейсыКлиентаAll()
        Loop
        Dim nomrs As String = Label3.Text
        Dim numb As Integer = НомРейса
        'Dim f3 = AllClass.РейсыКлиента.Where(Function(x) x.НомерРейса = numb).Select(Function(x) x).FirstOrDefault()
        Dim f3 = (From x In AllClass.РейсыКлиента
                  Where x.НомерРейса = numb
                  Select x).FirstOrDefault()

        Dim f As New ОтчетДляОплаты With {.Clickоплата = "Клиент", .Валюта = f3.Валюта, .Код = f3.Код, .СтоимостьФрахта = f3.СтоимостьФрахта,
                                                                                                .НазвОрганизации = f3.НазвОрганизации, .НомерРейса = nomrs}
        'Dim f = AllClass.РейсыКлиента.Where(Function(x) x.НомерРейса = НомРейса).Select(Function(x) New ОтчетДляОплаты With {.Clickоплата = "Клиент",
        '                                                                                        .Валюта = x.Валюта, .Код = x.Код, .СтоимостьФрахта = x.СтоимостьФрахта,
        '                                                                                        .НазвОрганизации = x.НазвОрганизации, .НомерРейса = nomrs}).FirstOrDefault()
        If f IsNot Nothing Then
            Dim f2 As New Оплата(f)
            f2.ShowDialog()
            Return f2.ostt
        Else
            Return "Рейс не найден!"
        End If



        'НомерРейса3 = Label3.Text
        'Clickоплата = "Клиент"
        'Dim strsql As String = "SELECT Код,НазвОрганизации,СтоимостьФрахта,Валюта FROM РейсыКлиента WHERE НомерРейса=" & НомРейса & ""
        'Dim ds As DataTable = Selects3(strsql)
        'ID = ds.Rows(0).Item(0)
        'назворг = ds.Rows(0).Item(1).ToString
        'datatbl = ds
        'Dim f As New Оплата
        'f.ShowDialog()
    End Function
    Private Sub СортПеревоз()


        Dim mo As New AllUpd
        Do While AllClass.РейсыПеревозчика Is Nothing
            mo.РейсыПеревозчикаAll()
        Loop


        If MaskedTextBox1.MaskCompleted = True Then

            Dim f = (From x In AllClass.РейсыПеревозчика
                     Where x.НомерРейса = НомРейса
                     Select x).FirstOrDefault()
            If f Is Nothing Then Return

            Dim str As String = Trim(f.СрокОплаты)
            Dim d As Integer

            If str.Length > 2 Then 'отбираем количество дней оплаты
                If str.Length > 5 Then 'если цифр больше 5
                    MessageBox.Show("Число цифр больше 5, сообщите администратору!")
                    Return
                End If
                If str.Length = 3 Then
                    d = CType(Strings.Right(str, 1), Integer)
                Else
                    d = CType(Strings.Right(str, 2), Integer)
                End If
            Else
                d = CType(str, Integer)
            End If
            If f.УсловияОплаты = "БанкПоДок" Or f.УсловияОплаты = "БанкПоВыг" Then
                Dim dt = AddWorkDays(MaskedTextBox1.Text, d)
                Using db As New dbAllDataContext()
                    Dim f4 = db.РейсыПеревозчика.Where(Function(x) x.НомерРейса = НомРейса).Select(Function(x) x).FirstOrDefault()
                    If f4 IsNot Nothing Then
                        With f4
                            .ДатаОплаты = dt
                            .ДатаПолученияДоков = MaskedTextBox1.Text
                            db.SubmitChanges()
                            mo.РейсыПеревозчикаAllAsync()
                        End With
                    End If
                End Using
            Else
                Dim ml = CDate(MaskedTextBox1.Text).AddDays(d)
                Using db As New dbAllDataContext()
                    Dim f4 = db.РейсыПеревозчика.Where(Function(x) x.НомерРейса = НомРейса).Select(Function(x) x).FirstOrDefault()
                    If f4 IsNot Nothing Then
                        With f4
                            .ДатаОплаты = ml
                            .ДатаПолученияДоков = MaskedTextBox1.Text
                            db.SubmitChanges()
                            mo.РейсыПеревозчикаAllAsync()
                        End With
                    End If
                End Using
            End If



            'strsql = "SELECT СрокОплаты, УсловияОплаты FROM РейсыПеревозчика WHERE НомерРейса=" & НомРейса & ""
            'ds = Selects3(strsql)
            'str = ds.Rows(0).Item(0).ToString
            'Dim dat As Date = MaskedTextBox1.Text

            'If str.Length > 2 Then 'отбираем количество дней оплаты
            '    If str.Length > 5 Then 'если цифр больше 5
            '        MessageBox.Show("Число цифр больше 5, сообщите администратору!")
            '        Exit Sub
            '    End If

            '    If str.Length = 3 Then
            '        d = CType(Strings.Right(str, 1), Integer)
            '    Else
            '        d = CType(Strings.Right(str, 2), Integer)
            '    End If
            'Else
            '    d = CType(str, Integer)
            'End If

            'If ds.Rows(0).Item(1).ToString = "БанкПоДок" Or ds.Rows(0).Item(1).ToString = "БанкПоВыг" Then

            '    If d <= 5 Then 'отсортировываем календарные и рабочие дни
            '        d += 2
            '    Else
            '        df = Math.Round(d / 5)
            '        d += (df * 2)
            '    End If
            'End If

            'dat = dat.AddDays(d)
            'strsql1 = "UPDATE РейсыПеревозчика SET ДатаОплаты='" & dat & "', ДатаПолученияДоков='" & MaskedTextBox1.Text & "' WHERE НомерРейса=" & НомРейса & ""
            'Updates3(strsql1)
        End If
    End Sub
    Private Shared Function AddWorkDays(ByVal date1 As DateTime, ByVal n As Integer) As DateTime
        Dim daysBeforeWeekEnd As Integer = 5 - CInt(date1.DayOfWeek)
        If n <= daysBeforeWeekEnd Then Return date1.AddDays(n)
        Dim fullWorkWeeks As Integer = n / 5
        Dim daysAfterWeekEnd As Integer = n - daysBeforeWeekEnd - fullWorkWeeks * 5
        Return date1.AddDays(daysBeforeWeekEnd + fullWorkWeeks * 7 + daysAfterWeekEnd + (If(daysAfterWeekEnd = 0, 0, 2)))
    End Function


End Class
Public Class ОтчетДляОплаты
    Public Property Код As Integer
    Public Property НазвОрганизации As String
    Public Property СтоимостьФрахта As String
    Public Property Валюта As String
    Public Property Clickоплата As String
    Public Property НомерРейса As String
End Class