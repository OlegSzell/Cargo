Option Explicit On
Imports System.ComponentModel
Imports System.Data.OleDb
Public Class Оплата
    Dim кто As String
    Dim NomerRejsa As Integer
    Dim f As Boolean = False
    Dim insertinbaza, Оплаты As String
    Dim ID, СтрокиВходящие As Integer
    Dim ostatok As Double = 0
    Dim lst As ОтчетДляОплаты
    Private grid1sel As List(Of Grid1Class)
    Private bsgrid1 As BindingSource
    Public ostt As String
    Private Flag As Boolean = False
    Sub New(ByVal _lst As ОтчетДляОплаты)
        lst = _lst
        ' Этот вызов является обязательным для конструктора.
        InitializeComponent()
        ПредзагрузкаAsync()
        ' Добавить код инициализации после вызова InitializeComponent().

    End Sub
    Private Async Sub ПредзагрузкаAsync()
        Await Task.Run(Sub() Предзагрузка())
    End Sub
    Private Sub Предзагрузка()
        Dim mo As New AllUpd
        Do While AllClass.РейсыПеревозчика Is Nothing
            mo.РейсыПеревозчикаAll()
        Loop
        Do While AllClass.РейсыКлиента Is Nothing
            mo.РейсыКлиентаAll()
        Loop
        Do While AllClass.ОплатыПер Is Nothing
            mo.ОплатыПерAll()
        Loop
        Do While AllClass.ОплатыКлиент Is Nothing
            mo.ОплатыКлиентAll()
        Loop
    End Sub

    Private Sub Оплата_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If lst Is Nothing Then Return
        Label2.Text = lst.НомерРейса
        кто = lst.Clickоплата
        Label20.Text = кто
        ID = lst.Код
        Label19.Text = lst.НазвОрганизации
        NomerRejsa = lst.НомерРейса
        Label23.Text = lst.СтоимостьФрахта
        Label5.Text = lst.Валюта


        grid1sel = New List(Of Grid1Class)
        bsgrid1 = New BindingSource
        bsgrid1.DataSource = grid1sel
        Grid1.DataSource = bsgrid1
        GridView(Grid1)
        'Grid1.SelectionMode = DataGridViewSelectionMode.CellSelect
        Grid1.Columns(0).Width = 60
        Grid1.Columns(0).HeaderText = "№"


        If кто = "Перевозчик" Then
            Историяперевоз()
            insertinbaza = "IDПер"
            Оплаты = "ОплатыПер"
        Else
            Историяклиент()
            insertinbaza = "IDКлиента"
            Оплаты = "ОплатыКлиент"
        End If
        Label24.Text = Math.Round(CDbl(lst.СтоимостьФрахта) - ostatok, 2)
        Label3.Text = ostatok

    End Sub
    Public Class Grid1Class
        Public Property Number As Integer
        Public Property Дата As Date
        Public Property Сумма As Double

    End Class
    Private Sub Историяперевоз()
        Dim mo As New AllUpd
        Do While AllClass.РейсыПеревозчика Is Nothing
            mo.РейсыПеревозчикаAll()
        Loop

        Do While AllClass.ОплатыПер Is Nothing
            mo.ОплатыПерAll()
        Loop

        Dim f = (From x In AllClass.РейсыПеревозчика
                 Join y In AllClass.ОплатыПер On x.Код Equals y.IDПер
                 Where x.НомерРейса = lst.НомерРейса
                 Select y).ToList()
        If grid1sel IsNot Nothing Then
            grid1sel.Clear()
        End If
        If f IsNot Nothing Then
            If f.Count > 0 Then
                Dim i As Integer = 1
                For Each b In f
                    Dim f2 As New Grid1Class With {.Number = i, .Дата = IIf(b.ДатаОплаты Is Nothing, Nothing, b.ДатаОплаты), .Сумма = Replace(b.Сумма, ".", ",")}
                    grid1sel.Add(f2)
                    ostatok += f2.Сумма
                    i += 1
                Next
                bsgrid1.ResetBindings(False)

            End If
        End If




    End Sub

    Private Sub Историяклиент()

        Dim mo As New AllUpd
        Do While AllClass.РейсыКлиента Is Nothing
            mo.РейсыКлиентаAll()
        Loop

        Do While AllClass.ОплатыКлиент Is Nothing
            mo.ОплатыКлиентAll()
        Loop

        Dim f = (From x In AllClass.РейсыКлиента
                 Join y In AllClass.ОплатыКлиент On x.Код Equals y.IDКлиента
                 Where x.НомерРейса = lst.НомерРейса
                 Select y).ToList()
        If grid1sel IsNot Nothing Then
            grid1sel.Clear()
        End If
        If f IsNot Nothing Then
            If f.Count > 0 Then
                Dim i As Integer = 1
                For Each b In f
                    Dim f2 As New Grid1Class With {.Number = i, .Дата = IIf(b.ДатаОплаты Is Nothing, Nothing, b.ДатаОплаты), .Сумма = Replace(b.Сумма, ".", ",")}
                    grid1sel.Add(f2)
                    ostatok += f2.Сумма
                    i += 1
                Next
                bsgrid1.ResetBindings(False)

            End If
        End If
    End Sub


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Dim opl = grid1sel.Select(Function(x) x.Сумма).Sum

        Label3.Text = CType(opl, String)

        Dim ost = CDbl(Replace(lst.СтоимостьФрахта, ".", ",")) - opl

        Label24.Text = CType(ost, String)

        ostt = CType(ost, String)

        Flag = True

        Dim mo As New AllUpd
        If lst.Clickоплата = "Перевозчик" Then
            'это данные для вставки в таблицу
            Dim f1 As New List(Of ОплатыПер)
            For Each b In grid1sel
                Dim f2 As New ОплатыПер With {.IDПер = lst.Код, .Рейс = lst.НомерРейса, .ДатаОплаты = b.Дата, .Сумма = b.Сумма}
                f1.Add(f2)
            Next

            'удаляем старые и вставляем новые
            Using db As New dbAllDataContext(_cn3)
                Dim f7 = (From x In db.ОплатыПер
                          Where x.Рейс = lst.НомерРейса And x.IDПер = lst.Код
                          Select x).ToList()

                If f7 IsNot Nothing Then
                    If f7.Count > 0 Then
                        db.ОплатыПер.DeleteAllOnSubmit(f7)
                        db.SubmitChanges()
                        db.ОплатыПер.InsertAllOnSubmit(f1)
                        db.SubmitChanges()

                    Else
                        If f1 IsNot Nothing Then
                            If f1.Count > 0 Then
                                db.ОплатыПер.InsertAllOnSubmit(f1)
                                db.SubmitChanges()

                            End If

                        End If
                    End If
                Else
                    If f1 IsNot Nothing Then
                        If f1.Count > 0 Then
                            db.ОплатыПер.InsertAllOnSubmit(f1)
                            db.SubmitChanges()

                        End If

                    End If
                End If

            End Using

            mo.ОплатыПерAllAsync()
            If ostt = "0" Then
                Using db As New dbAllDataContext(_cn3)
                    Dim f8 = db.РейсыПеревозчика.Where(Function(x) x.Код = lst.Код).Select(Function(x) x).FirstOrDefault()
                    If f8 IsNot Nothing Then
                        f8.ОстатокОплаты = 0
                        db.SubmitChanges()
                        mo.РейсыПеревозчикаAllAsync()
                    End If

                End Using
            End If

            Dim subj As String = "Рейс: " & Label2.Text & vbCrLf & "Фрахт: " & Label23.Text & vbCrLf & "Валюта: " & Label5.Text & vbCrLf _
            & "Оплачено: " & Label3.Text & vbCrLf & "Остаток: " & Label24.Text & vbCrLf & "Пользователь: " & Environment.MachineName
            Dim mail As New MySendMail(subj, "Оплаты перевозчику")
            mail.Mail()


            MessageBox.Show("Данные приняты!", Рик)

        Else 'клиент
            'это данные для вставки в таблицу
            Dim f1 As New List(Of ОплатыКлиент)
            For Each b In grid1sel
                Dim f2 As New ОплатыКлиент With {.IDКлиента = lst.Код, .Рейс = lst.НомерРейса, .ДатаОплаты = b.Дата, .Сумма = b.Сумма}
                f1.Add(f2)
            Next

            Using db As New dbAllDataContext(_cn3)
                Dim f = (From x In db.ОплатыКлиент
                         Where x.Рейс = lst.НомерРейса And x.IDКлиента = lst.Код
                         Select x).ToList()

                If f IsNot Nothing Then
                    If f.Count > 0 Then
                        db.ОплатыКлиент.DeleteAllOnSubmit(f)
                        db.SubmitChanges()
                        db.ОплатыКлиент.InsertAllOnSubmit(f1)
                        db.SubmitChanges()

                    Else
                        If f1 IsNot Nothing Then
                            If f1.Count > 0 Then
                                db.ОплатыКлиент.InsertAllOnSubmit(f1)
                                db.SubmitChanges()


                            End If

                        End If

                    End If
                Else
                    If f1 IsNot Nothing Then
                        If f1.Count > 0 Then
                            db.ОплатыКлиент.InsertAllOnSubmit(f1)
                            db.SubmitChanges()
                        End If
                    End If
                End If

            End Using
            mo.ОплатыКлиентAllAsync()
            If ostt = "0" Then
                Using db As New dbAllDataContext(_cn3)
                    Dim f8 = (From x In db.РейсыКлиента
                              Where x.Код = lst.Код
                              Select x).FirstOrDefault()
                    If f8 IsNot Nothing Then
                        f8.ОстатокОплаты = 0
                        db.SubmitChanges()
                        mo.РейсыКлиентаAllAsync()
                    End If

                End Using
            End If




            Dim subj As String = "Рейс: " & Label2.Text & vbCrLf & "Фрахт: " & Label23.Text & vbCrLf & "Валюта: " & Label5.Text & vbCrLf _
            & "Оплачено: " & Label3.Text & vbCrLf & "Остаток: " & Label24.Text & vbCrLf & "Пользователь: " & Environment.MachineName
            Dim mail As New MySendMail(subj, "Оплаты от клиента")
            mail.Mail()


            MessageBox.Show("Данные приняты!", Рик)
        End If




    End Sub

    Private Sub Обновление()
        Dim dbl As Double = Nothing
        Dim stsql As String = "ThenSELECT * FROM " & Оплаты & " WHERE " & insertinbaza & "=" & ID & " And Рейс=" & NomerRejsa & ""
        Dim ds As DataTable = Selects3(stsql)
        For i As Integer = 0 To ds.Rows.Count - 1
            If ds.Rows(i).Item(4).ToString.Contains(".") Then
                ds.Rows(i).Item(4) = Replace(ds.Rows(i).Item(4).ToString, ".", ",")
            End If
            dbl += CDbl(ds.Rows(i).Item(4).ToString)
        Next
        'Label24.Text = Math.Round(CDbl(Отчет.datatbl.Rows(0).Item(2).ToString) - dbl, 2) & " " & Отчет.datatbl.Rows(0).Item(3).ToString
        'Dim strsql1 As String

        'If кто = "Перевозчик" Then
        '    strsql1 = "UPDATE РейсыПеревозчика SET ОстатокОплаты='" & CType(Math.Round(CDbl(Отчет.datatbl.Rows(0).Item(2).ToString) - dbl, 2), String) & "' WHERE Код=" & Отчет.datatbl.Rows(0).Item(0) & ""

        'Else
        '    strsql1 = "UPDATE РейсыКлиента SET ОстатокОплаты='" & CType(Math.Round(CDbl(Отчет.datatbl.Rows(0).Item(2).ToString) - dbl, 2), String) & "' WHERE Код=" & Отчет.datatbl.Rows(0).Item(0) & ""
        'End If
        'Updates3(strsql1)
        'Отчет.ОсткиОплат()

    End Sub

    Private Sub Оплата_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        If Flag = False Then
            If Label3.Text = "0" Then
                ostt = "Рейс не оплачен!"

            Else
                If CType(Label23.Text, Integer) = CType(Label3.Text, Integer) Then
                    ostt = "Рейс оплачен!"
                Else
                    ostt = Label24.Text
                End If

            End If
        End If
    End Sub

End Class