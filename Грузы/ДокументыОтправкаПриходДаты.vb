Imports System.ComponentModel

Public Class ДокументыОтправкаПриходДаты
    Private grid1all As New BindingList(Of Grid1Class)
    Private bsgrid1 As New BindingSource
    Private НомРейса As Integer
    Dim БанкПоДок As String = "БанкПоДок"
    Dim КалПоДок As String = "КалПоДок"
    Dim КалПоВыг As String = "КалПоВыг"

    Sub New(ByVal _НомРейса As Integer)
        НомРейса = _НомРейса
        AllDannyeAsync()
        ' Этот вызов является обязательным для конструктора.
        InitializeComponent()
        Label5.Text = НомРейса
        ' Добавить код инициализации после вызова InitializeComponent().

    End Sub
    Private Sub ДокументыОтправкаПриходДаты_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        bsgrid1.DataSource = grid1all
        Grid1.DataSource = bsgrid1
        GridView(Grid1)
        Grid1.SelectionMode = DataGridViewSelectionMode.CellSelect
        Grid1.Columns(0).Width = 60
        Grid1.Columns(1).HeaderText = "Клиент или Перевозчик"
        Grid1.Columns(3).HeaderText = "Дата документов"
        Grid1.Columns(4).HeaderText = "Срок оплаты"
        Grid1.Columns(5).HeaderText = "Дата оплаты"
        Grid1.Columns(4).Width = 60
        Grid1.Columns(0).ReadOnly = True
        Grid1.Columns(1).ReadOnly = True
        Grid1.Columns(2).ReadOnly = True
        Grid1.Columns(4).ReadOnly = True
        Grid1.Columns(6).Visible = False

        Dim mo As New AllUpd
        Do While AllClass.РейсыКлиента Is Nothing
            mo.РейсыКлиентаAll()
        Loop
        Do While AllClass.РейсыПеревозчика Is Nothing
            mo.РейсыПеревозчикаAll()
        Loop

        Dim f = (From x In AllClass.РейсыКлиента
                 Where x.НомерРейса = НомРейса
                 Select New Grid1Class With {.Номер = 1, .КлиентИлиПеревозчик = "Клиент",
                     .ДатаДокументов = IIf(x.ДатаОтправкиДоков Is Nothing, Nothing, x.ДатаОтправкиДоков),
                     .ДатаОплаты = IIf(x.ДатаОплаты Is Nothing, Nothing, x.ДатаОплаты), .Организация = x.НазвОрганизации, .СрокОплаты = x.СрокОплаты, .КакиеДниОплаты = x.УсловияОплаты}).FirstOrDefault()

        Dim f1 = (From x In AllClass.РейсыПеревозчика
                  Where x.НомерРейса = НомРейса
                  Select New Grid1Class With {.Номер = 2, .КлиентИлиПеревозчик = "Перевозчик",
                      .ДатаДокументов = IIf(x.ДатаПолученияДоков Is Nothing, Nothing, x.ДатаПолученияДоков),
                      .ДатаОплаты = IIf(x.ДатаОплаты Is Nothing, Nothing, x.ДатаОплаты), .Организация = x.НазвОрганизации, .СрокОплаты = x.СрокОплаты, .КакиеДниОплаты = x.УсловияОплаты}).FirstOrDefault()


        If grid1all IsNot Nothing Then
            grid1all.Clear()
        End If
        grid1all.Add(f)
        grid1all.Add(f1)


    End Sub
    Public Class Grid1Class
        Public Property Номер As Integer
        Public Property КлиентИлиПеревозчик As String
        Public Property Организация As String
        Public Property ДатаДокументов As Date
        Public Property СрокОплаты As String
        Public Property ДатаОплаты As Date
        Public Property КакиеДниОплаты As String

    End Class
    Private Async Sub AllDannyeAsync()
        Await Task.Run(Sub() AllDannye())
    End Sub
    Private Sub AllDannye()
        Dim mo As New AllUpd
        mo.РейсыКлиентаAll()
        mo.РейсыПеревозчикаAll()

    End Sub
    Private Sub UpdClient(ByVal dtstar As Date, ByVal dtOplata As Date)


    End Sub

    Private Sub Grid1_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles Grid1.CellEndEdit
        If e.ColumnIndex = 3 Then
            Dim mo As New AllUpd

            If e.RowIndex = 0 Then 'клиент
                Dim m As Grid1Class = grid1all.ElementAt(e.RowIndex)

                If m.ДатаДокументов = Nothing Then
                    grid1all(0).ДатаДокументов = Nothing
                    grid1all(0).ДатаОплаты = Nothing
                    bsgrid1.ResetBindings(False)
                    Deletes("Zak")
                    Return
                End If


                Dim m1 As String
                Dim EndDay As Date = Nothing
                If m.СрокОплаты.Length <= 2 Then
                    m1 = m.СрокОплаты
                Else
                    If m.СрокОплаты.Length = 3 Then
                        m1 = Strings.Right(m.СрокОплаты, 1)
                    ElseIf m.СрокОплаты.Length >= 4 And m.СрокОплаты.Length <= 5 Then
                        m1 = Strings.Right(m.СрокОплаты, 2)
                    Else
                        m1 = Strings.Right(m.СрокОплаты, 3)

                    End If

                End If

                Select Case m.КакиеДниОплаты
                    Case БанкПоДок
                        EndDay = m.ДатаДокументов.AddDays(m1 + 2 + (Math.Round(m1 / 5) * 2))
                    Case КалПоДок
                        EndDay = m.ДатаДокументов.AddDays(Math.Round(m1 + 2))
                    Case КалПоВыг
                        EndDay = m.ДатаДокументов.AddDays(m1)
                End Select

                grid1all(0).ДатаОплаты = EndDay
                bsgrid1.ResetBindings(False)
                UpdDokiDateAsync(m.ДатаДокументов, EndDay)
            Else 'перевозчик


                Dim m As Grid1Class = grid1all.ElementAt(e.RowIndex)
                If m.ДатаДокументов = Nothing Then
                    grid1all(1).ДатаДокументов = Nothing
                    grid1all(1).ДатаОплаты = Nothing
                    bsgrid1.ResetBindings(False)
                    Deletes("Per")
                    Return
                End If
                Dim m1 As String
                Dim EndDay As Date = Nothing
                If m.СрокОплаты.Length <= 2 Then
                    m1 = m.СрокОплаты
                Else
                    If m.СрокОплаты.Length = 3 Then
                        m1 = Strings.Right(m.СрокОплаты, 1)
                    ElseIf m.СрокОплаты.Length >= 4 And m.СрокОплаты.Length <= 5 Then
                        m1 = Strings.Right(m.СрокОплаты, 2)
                    Else
                        m1 = Strings.Right(m.СрокОплаты, 3)

                    End If

                End If

                Select Case m.КакиеДниОплаты
                    Case БанкПоДок
                        EndDay = m.ДатаДокументов.AddDays(m1 + (Math.Round(m1 / 5) * 2))
                    Case КалПоДок
                        EndDay = m.ДатаДокументов.AddDays(m1)
                    Case КалПоВыг
                        EndDay = m.ДатаДокументов.AddDays(m1)
                End Select

                grid1all(1).ДатаОплаты = EndDay
                bsgrid1.ResetBindings(False)
                UpdDokiDateAsyncPer(m.ДатаДокументов, EndDay)
            End If
        End If
    End Sub
    Private Sub Deletes(ByVal Who As String)
        If Who = "Per" Then
            Using db As New dbAllDataContext(_cn3)
                Dim f = db.РейсыПеревозчика.Where(Function(x) x.НомерРейса = НомРейса).FirstOrDefault()
                If f IsNot Nothing Then
                    f.ДатаПолученияДоков = Nothing
                    f.ДатаОплаты = Nothing
                    db.SubmitChanges()
                    Dim mo As New AllUpd
                    mo.РейсыПеревозчикаAll()
                End If
            End Using
        Else
            Using db As New dbAllDataContext(_cn3)
                Dim f = db.РейсыКлиента.Where(Function(x) x.НомерРейса = НомРейса).FirstOrDefault()
                If f IsNot Nothing Then
                    f.ДатаОтправкиДоков = Nothing
                    f.ДатаОплаты = Nothing
                    db.SubmitChanges()
                    Dim mo As New AllUpd
                    mo.РейсыКлиентаAll()
                End If
            End Using

        End If

    End Sub
    Private Async Sub UpdDokiDateAsync(ByVal ОТправка As Date, Oplata As Date)
        Await Task.Run(Sub() UpdDokiDate(ОТправка, Oplata))
    End Sub
    Private Sub UpdDokiDate(ByVal ОТправка As Date, Oplata As Date)
        Using db As New dbAllDataContext(_cn3)
            Dim f = db.РейсыКлиента.Where(Function(x) x.НомерРейса = НомРейса).FirstOrDefault()
            If f IsNot Nothing Then
                f.ДатаОтправкиДоков = ОТправка
                f.ДатаОплаты = Oplata
                db.SubmitChanges()
                Dim mo As New AllUpd
                mo.РейсыКлиентаAll()
            End If
        End Using

    End Sub

    Private Async Sub UpdDokiDateAsyncPer(ByVal ОТправка As Date, Oplata As Date)
        Await Task.Run(Sub() UpdDokiDatePer(ОТправка, Oplata))
    End Sub
    Private Sub UpdDokiDatePer(ByVal ОТправка As Date, Oplata As Date)
        Using db As New dbAllDataContext(_cn3)
            Dim f = db.РейсыПеревозчика.Where(Function(x) x.НомерРейса = НомРейса).FirstOrDefault()
            If f IsNot Nothing Then
                f.ДатаПолученияДоков = ОТправка
                f.ДатаОплаты = Oplata
                db.SubmitChanges()
                Dim mo As New AllUpd
                mo.РейсыПеревозчикаAll()
            End If
        End Using

    End Sub

    Private Sub Grid1_DataError(sender As Object, e As DataGridViewDataErrorEventArgs) Handles Grid1.DataError


        'MessageBox.Show("Error happened " & e.Context.ToString())

        If (e.Context = DataGridViewDataErrorContexts.Commit) Then
            grid1all(e.RowIndex).ДатаДокументов = Nothing
            'MessageBox.Show("Commit error")
        End If

        'If (e.Context = DataGridViewDataErrorContexts _
        '.CurrentCellChange) Then
        '    MessageBox.Show("Cell change")
        'End If
        'If (e.Context = DataGridViewDataErrorContexts.Parsing) _
        'Then
        '    MessageBox.Show("parsing error")
        'End If
        'If (e.Context =
        'DataGridViewDataErrorContexts.LeaveControl) Then
        '    MessageBox.Show("leave control error")
        'End If


        If (TypeOf (e.Exception) Is ConstraintException) Then
            Dim view As DataGridView = CType(sender, DataGridView)
            view.Rows(e.RowIndex).ErrorText = "an error"
            view.Rows(e.RowIndex).Cells(e.ColumnIndex) _
                .ErrorText = "an error"

            e.ThrowException = False
        End If
    End Sub
End Class