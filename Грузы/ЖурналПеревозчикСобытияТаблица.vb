Public Class ЖурналПеревозчикСобытияТаблица
    Private Grid1all As New List(Of Grid1Class)
    Private bsGrid1 As New BindingSource
    Private ReadOnly ID As Integer
    Public Rez As String = Nothing
    Sub New(ByVal _Id As Integer)
        ID = _Id

        ' Этот вызов является обязательным для конструктора.
        InitializeComponent()
        predAsync()
        ' Добавить код инициализации после вызова InitializeComponent().

    End Sub
    Private Sub ЖурналПеревозчикСобытияТаблица_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        bsGrid1.DataSource = Grid1all
        Grid1.DataSource = bsGrid1
        GridView(Grid1)
        Grid1.SelectionMode = DataGridViewSelectionMode.CellSelect
        Grid1.Columns(0).Width = 60
        Grid1.Columns(1).Width = 80
        Grid1.Columns(2).Width = 120

        Dim mo As New AllUpd
        Do While AllClass.ЖурналПеревозчик Is Nothing
            mo.ЖурналПеревозчикAll()
        Loop
        Do While AllClass.ЖурналПеревозчикСобытия Is Nothing
            mo.ЖурналПеревозчикСобытияAll()
        Loop
        Dim i As Integer = 1
        Dim f = (From x In AllClass.ЖурналПеревозчик
                 Join y In AllClass.ЖурналПеревозчикСобытия On x.Код Equals y.IDЖурналПеревозчик
                 Where x.Код = ID
                 Select New Grid1Class With {.Дата = y.Дата, .Сообщение = y.Событие, .Маршрут = y.Маршрут}).ToList()
        If f IsNot Nothing Then
            For Each b In f
                b.Номер = i
                Grid1all.Add(b)
                i += 1
            Next

            bsGrid1.ResetBindings(False)
        End If
    End Sub
    Private Async Sub predAsync()
        Await Task.Run(Sub() pred())
    End Sub
    Private Sub pred()
        Dim mo As New AllUpd
        Do While AllClass.ЖурналПеревозчик Is Nothing
            mo.ЖурналПеревозчикAll()
        Loop
        Do While AllClass.ЖурналПеревозчикСобытия Is Nothing
            mo.ЖурналПеревозчикСобытияAll()
        Loop
    End Sub
    Public Class Grid1Class
        Public Property Номер As Integer
        Public Property Дата As String
        Public Property Маршрут As String
        Public Property Сообщение As String

    End Class

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Close()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        addnewAsync(Grid1all)
        MessageBox.Show("Данные приняты!", Рик)
    End Sub
    Private Async Sub addnewAsync(ByVal lst As List(Of Grid1Class))
        Await Task.Run(Sub() addnew(lst))
    End Sub
    Private Sub addnew(ByVal lst As List(Of Grid1Class))
        Dim mo As New AllUpd
        Using db As New dbAllDataContext(_cn3)
            Dim f = db.ЖурналПеревозчикСобытия.Where(Function(x) x.IDЖурналПеревозчик = ID).ToList()
            If f IsNot Nothing Then
                db.ЖурналПеревозчикСобытия.DeleteAllOnSubmit(f)
                db.SubmitChanges()
                If lst IsNot Nothing Then
                    Dim mnew As New List(Of ЖурналПеревозчикСобытия)
                    For Each b In lst
                        Dim m As New ЖурналПеревозчикСобытия With {.IDЖурналПеревозчик = ID, .Дата = b.Дата, .Событие = b.Сообщение, .Маршрут = b.Маршрут}
                        mnew.Add(m)
                    Next
                    db.ЖурналПеревозчикСобытия.InsertAllOnSubmit(mnew)
                    db.SubmitChanges()

                    Dim ml = db.ЖурналПеревозчик.Where(Function(x) x.Код = ID).FirstOrDefault()
                    If ml IsNot Nothing Then
                        ml.Состояние = lst.Count & " Сообщение(я)"
                        Rez = lst.Count & " Сообщение(я)"
                        db.SubmitChanges()
                    End If

                    mo.ЖурналПеревозчикСобытияAll()
                End If
            Else
                If lst IsNot Nothing Then
                    Dim mnew As New List(Of ЖурналПеревозчикСобытия)
                    For Each b In lst
                        Dim m As New ЖурналПеревозчикСобытия With {.IDЖурналПеревозчик = ID, .Дата = b.Дата, .Событие = b.Сообщение, .Маршрут = b.Маршрут}
                        mnew.Add(m)
                    Next
                    db.ЖурналПеревозчикСобытия.InsertAllOnSubmit(mnew)
                    db.SubmitChanges()

                    Dim ml = db.ЖурналПеревозчик.Where(Function(x) x.Код = ID).FirstOrDefault()
                    If ml IsNot Nothing Then
                        ml.Состояние = lst.Count & " Сообщение(я)"
                        Rez = lst.Count & " Сообщение(я)"
                        db.SubmitChanges()
                    End If



                    mo.ЖурналПеревозчикСобытияAll()
                End If

            End If


        End Using
    End Sub
    Private RowInd As Integer = Nothing

    Private Sub Grid1_MouseDown(sender As Object, e As MouseEventArgs) Handles Grid1.MouseDown
        If e.Button = MouseButtons.Right Then
            RowInd = 0
            Dim k = TryCast(sender, DataGridView)
            Dim m = k.CurrentCell.ColumnIndex

            If m = 1 Then
                RowInd = k.CurrentCell.RowIndex
                ContextMenuStrip1.Show(MousePosition, ToolStripDropDownDirection.Right)
            End If
        End If
    End Sub

    Private Sub СегодняToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles СегодняToolStripMenuItem.Click
        If RowInd >= 0 Then
            Grid1all(RowInd).Дата = Now
            bsGrid1.ResetBindings(False)
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If Grid1.CurrentCell.RowIndex >= 0 Then
            Grid1all.RemoveAt(Grid1.CurrentCell.RowIndex)
            bsGrid1.ResetBindings(False)
        End If
    End Sub
End Class