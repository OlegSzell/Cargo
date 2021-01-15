Imports System.ComponentModel

Public Class ЖурналКлиентДатыТаблица
    Private D As New List(Of ЖурналКлиентДаты)
    Dim Grid1all As List(Of Grid1Class)
    Private bsD As BindingSource
    Private ID As Integer
    Public GRD As New List(Of Grid1Class)
    Sub New(ByVal _d As List(Of ЖурналКлиентДаты), ByVal _ID As Integer)
        D = _d
        ID = _ID
        ' Этот вызов является обязательным для конструктора.
        InitializeComponent()

        ' Добавить код инициализации после вызова InitializeComponent().

    End Sub
    Private Sub ЖурналКлиентДатыТаблица_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Grid1all = New List(Of Grid1Class)
        bsD = New BindingSource
        bsD.DataSource = Grid1all
        Grid1.DataSource = bsD
        GridView(Grid1)
        Grid1.SelectionMode = DataGridViewSelectionMode.CellSelect
        Grid1.Columns(0).Visible = False
        Grid1.Columns(8).Visible = False
        Grid1.Columns(1).Width = 60
        Grid1.Columns(6).ReadOnly = True
        Grid1.Columns(7).Width = 90
        Grid1.Columns(7).HeaderText = "Добавть в новые рейсы"

        If D IsNot Nothing Then
            Dim i As Integer = 1
            For Each b In D
                Dim m As New Grid1Class With {.Номер = i, .ID = ID, .ДатаЗагрузки = b.ДатаЗагрузки, .ДатаВыгрузки = b.ДатаДоставки, .Ставка = b.Ставка, .ДопУсловия = b.ДопУсловия, .Состояние = b.Состояние, .IDЖурналКлиентДаты = b.ID}
                Grid1all.Add(m)
                i += 1
            Next
            bsD.ResetBindings(False)
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        GRD.AddRange(Grid1all)
        Close()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        GRD.AddRange(Grid1all)
        UpdAsync(Grid1all)
        MessageBox.Show("Данные приняты!", Рик)
    End Sub
    Private Async Sub UpdAsync(ByVal d As List(Of Grid1Class))
        Await Task.Run(Sub() Upd(d))
    End Sub
    Private Sub Upd(ByVal d As List(Of Grid1Class))
        Using db As New dbAllDataContext()
            For Each b In d
                If b.IDЖурналКлиентДаты > 0 Then
                    Dim f2 = db.ЖурналКлиентДаты.Where(Function(x) x.IDЖурналКлиентГруз = b.ID And x.ID = b.IDЖурналКлиентДаты).FirstOrDefault()
                    If f2 IsNot Nothing Then
                        With f2
                            .ДатаЗагрузки = b.ДатаЗагрузки
                            .ДатаДоставки = b.ДатаВыгрузки
                            .ДопУсловия = b.ДопУсловия
                            .Ставка = b.Ставка
                            .Состояние = b.Состояние
                        End With
                        db.SubmitChanges()
                    End If
                Else
                    Dim f2 As New ЖурналКлиентДаты
                    With f2
                        .IDЖурналКлиентГруз = b.ID
                        .ДатаЗагрузки = b.ДатаЗагрузки
                        .ДатаДоставки = b.ДатаВыгрузки
                        .ДопУсловия = b.ДопУсловия
                        .Ставка = b.Ставка
                        .Состояние = b.Состояние
                    End With
                    db.ЖурналКлиентДаты.InsertOnSubmit(f2)
                    db.SubmitChanges()

                End If

            Next

            Dim mo As New AllUpd
            mo.ЖурналКлиентДатыAll()

        End Using
    End Sub

    Private Sub Grid1_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles Grid1.CellEndEdit
        If Grid1all.ElementAt(e.RowIndex).ID = 0 Then
            Grid1.Rows(e.RowIndex).Cells(0).Value = ID
        End If
        If Grid1all.ElementAt(e.RowIndex).Номер Is Nothing Then
            Grid1all.ElementAt(e.RowIndex).Номер = e.RowIndex + 1
        End If
    End Sub

    Public Class Grid1Class

        Public Property ID As Integer
        Public Property Номер As String
        Public Property ДатаЗагрузки As String
        Public Property ДатаВыгрузки As String
        Public Property Ставка As String
        Public Property ДопУсловия As String
        Public Property Состояние As String
        Public Property ДобВТабл As Boolean
        Public Property IDЖурналКлиентДаты As Integer

    End Class

    Private Sub Grid1_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles Grid1.CellDoubleClick
        If e.ColumnIndex = 6 Then
            Dim m = Grid1all.ElementAt(e.RowIndex)
            Dim f As New ЖурналКлиентДатыТаблицаАктуальность(m.IDЖурналКлиентДаты)
            f.ShowDialog()
            If f.Rez IsNot Nothing Then
                m.Состояние = f.Rez
                bsD.ResetBindings(False)
            End If
        End If
    End Sub
    Private Async Sub DelAsync(f As List(Of Grid1Class))
        Await Task.Run(Sub() Del(f))

    End Sub
    Private Sub Del(f As List(Of Grid1Class))
        Dim mo As New AllUpd
        Using db As New dbAllDataContext()
            For Each b In f
                Dim m = db.ЖурналКлиентДаты.Where(Function(x) x.ID = b.IDЖурналКлиентДаты).FirstOrDefault()
                If m IsNot Nothing Then
                    db.ЖурналКлиентДаты.DeleteOnSubmit(m)
                    db.SubmitChanges()
                End If
            Next

        End Using
        mo.ЖурналКлиентДатыAll()

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim f = Grid1all.Where(Function(x) x.ДобВТабл = True).ToList()
        If f IsNot Nothing Then
            If f.Count > 0 Then
                If MessageBox.Show("Удалить выбранные даты загрузок", Рик, MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.No Then
                    Return
                End If

                Del(f)
                For Each b In f
                    Grid1all.Remove(b)
                Next
                Dim i As Integer = 1
                For Each v In Grid1all
                    v.Номер = i
                    i += 1
                Next
                bsD.ResetBindings(False)

            End If
        End If
    End Sub
End Class