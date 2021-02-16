Imports System.ComponentModel

Public Class ЖурналДопФормаВыбора
    Private lst As List(Of Grid2ЖурналClass)
    Private bslst As BindingSource
    Private Grid1all As List(Of Grid2ЖурналClass)

    Private list1 As BindingList(Of String)
    Private bslist1 As BindingSource
    Public Sel As Grid2ЖурналClass = Nothing

    Sub New(ByVal _lst As List(Of Grid2ЖурналClass))
        lst = _lst
        ' Этот вызов является обязательным для конструктора.
        InitializeComponent()

        ' Добавить код инициализации после вызова InitializeComponent().

    End Sub

    Private Sub ЖурналДопФормаВыбора_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Grid1all = New List(Of Grid2ЖурналClass)
        bslst = New BindingSource
        bslst.DataSource = Grid1all
        Grid1.DataSource = bslst
        GridView(Grid1)
        Grid1.Columns(0).Width = 30
        Grid1.Columns(1).Width = 80
        Grid1.Columns(2).Width = 80
        Grid1.Columns(3).Width = 100
        Grid1.Columns(0).HeaderText = "№"
        Grid1.Columns(4).Width = 80
        Grid1.Columns(5).Width = 120
        Grid1.Columns(7).Width = 80
        Grid1.Columns(8).Width = 120
        Grid1.Columns(9).Width = 50


        Grid1.Columns(7).Visible = False
        Grid1.Columns(9).Visible = False
        Grid1.Columns(2).Visible = False
        Grid1.Columns(10).Visible = False
        Grid1.Columns(11).Visible = False
        Grid1.Columns(12).Visible = False
        Grid1.Columns(13).Visible = False
        Grid1.Columns(14).Visible = False

        Grid1.Columns(12).Width = 50


        list1 = New BindingList(Of String)
        bslist1 = New BindingSource
        bslist1.DataSource = list1
        ListBox1.DataSource = bslist1

        Dim jk = lst.OrderBy(Function(x) x.Клиент).ToList()
        Dim i As Integer = 1
        For Each b In jk
            b.Номер = i
            i += 1
        Next
        Grid1all.AddRange(jk)
        bslst.ResetBindings(False)




        'Dim i As Integer = 1
        'For Each b In lst
        '    Dim b1 As Com1ЖурналClass = b.Данные
        '    Dim маршрут As String = Nothing
        '    For Each b2 In b1.Lst
        '        If маршрут Is Nothing Then
        '            маршрут = b2.СтранаПогрузки & ", " & IIf(b2.КвадратПогрузки Is Nothing, b2.ГородПогрузки, b2.КвадратПогрузки) & IIf(b2.КвадратПогрузки IsNot Nothing, ", " & b2.ГородПогрузки, "") & " - " &
        '                b2.СтранаВыгрузки & ", " & IIf(b2.КвадратВыгрузки Is Nothing, b2.ГородВыгрузки, b2.КвадратВыгрузки) & IIf(b2.КвадратВыгрузки IsNot Nothing, ", " & b2.ГородВыгрузки, "")
        '        Else
        '            маршрут = маршрут & vbCrLf & b2.СтранаПогрузки & ", " & IIf(b2.КвадратПогрузки Is Nothing, b2.ГородПогрузки, b2.КвадратПогрузки) & IIf(b2.КвадратПогрузки IsNot Nothing, ", " & b2.ГородПогрузки, "") & " - " &
        '                b2.СтранаВыгрузки & ", " & IIf(b2.КвадратВыгрузки Is Nothing, b2.ГородВыгрузки, b2.КвадратВыгрузки) & IIf(b2.КвадратВыгрузки IsNot Nothing, ", " & b2.ГородВыгрузки, "")
        '        End If
        '    Next
        '    Dim f As New Grid2ЖурналClass With {.Страны = маршрут, .КодЖурналГруз = b1.КодГруза, .Дата = b1.СписокДат, .Номер = i, .Груз = b1.Наименование, .ДатаЗагрузки = b1.ДатаЗагрузки, .Клиент = b1.Lst(0).Клиент}
        '    Grid1all.Add(f)
        '    i += 1
        'Next



    End Sub

    Private Sub Grid1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles Grid1.CellClick
        Dim f As Grid2ЖурналClass = Grid1all.ElementAt(e.RowIndex)


        '     Public Property Номер As Integer
        'Public Property Клиент As String
        'Public Property Дата As Date
        'Public Property Страны As String
        'Public Property ДатаЗагрузки As String
        'Public Property Города As String
        'Public Property Груз As String
        'Public Property Состояние As String
        'Public Property Ставка As String
        'Public Property Инфо As String






        'Dim f1 As Com1ЖурналClass = lst.Where(Function(x) x.Данные.КодГруза = f.КодЖурналГруз).Select(Function(x) x.Данные).FirstOrDefault()


        Dim kl1 = "Клиент: " & f.Клиент & vbCrLf
        Dim mr = "Страны: " & f.Страны & vbCrLf
        Dim dat = "Дата загрузки: " & f.ДатаЗагрузки & "." & vbCrLf
        Dim dat1 = "Города: " & f.Города & "." & vbCrLf
        Dim gr = "Груз: " & f.Груз & vbCrLf
        Dim kl2 = "Состояние: " & f.Состояние & vbCrLf
        Dim kl = "Ставка: " & f.Ставка
        Dim mfg = " -- "

        Dim f4 As String() = {kl1, mfg, mr, mfg, dat, mfg, dat1, mfg, gr, mfg, kl2, mfg, kl}
        Dim f2 As New List(Of String)
        f2.AddRange(f4)

        If list1 IsNot Nothing Then
            list1.Clear()
        End If
        If f2 IsNot Nothing Then
            If f2.Count > 0 Then
                For Each b In f2
                    list1.Add(b)
                Next

            End If
        End If
        bslist1.ResetBindings(False)



    End Sub

    Private Sub Grid1_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles Grid1.CellDoubleClick
        'Dim f As Grid2ЖурналClass = Grid1all.ElementAt(e.RowIndex)
        'Sel = f
        'Close()
    End Sub
End Class