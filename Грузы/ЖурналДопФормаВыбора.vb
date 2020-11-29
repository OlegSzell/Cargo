Imports System.ComponentModel

Public Class ЖурналДопФормаВыбора
    Private lst As List(Of Журнал.ForSeachtxt1)
    Private bslst As BindingSource
    Private Grid1all As BindingList(Of Grid2ЖурналClass)
    Private list1 As BindingList(Of String)
    Private bslist1 As BindingSource
    Public Sel As Grid2ЖурналClass = Nothing

    Sub New(ByVal _lst As List(Of Журнал.ForSeachtxt1))
        lst = _lst
        ' Этот вызов является обязательным для конструктора.
        InitializeComponent()

        ' Добавить код инициализации после вызова InitializeComponent().

    End Sub

    Private Sub ЖурналДопФормаВыбора_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Grid1all = New BindingList(Of Grid2ЖурналClass)
        bslst = New BindingSource
        bslst.DataSource = Grid1all
        Grid1.DataSource = bslst
        GridView(Grid1)
        Grid1.Columns(0).Width = 60
        Grid1.Columns(0).HeaderText = "№"
        Grid1.Columns(2).HeaderText = "Дата создания"
        Grid1.Columns(2).Width = 80
        Grid1.Columns(3).Width = 80
        Grid1.Columns(3).HeaderText = "Дата загрузки"
        Grid1.Columns(6).Visible = False
        Grid1.Columns(7).Visible = False
        Grid1.Columns(8).Visible = False


        list1 = New BindingList(Of String)
        bslist1 = New BindingSource
        bslist1.DataSource = list1
        ListBox1.DataSource = bslist1


        Dim i As Integer = 1
        For Each b In lst
            Dim b1 As Com1ЖурналClass = b.Данные
            Dim маршрут As String = Nothing
            For Each b2 In b1.Lst
                If маршрут Is Nothing Then
                    маршрут = b2.СтранаПогрузки & ", " & IIf(b2.КвадратПогрузки Is Nothing, b2.ГородПогрузки, b2.КвадратПогрузки) & IIf(b2.КвадратПогрузки IsNot Nothing, ", " & b2.ГородПогрузки, "") & " - " &
                        b2.СтранаВыгрузки & ", " & IIf(b2.КвадратВыгрузки Is Nothing, b2.ГородВыгрузки, b2.КвадратВыгрузки) & IIf(b2.КвадратВыгрузки IsNot Nothing, ", " & b2.ГородВыгрузки, "")
                Else
                    маршрут = маршрут & vbCrLf & b2.СтранаПогрузки & ", " & IIf(b2.КвадратПогрузки Is Nothing, b2.ГородПогрузки, b2.КвадратПогрузки) & IIf(b2.КвадратПогрузки IsNot Nothing, ", " & b2.ГородПогрузки, "") & " - " &
                        b2.СтранаВыгрузки & ", " & IIf(b2.КвадратВыгрузки Is Nothing, b2.ГородВыгрузки, b2.КвадратВыгрузки) & IIf(b2.КвадратВыгрузки IsNot Nothing, ", " & b2.ГородВыгрузки, "")
                End If
            Next
            Dim f As New Grid2ЖурналClass With {.Mаршрут = маршрут, .КодЖурналГруз = b1.КодГруза, .Дата = b1.СписокДат, .Номер = i, .Груз = b1.Наименование, .ДатаЗагрузки = b1.ДатаЗагрузки, .Клиент = b1.Lst(0).Клиент}
            Grid1all.Add(f)
            i += 1
        Next



    End Sub

    Private Sub Grid1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles Grid1.CellClick
        Dim f As Grid2ЖурналClass = Grid1all.ElementAt(e.RowIndex)
        Dim f1 As Com1ЖурналClass = lst.Where(Function(x) x.Данные.КодГруза = f.КодЖурналГруз).Select(Function(x) x.Данные).FirstOrDefault()
        Dim mar As String = Nothing
        Dim mar1 As String = Nothing
        Dim tz1 As String = Nothing
        Dim tz2 As String = Nothing
        For Each b In f1.Lst
            If mar Is Nothing Then
                mar = b.СтранаПогрузки & " (" & b.ГородПогрузки & ") " & vbCrLf & b.СтранаВыгрузки & " (" & b.ГородВыгрузки & ") "
            Else
                mar = mar & vbCrLf & b.СтранаПогрузки & " (" & b.ГородПогрузки & ") " & vbCrLf & b.СтранаВыгрузки & " (" & b.ГородВыгрузки & ") "
            End If


            If tz1 Is Nothing Then
                tz1 = "Таможня отправления: " & b.ТаможняОтправления
            Else
                tz1 = tz1 & vbCrLf & "Таможня отправления: " & vbCrLf & b.ТаможняОтправления
            End If

            If tz2 Is Nothing Then
                tz2 = "Таможня назначения: " & b.ТаможняНазначения
            Else
                tz2 = tz2 & vbCrLf & "Таможня назначения: " & b.ТаможняНазначения
            End If
        Next
        Dim mr = "Маршрут: " & mar
        Dim dat = "Дата загрузки: " & f1.ДатаЗагрузки & "."
        Dim dat1 = "Дата выгрузки: " & f1.ДатаВыгрузки & "."
        Dim gr = "Груз: " & f1.Наименование
        Dim kl = "Клиент: " & f1.Lst(0).Клиент

        Dim f4 As String() = {kl, mr, dat, dat1, gr, tz1, tz2}
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




    End Sub

    Private Sub Grid1_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles Grid1.CellDoubleClick
        Dim f As Grid2ЖурналClass = Grid1all.ElementAt(e.RowIndex)
        Sel = f
        Close()
    End Sub
End Class