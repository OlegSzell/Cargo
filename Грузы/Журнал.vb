Public Class Журнал
    Private com1All As List(Of AllClientClass)
    Private Grid2all As List(Of Grid2ЖурналClass)
    Private bsgrid2 As BindingSource
    Private Sub Журнал_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.MdiParent = MDIParent1
        Label1.Text = Now.Date
        ПредзагрузкаAsync()
        Grid2Met()
    End Sub
    Private Sub Grid2Met()
        Grid2all = New List(Of Grid2ЖурналClass)
        bsgrid2 = New BindingSource
        bsgrid2.DataSource = Grid2all
        Grid2.DataSource = bsgrid2
        GridView(Grid2)
        Grid2.Columns(0).Width = 60
        Grid2.Columns(0).HeaderText = "№"
        Grid2.Columns(2).Width = 80
        Grid2.Columns(3).Width = 80
        Grid2.Columns(3).HeaderText = "Дата загрузки"

        Dim mo As New AllUpd
        Do While AllClass.ЖурналДата Is Nothing
            mo.ЖурналДатаAll()
        Loop
        Do While AllClass.ЖурналКлиентГруз Is Nothing
            mo.ЖурналКлиентГрузAll()
        Loop
        Do While AllClass.ЖурналКлиентМаршрут Is Nothing
            mo.ЖурналКлиентМаршрутAll()
        Loop

        Dim f = (From x In AllClass.ЖурналДата
                 Join y In AllClass.ЖурналКлиентГруз On x.Код Equals y.КодЖурналДата
                 Join z In AllClass.ЖурналКлиентМаршрут On y.Код Equals z.КодЖурналКлиентГруз
                 Order By x.Дата Descending
                 Select x, y, z).ToList

        Dim f1 = (From x In f
                  Group x By Keys = New With {Key x.y.Клиент, Key x.x.Дата, Key x.y.ДатаЗагрузки}
                      Into Group
                  Select New Grid2ЖурналClass With {.Клиент = Keys.Клиент, .Дата = Keys.Дата, .ДатаЗагрузки = IIf(Keys.ДатаЗагрузки Is Nothing, Nothing, Keys.ДатаЗагрузки),
                     .МаршрутList = (From p In Group
                                     Select "( " & p.z.СтранаПогрузки & " ) " & p.z.ГородПогрузки & " - " & "( " & p.z.СтранаВыгрузки & " ) " & p.z.ГородВыгрузки).ToList()}).ToList()



        If Grid2all IsNot Nothing Then
            Grid2all.Clear()
        End If
        Dim i As Integer = 1
        For Each b In f1
            Dim k As String = Nothing
            For Each b1 In b.МаршрутList
                If k Is Nothing Then
                    k = b1
                Else
                    k = k & vbCrLf & b1
                End If

            Next
            Dim f3 As New Grid2ЖурналClass With {.Дата = b.Дата, .ДатаЗагрузки = b.ДатаЗагрузки, .Клиент = b.Клиент, .Номер = i, .Mаршрут = k}

            i += 1
            Grid2all.Add(f3)
        Next

        bsgrid2.ResetBindings(False)

    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        Dim f As New ЖурналДобавитьГруз
        f.ShowDialog()

    End Sub
    Private Async Sub ПредзагрузкаAsync()
        Await Task.Run(Sub() Предзагрузка())
    End Sub
    Private Sub Предзагрузка()
        Dim mo As New AllUpd
        Do While AllClass.Клиент Is Nothing
            mo.КлиентAll()
        Loop
        Do While AllClass.ЖурналДата Is Nothing
            mo.ЖурналДатаAll()
        Loop
        Do While AllClass.ЖурналКлиентГруз Is Nothing
            mo.ЖурналКлиентГрузAll()
        Loop
        Do While AllClass.ЖурналКлиентМаршрут Is Nothing
            mo.ЖурналКлиентМаршрутAll()
        Loop

        Do While AllClass.ЖурналКлиентСписок Is Nothing
            mo.ЖурналКлиентСписокAll()
        Loop
    End Sub
End Class