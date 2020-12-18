Public Class ЖурналCombo3СводнаяПоКлиентам
    Private Naz As String = Nothing
    Private Grid1all As List(Of Grid1Class)
    Private bsgrid1 As BindingSource
    Sub New(ByVal _Naz As String)
        Naz = _Naz

        ' Этот вызов является обязательным для конструктора.
        InitializeComponent()
        Grid1all = New List(Of Grid1Class)
        bsgrid1 = New BindingSource
        bsgrid1.DataSource = Grid1all
        Grid1.DataSource = bsgrid1
        GridView(Grid1)

        ' Добавить код инициализации после вызова InitializeComponent().

    End Sub
    Private Sub ЖурналCombo3СводнаяПоКлиентам_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Grid1.Columns(0).Width = 50
        If Naz Is Nothing Then Return
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

        Do While AllClass.ЖурналПеревозчик Is Nothing
            mo.ЖурналПеревозчикAll()
        Loop

        Dim f = (From x In AllClass.ЖурналДата
                 Join y In AllClass.ЖурналКлиентГруз On x.Код Equals y.КодЖурналДата
                 Join z In AllClass.ЖурналКлиентМаршрут On y.Код Equals z.КодЖурналКлиентГруз
                 Join c In AllClass.ЖурналПеревозчик On y.Код Equals c.КодЖурналКлиентГруз
                 Where y.Клиент = Naz
                 Select x, y, z, c).ToList()
        If f IsNot Nothing Then
            If f.Count = 0 Then
                Dim f1 = (From x In AllClass.ЖурналДата
                          Join y In AllClass.ЖурналКлиентГруз On x.Код Equals y.КодЖурналДата
                          Join z In AllClass.ЖурналКлиентМаршрут On y.Код Equals z.КодЖурналКлиентГруз
                          Where y.Клиент = Naz
                          Select x, y, z).ToList()

                If f1 IsNot Nothing Then
                    If f1.Count > 0 Then
                        БезПеревозов(f1)
                        Dim f2 = From v In f1
                                 Group v By Keys = New With {Key v.x.Дата}
                             Into Group
                                 Select New Grid1Class With {.Дата = Keys.Дата, .Груз = Group.Select(Function(n) n.y?.НаименованиеГруза _
                                     & vbCrLf & n.y.Вес & vbCrLf & n.y.Обьем).FirstOrDefault, .ДопИнфо = Group.Select(Function(n) n.y.ДополнитИнформация).FirstOrDefault}


                        Grid1all.AddRange(f2)
                        bsgrid1.ResetBindings(False)
                    End If



                End If
            Else ' с перевозами

            End If
        End If
    End Sub
    Private Sub БезПеревозов(Of T)(ByVal D As List(Of T))


    End Sub
    Public Class Grid1Class
        Public Property Номер As Integer
        Public Property Дата As String
        Public Property Груз As String
        Public Property Перевозчик As String
        Public Property Состояние As String
        Public Property ДопИнфо As String

    End Class
End Class