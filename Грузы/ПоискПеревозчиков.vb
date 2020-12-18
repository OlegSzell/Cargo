Option Explicit On
Imports System.Data.OleDb
Imports System.Threading
Public Class ПоискПеревозчиков
    Public Da As New OleDbDataAdapter 'Адаптер
    Public Ds As New DataSet 'Пустой набор записей
    Dim tbl As New DataTable
    Dim загр As String
    Dim ind As Boolean = False
    Dim удал As Integer
    Private Delegate Sub comb2()
    Private Delegate Sub comb31()
    Private Delegate Sub comb32()


    Private com1all As List(Of Com1Class)
    Private bs1com As BindingSource

    Private com2all As List(Of IDNaz)
    Private bs2com As BindingSource

    Private com3all As List(Of IDNaz)
    Private bs3com As BindingSource

    Private Grid1all As List(Of Grid1Class)
    Private bsGrid1all As BindingSource
    Sub New()

        ' Этот вызов является обязательным для конструктора.
        InitializeComponent()
        ПредзагрузкаAsync()
        ' Добавить код инициализации после вызова InitializeComponent().

    End Sub
    Public Class Grid1Class
        Public Property Номер As Integer
        Public Property Организация As String
        Public Property Контакт As String
        Public Property Телефон As String
        Public Property Страны As String
        Public Property Регионы As String
        Public Property ADR As String
        Public Property КоличАвто As String
        Public Property ТипАвто As String
        Public Property Тоннаж As String
        Public Property Объем As String
        Public Property Ставка As String
        Public Property Примечание As String
        Public Property IDПеревоз As String

    End Class
    Public Class Com1Class
        Public Property Naz As String
        Public Property lst As List(Of String)
    End Class
    Private Sub COM4()

        If ComboBox1.InvokeRequired Then
            Me.Invoke(New comb2(AddressOf COM4))
        Else
            Dim strsql As String = "SELECT DISTINCT Страна FROM Страна ORDER BY Страна"
            Dim ds As DataTable = Selects3(strsql)
            Me.ComboBox1.AutoCompleteCustomSource.Clear()
            Me.ComboBox1.Items.Clear()
            For Each r As DataRow In ds.Rows
                Me.ComboBox1.AutoCompleteCustomSource.Add(r.Item(0).ToString())
                Me.ComboBox1.Items.Add(r(0).ToString)
            Next
        End If
    End Sub
    Private Sub COM5()

        If ComboBox3.InvokeRequired Then
            Me.Invoke(New comb31(AddressOf COM5))
        Else
            Dim strsql As String = "SELECT DISTINCT [Наименование фирмы] FROM ПеревозчикиБаза ORDER BY [Наименование фирмы]"
            Dim ds As DataTable = Selects3(strsql)
            Me.ComboBox3.AutoCompleteCustomSource.Clear()
            Me.ComboBox3.Items.Clear()
            For Each r As DataRow In ds.Rows
                Me.ComboBox3.AutoCompleteCustomSource.Add(r.Item(0).ToString())
                Me.ComboBox3.Items.Add(r(0).ToString)
            Next
        End If
    End Sub
    Public Sub Запуск()
        If TextBox1.InvokeRequired Then
            Me.Invoke(New comb32(AddressOf Запуск))
        Else
            TextBox1.Text = DateTime.Now.ToString("dd.MM.yyyy")
        End If
        Dim h As New Thread(AddressOf УскорЗагр)
        h.IsBackground = True
        h.SetApartmentState(ApartmentState.STA)
        h.Start()
    End Sub
    Private Sub com1Metod()
        Dim mo As New AllUpd
        Do While AllClass.Страна Is Nothing
            mo.СтранаAll()
        Loop

        Dim f = (From b In (From x In AllClass.Страна
                            Join y In AllClass.РегионыРоссии On x.Код Equals y.Страны
                            Order By x.Страна).ToList
                 Group b By Keys = New With {Key b.x.Страна}
                      Into Group
                 Select New Com1Class With {.Naz = Keys.Страна, .lst = (From c In Group
                                                                        Order By c.y.Регионы
                                                                        Select c.y.Регионы).ToList()}).ToList()
        If com1all IsNot Nothing Then
            com1all.Clear()
        End If


        If f IsNot Nothing Then
            com1all.AddRange(f)
            bs1com.ResetBindings(False)
        End If

        ComboBox1.Text = String.Empty

    End Sub
    Private Sub com3Metod()
        Dim mo As New AllUpd
        Do While AllClass.ПеревозчикиБаза Is Nothing
            mo.ПеревозчикиБазаAll()
        Loop
        If com3all IsNot Nothing Then
            com3all.Clear()
        End If
        Dim f = AllClass.ПеревозчикиБаза.OrderBy(Function(x) x.Наименование_фирмы).Select(Function(x) x.Наименование_фирмы).Distinct.ToList()
        If f IsNot Nothing Then
            For Each b In f
                Dim m As New IDNaz With {.Naz = b}
                com3all.Add(m)
            Next
            bs3com.ResetBindings(False)
        End If
    End Sub
    Private Async Sub ПредзагрузкаAsync()
        Await Task.Run(Sub() Предзагрузка())
    End Sub
    Private Sub Предзагрузка()
        Dim mo As New AllUpd
        Do While AllClass.Страна Is Nothing
            mo.СтранаAll()
        Loop

        Do While AllClass.ПеревозчикиБаза Is Nothing
            mo.ПеревозчикиБазаAll()
        Loop
        Do While AllClass.РегионыРоссии Is Nothing
            mo.РегионыРоссииAll()
        Loop
    End Sub

    Private Sub ПоискПеревозчиков_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        com1all = New List(Of Com1Class)
        bs1com = New BindingSource
        bs1com.DataSource = com1all
        ComboBox1.DataSource = bs1com
        ComboBox1.DisplayMember = "Naz"

        com2all = New List(Of IDNaz)
        bs2com = New BindingSource
        bs2com.DataSource = com2all
        ComboBox2.DataSource = bs2com
        ComboBox2.DisplayMember = "Naz"

        com3all = New List(Of IDNaz)
        bs3com = New BindingSource
        bs3com.DataSource = com3all
        ComboBox3.DataSource = bs3com
        ComboBox3.DisplayMember = "Naz"

        Grid1all = New List(Of Grid1Class)
        bsGrid1all = New BindingSource
        bsGrid1all.DataSource = Grid1all
        Grid1.DataSource = bsGrid1all
        GridView(Grid1)
        Grid1.Columns(12).Visible = False
        Grid1.SelectionMode = DataGridViewSelectionMode.FullRowSelect

        Grid1.Columns(0).Width = 150
        Grid1.Columns(1).Width = 150
        Grid1.Columns(2).Width = 150
        Grid1.Columns(4).Width = 150
        Grid1.Columns(11).Width = 300

        TextBox1.Text = DateTime.Now.ToString("dd.MM.yyyy")


        ComboBox1.BeginInvoke(New MethodInvoker(Sub() com1Metod()))



        Запуск()

    End Sub
    Private Sub УскорЗагр()
        Dim bn As New Thread(AddressOf COM4)
        Dim bn1 As New Thread(AddressOf COM5)
        bn.IsBackground = True
        bn1.IsBackground = True
        bn.Start()
        bn1.Start()


    End Sub


    Public Sub ОбнКом1()

        If Not IsDBNull(tbl) Then
            tbl.Clear()
        End If

        ComboBox2.Text = ""
        CheckBox1.Checked = False



        Dim StrSql3 As String
        StrSql3 = "SELECT DISTINCT Регионы
FROM РегионыРоссии INNER JOIN Страна ON РегионыРоссии.Страны = Страна.Код
Where Страна.Страна = '" & ComboBox1.Text & "'"
        Dim fg As New Thread(Sub() COMxt(Me, StrSql3, ComboBox2))
        fg.IsBackground = True
        fg.Start()

        'ComboBox2.Items.Clear()
        'ComboBox2.AutoCompleteCustomSource.Clear()

        'For Each r As DataRow In ds2.Rows
        '    Me.ComboBox2.AutoCompleteCustomSource.Add(r.Item(0).ToString())
        '    Me.ComboBox2.Items.Add(r(0).ToString)
        'Next

        refreshgrid()
    End Sub
    Private Sub refreshgrid()
        Me.Cursor = Cursors.WaitCursor
        If Not IsDBNull(tbl) Then
            tbl.Clear()
        End If



        Dim StrSql2 As String


        If ComboBox2.Text = "" Then
            StrSql2 = "SELECT ПеревозчикиБаза.[Наименование фирмы], ПеревозчикиБаза.[Контактное лицо],
ПеревозчикиБаза.Телефоны, ПеревозчикиБаза.[Страны перевозок], ПеревозчикиБаза.Регионы, ПеревозчикиБаза.ADR,
ПеревозчикиБаза.[Кол-во авто], ПеревозчикиБаза.Вид_авто, ПеревозчикиБаза.Тоннаж, ПеревозчикиБаза.Объем, ПеревозчикиБаза.Ставка,
ПеревозчикиБаза.Примечание, ПеревозчикиБаза.ID
From ПеревозчикиБаза Where ПеревозчикиБаза.[Страны перевозок] LIKE '%" & ComboBox1.Text & "%'"
        Else
            StrSql2 = "SELECT ПеревозчикиБаза.[Наименование фирмы], ПеревозчикиБаза.[Контактное лицо],
ПеревозчикиБаза.Телефоны, ПеревозчикиБаза.[Страны перевозок], ПеревозчикиБаза.Регионы, ПеревозчикиБаза.ADR,
ПеревозчикиБаза.[Кол-во авто], ПеревозчикиБаза.Вид_авто, ПеревозчикиБаза.Тоннаж, ПеревозчикиБаза.Объем, ПеревозчикиБаза.Ставка,
ПеревозчикиБаза.Примечание, ПеревозчикиБаза.ID
From ПеревозчикиБаза Where ПеревозчикиБаза.[Страны перевозок] LIKE '%" & ComboBox1.Text & "%' AND ПеревозчикиБаза.Регионы LIKE '%" & ComboBox2.Text & "%'"
        End If
        tbl = Selects3(StrSql2)

        Grid1.DataSource = tbl
        GridView(Grid1)
        Grid1.Columns(12).Visible = False
        Grid1.SelectionMode = DataGridViewSelectionMode.FullRowSelect

        Grid1.Columns(0).Width = 150
        Grid1.Columns(1).Width = 150
        Grid1.Columns(2).Width = 150
        Grid1.Columns(4).Width = 150
        Grid1.Columns(11).Width = 300
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged

        CheckBox1.Checked = False
        refreshgrid()

    End Sub

    Private Sub ComboBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox2.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True

            refreshgrid()
        End If
    End Sub


    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged

        If CheckBox1.Checked = True Then
            tbl.Clear()
            Dim adr As String = "есть"
            Dim StrSql2 As String
            StrSql2 = "SELECT ПеревозчикиБаза.[Наименование фирмы], ПеревозчикиБаза.[Контактное лицо],
ПеревозчикиБаза.Телефоны, ПеревозчикиБаза.[Страны перевозок], ПеревозчикиБаза.Регионы, ПеревозчикиБаза.ADR,
ПеревозчикиБаза.[Кол-во авто], ПеревозчикиБаза.Вид_авто, ПеревозчикиБаза.Тоннаж, ПеревозчикиБаза.Объем, ПеревозчикиБаза.Ставка,
ПеревозчикиБаза.Примечание,ПеревозчикиБаза.ID
From ПеревозчикиБаза
Where ПеревозчикиБаза.ADR='" & adr & "'"

            tbl = Selects3(StrSql2)

            Grid1.DataSource = tbl
            Grid1.Columns(12).Visible = False
            Grid1.Columns(0).Width = 250
            Grid1.Columns(1).Width = 250
            Grid1.Columns(2).Width = 150
            Grid1.Columns(11).Width = 320
            Grid1.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        End If
    End Sub
    Private Sub comb3()
        ComboBox1.Text = ""
        ComboBox2.Text = ""
        CheckBox1.Checked = False
        tbl.Clear()
        Dim StrSql6 As String
        StrSql6 = "SELECT ПеревозчикиБаза.[Наименование фирмы], ПеревозчикиБаза.[Контактное лицо],
ПеревозчикиБаза.Телефоны, ПеревозчикиБаза.[Страны перевозок], ПеревозчикиБаза.Регионы, ПеревозчикиБаза.ADR,
ПеревозчикиБаза.[Кол-во авто], ПеревозчикиБаза.Вид_авто, ПеревозчикиБаза.Тоннаж, ПеревозчикиБаза.Объем, ПеревозчикиБаза.Ставка,
ПеревозчикиБаза.Примечание, ПеревозчикиБаза.ID
From ПеревозчикиБаза
Where ПеревозчикиБаза.[Наименование фирмы] ='" & ComboBox3.Text & "'"

        tbl = Selects3(StrSql:=StrSql6)
        Grid1.DataSource = tbl
        Grid1.Columns(12).Visible = False
        Grid1.Columns(0).Width = 250
        Grid1.Columns(1).Width = 250
        Grid1.Columns(2).Width = 150
        Grid1.Columns(11).Width = 320
        Grid1.SelectionMode = DataGridViewSelectionMode.FullRowSelect
    End Sub
    Private Sub ComboBox3_SelectedValueChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedValueChanged
        comb3()

    End Sub

    'Private Sub Grid1_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs)
    '    idtabl = Nothing
    '    idtabl = Grid1.CurrentRow.Cells(12).Value
    '    ИзменПеревоз.Show()

    'End Sub
    'Private Sub Grid1_CellClick(sender As Object, e As DataGridViewCellEventArgs)


    '    If IsDBNull(e) Then
    '        Exit Sub
    '    End If
    '    удал = Nothing
    '    удал = Grid1.CurrentRow.Cells(12).Value
    'End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim StrSql6 As String
        If MessageBox.Show("Удалить перевозчика?", Рик, MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.No Then Exit Sub
        If Not удал = Nothing Then
            StrSql6 = "DELETE FROM ПеревозчикиБаза Where ID =" & удал & ""
            Updates3(stroka:=StrSql6)
            comb3()
        Else
            MessageBox.Show("Для удаления выберите перевозчика в таблице!", Рик)
        End If

    End Sub

    'Private Sub Grid1_ColumnHeaderMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs)
    '    For Each column As DataGridViewColumn In Grid1.Columns
    '        column.SortMode = DataGridViewColumnSortMode.NotSortable
    '    Next
    'End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        ДобПер = 1
        ОбнПер = 1
        ДобавитьПеревозчика.WindowState = 0
        ДобавитьПеревозчика.StartPosition = FormStartPosition.CenterScreen
        ДобавитьПеревозчика.ShowDialog()
    End Sub

    Private Sub Grid1_ColumnHeaderMouseClick_1(sender As Object, e As DataGridViewCellMouseEventArgs) Handles Grid1.ColumnHeaderMouseClick
        For Each column As DataGridViewColumn In Grid1.Columns
            column.SortMode = DataGridViewColumnSortMode.NotSortable
        Next
    End Sub

    Private Sub Grid1_CellClick_1(sender As Object, e As DataGridViewCellEventArgs) Handles Grid1.CellClick
        If IsDBNull(e) Then
            Exit Sub
        End If
        удал = Nothing
        удал = Grid1.CurrentRow.Cells(12).Value
    End Sub

    Private Sub Grid1_CellDoubleClick_1(sender As Object, e As DataGridViewCellEventArgs) Handles Grid1.CellDoubleClick
        idtabl = Nothing
        idtabl = Grid1.CurrentRow.Cells(12).Value
        Dim f As New ИзменПеревоз
        f.Show()
    End Sub

    Private Sub ComboBox1_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles ComboBox1.SelectionChangeCommitted
        ОбнКом1()
    End Sub
End Class