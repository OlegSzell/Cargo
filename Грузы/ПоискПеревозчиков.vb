Option Explicit On
Imports System.ComponentModel
Imports System.Data.OleDb
Imports System.Threading
Public Class ПоискПеревозчиков
    Public Da As New OleDbDataAdapter 'Адаптер
    Public Ds As New DataSet 'Пустой набор записей
    Dim tbl As New DataTable
    Dim загр As String
    Dim ind As Boolean = False
    Dim удал As Grid1Class = Nothing
    Private Delegate Sub comb2()
    Private Delegate Sub comb31()
    Private Delegate Sub comb32()


    Private com1all As List(Of Com1Class)
    Private bs1com As BindingSource

    Private com2all As List(Of IDNaz)
    Private bs2com As BindingSource

    Private com3all As List(Of IDNaz)
    Private bs3com As BindingSource

    Private Grid1all As BindingList(Of Grid1Class)
    'Private Grid1all As List(Of Grid1Class)
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
    'Private Sub COM4()

    '    If ComboBox1.InvokeRequired Then
    '        Me.Invoke(New comb2(AddressOf COM4))
    '    Else
    '        Dim strsql As String = "SELECT DISTINCT Страна FROM Страна ORDER BY Страна"
    '        Dim ds As DataTable = Selects3(strsql)
    '        Me.ComboBox1.AutoCompleteCustomSource.Clear()
    '        Me.ComboBox1.Items.Clear()
    '        For Each r As DataRow In ds.Rows
    '            Me.ComboBox1.AutoCompleteCustomSource.Add(r.Item(0).ToString())
    '            Me.ComboBox1.Items.Add(r(0).ToString)
    '        Next
    '    End If
    'End Sub
    'Private Sub COM5()

    '    If ComboBox3.InvokeRequired Then
    '        Me.Invoke(New comb31(AddressOf COM5))
    '    Else
    '        Dim strsql As String = "SELECT DISTINCT [Наименование фирмы] FROM ПеревозчикиБаза ORDER BY [Наименование фирмы]"
    '        Dim ds As DataTable = Selects3(strsql)
    '        Me.ComboBox3.AutoCompleteCustomSource.Clear()
    '        Me.ComboBox3.Items.Clear()
    '        For Each r As DataRow In ds.Rows
    '            Me.ComboBox3.AutoCompleteCustomSource.Add(r.Item(0).ToString())
    '            Me.ComboBox3.Items.Add(r(0).ToString)
    '        Next
    '    End If
    'End Sub
    'Public Sub Запуск()
    '    If TextBox1.InvokeRequired Then
    '        Me.Invoke(New comb32(AddressOf Запуск))
    '    Else
    '        TextBox1.Text = DateTime.Now.ToString("dd.MM.yyyy")
    '    End If
    '    Dim h As New Thread(AddressOf УскорЗагр)
    '    h.IsBackground = True
    '    h.SetApartmentState(ApartmentState.STA)
    '    h.Start()
    'End Sub
    Private Sub com1Metod()
        Dim mo As New AllUpd
        Do While AllClass.Страна Is Nothing
            mo.СтранаAll()
        Loop

        Do While AllClass.РегионыРоссии Is Nothing
            mo.РегионыРоссииAll()
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
        ComboBox3.Text = String.Empty
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

        Grid1all = New BindingList(Of Grid1Class)
        'Grid1all = New List(Of Grid1Class)
        bsGrid1all = New BindingSource
        bsGrid1all.DataSource = Grid1all
        Grid1.DataSource = bsGrid1all
        GridView(Grid1)
        Grid1.Columns(13).Visible = False
        'Grid1.SelectionMode = DataGridViewSelectionMode.FullRowSelect

        Grid1.Columns(0).Width = 60
        Grid1.Columns(1).Width = 150
        Grid1.Columns(2).Width = 100
        Grid1.Columns(3).Width = 110
        Grid1.Columns(4).Width = 150
        Grid1.Columns(5).MinimumWidth = 120
        Grid1.Columns(12).Width = 300

        TextBox1.Text = DateTime.Now.ToString("dd.MM.yyyy")


        ComboBox1.BeginInvoke(New MethodInvoker(Sub() com1Metod()))

        ComboBox3.BeginInvoke(New MethodInvoker(Sub() com3Metod()))

        'Запуск()

    End Sub
    'Private Sub УскорЗагр()
    '    Dim bn As New Thread(AddressOf COM4)
    '    Dim bn1 As New Thread(AddressOf COM5)
    '    bn.IsBackground = True
    '    bn1.IsBackground = True
    '    bn.Start()
    '    bn1.Start()


    'End Sub


    Public Sub ОбнКом1()
        If ComboBox1.SelectedItem Is Nothing Then Return
        Dim com1sel As Com1Class = ComboBox1.SelectedItem
        If com2all IsNot Nothing Then
            com2all.Clear()
        End If

        For Each b In com1sel.lst
            Dim m As New IDNaz With {.Naz = b}
            com2all.Add(m)
        Next
        bs2com.ResetBindings(False)
        ComboBox2.Text = String.Empty

        Dim mo As New AllUpd

        Do While AllClass.ПеревозчикиБаза Is Nothing
            mo.ПеревозчикиБазаAll()
        Loop

        If Grid1all IsNot Nothing Then
            Grid1all.Clear()
        End If

        Dim f = (From x In AllClass.ПеревозчикиБаза
                 Where x?.Страны_перевозок?.ToUpper.Contains(com1sel.Naz.ToUpper)
                 Select New Grid1Class With {.ADR = x.ADR, .IDПеревоз = x.ID, .КоличАвто = x.Кол_во_авто, .Контакт = x.Контактное_лицо, .Объем = x.Объем,
                    .Организация = x.Наименование_фирмы, .Примечание = x.Примечание, .Регионы = x.Регионы, .Ставка = x.Ставка, .Страны = x.Страны_перевозок,
                    .Телефон = x.Телефоны, .ТипАвто = x.Вид_авто, .Тоннаж = x.Тоннаж}).ToList()
        If f IsNot Nothing Then
            Dim f1 = f.OrderBy(Function(x) x.Организация).ToList()
            Dim i As Integer = 1
            For Each b In f1
                b.Номер = i
                Grid1all.Add(b)
                i += 1
            Next
        End If
        'bsGrid1all.ResetBindings(False)

    End Sub

    Private Sub refreshgrid()
        Dim mo As New AllUpd

        Do While AllClass.ПеревозчикиБаза Is Nothing
            mo.ПеревозчикиБазаAll()
        Loop

        Dim com1 As Com1Class = ComboBox1.SelectedItem
        Dim com2 As IDNaz = ComboBox2.SelectedItem



        If Grid1all IsNot Nothing Then
            Grid1all.Clear()
        End If


        If Not com1.Naz = com2.Naz Then

            Dim f = (From x In AllClass.ПеревозчикиБаза
                     Where x?.Страны_перевозок?.ToUpper.Contains(com1.Naz?.ToUpper) And x?.Регионы?.ToUpper.Contains(com2.Naz?.ToUpper)
                     Select New Grid1Class With {.ADR = x.ADR, .IDПеревоз = x.ID, .КоличАвто = x.Кол_во_авто, .Контакт = x.Контактное_лицо, .Объем = x.Объем,
                        .Организация = x.Наименование_фирмы, .Примечание = x.Примечание, .Регионы = x.Регионы, .Ставка = x.Ставка, .Страны = x.Страны_перевозок,
                        .Телефон = x.Телефоны, .ТипАвто = x.Вид_авто, .Тоннаж = x.Тоннаж}).ToList()
            If f IsNot Nothing Then
                Dim f1 = f.OrderBy(Function(x) x.Организация).ToList()
                Dim i As Integer = 1
                For Each b In f1
                    b.Номер = i
                    Grid1all.Add(b)
                    i += 1
                Next
            End If
            'bsGrid1all.ResetBindings(False)


        End If
    End Sub
    '    Private Sub refreshgrid()
    '        Me.Cursor = Cursors.WaitCursor
    '        If Not IsDBNull(tbl) Then
    '            tbl.Clear()
    '        End If



    '        Dim StrSql2 As String


    '        If ComboBox2.Text = "" Then
    '            StrSql2 = "SELECT ПеревозчикиБаза.[Наименование фирмы], ПеревозчикиБаза.[Контактное лицо],
    'ПеревозчикиБаза.Телефоны, ПеревозчикиБаза.[Страны перевозок], ПеревозчикиБаза.Регионы, ПеревозчикиБаза.ADR,
    'ПеревозчикиБаза.[Кол-во авто], ПеревозчикиБаза.Вид_авто, ПеревозчикиБаза.Тоннаж, ПеревозчикиБаза.Объем, ПеревозчикиБаза.Ставка,
    'ПеревозчикиБаза.Примечание, ПеревозчикиБаза.ID
    'From ПеревозчикиБаза Where ПеревозчикиБаза.[Страны перевозок] LIKE '%" & ComboBox1.Text & "%'"
    '        Else
    '            StrSql2 = "SELECT ПеревозчикиБаза.[Наименование фирмы], ПеревозчикиБаза.[Контактное лицо],
    'ПеревозчикиБаза.Телефоны, ПеревозчикиБаза.[Страны перевозок], ПеревозчикиБаза.Регионы, ПеревозчикиБаза.ADR,
    'ПеревозчикиБаза.[Кол-во авто], ПеревозчикиБаза.Вид_авто, ПеревозчикиБаза.Тоннаж, ПеревозчикиБаза.Объем, ПеревозчикиБаза.Ставка,
    'ПеревозчикиБаза.Примечание, ПеревозчикиБаза.ID
    'From ПеревозчикиБаза Where ПеревозчикиБаза.[Страны перевозок] LIKE '%" & ComboBox1.Text & "%' AND ПеревозчикиБаза.Регионы LIKE '%" & ComboBox2.Text & "%'"
    '        End If
    '        tbl = Selects3(StrSql2)

    '        Grid1.DataSource = tbl
    '        GridView(Grid1)
    '        Grid1.Columns(12).Visible = False
    '        Grid1.SelectionMode = DataGridViewSelectionMode.FullRowSelect

    '        Grid1.Columns(0).Width = 150
    '        Grid1.Columns(1).Width = 150
    '        Grid1.Columns(2).Width = 150
    '        Grid1.Columns(4).Width = 150
    '        Grid1.Columns(11).Width = 300
    '        Me.Cursor = Cursors.Default
    '    End Sub

    Private Sub ComboBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox2.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True

            refreshgrid()
        End If
    End Sub


    '    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged

    '        If CheckBox1.Checked = True Then
    '            tbl.Clear()
    '            Dim adr As String = "есть"
    '            Dim StrSql2 As String
    '            StrSql2 = "SELECT ПеревозчикиБаза.[Наименование фирмы], ПеревозчикиБаза.[Контактное лицо],
    'ПеревозчикиБаза.Телефоны, ПеревозчикиБаза.[Страны перевозок], ПеревозчикиБаза.Регионы, ПеревозчикиБаза.ADR,
    'ПеревозчикиБаза.[Кол-во авто], ПеревозчикиБаза.Вид_авто, ПеревозчикиБаза.Тоннаж, ПеревозчикиБаза.Объем, ПеревозчикиБаза.Ставка,
    'ПеревозчикиБаза.Примечание,ПеревозчикиБаза.ID
    'From ПеревозчикиБаза
    'Where ПеревозчикиБаза.ADR='" & adr & "'"

    '            tbl = Selects3(StrSql2)

    '            Grid1.DataSource = tbl
    '            Grid1.Columns(12).Visible = False
    '            Grid1.Columns(0).Width = 250
    '            Grid1.Columns(1).Width = 250
    '            Grid1.Columns(2).Width = 150
    '            Grid1.Columns(11).Width = 320
    '            Grid1.SelectionMode = DataGridViewSelectionMode.FullRowSelect
    '        End If
    '    End Sub
    Private Sub comb3()
        Dim mo As New AllUpd

        Do While AllClass.ПеревозчикиБаза Is Nothing
            mo.ПеревозчикиБазаAll()
        Loop

        If Grid1all IsNot Nothing Then
            Grid1all.Clear()
        End If

        ComboBox1.Text = String.Empty
        ComboBox2.Text = String.Empty
        CheckBox1.Checked = False

        Dim nz As IDNaz = ComboBox3.SelectedItem

        Dim f = AllClass.ПеревозчикиБаза.Where(Function(x) x.Наименование_фирмы = nz.Naz).FirstOrDefault()
        If f IsNot Nothing Then

            Dim f1 As New Grid1Class With {.ADR = f.ADR, .IDПеревоз = f.ID, .КоличАвто = f.Кол_во_авто, .Контакт = f.Контактное_лицо, .Объем = f.Объем,
                       .Организация = f.Наименование_фирмы, .Примечание = f.Примечание, .Регионы = f.Регионы, .Ставка = f.Ставка, .Страны = f.Страны_перевозок,
                       .Телефон = f.Телефоны, .ТипАвто = f.Вид_авто, .Тоннаж = f.Тоннаж, .Номер = 1}

            Grid1all.Add(f1)
        End If



        '        Dim StrSql6 As String
        '        StrSql6 = "SELECT ПеревозчикиБаза.[Наименование фирмы], ПеревозчикиБаза.[Контактное лицо],
        'ПеревозчикиБаза.Телефоны, ПеревозчикиБаза.[Страны перевозок], ПеревозчикиБаза.Регионы, ПеревозчикиБаза.ADR,
        'ПеревозчикиБаза.[Кол-во авто], ПеревозчикиБаза.Вид_авто, ПеревозчикиБаза.Тоннаж, ПеревозчикиБаза.Объем, ПеревозчикиБаза.Ставка,
        'ПеревозчикиБаза.Примечание, ПеревозчикиБаза.ID
        'From ПеревозчикиБаза
        'Where ПеревозчикиБаза.[Наименование фирмы] ='" & ComboBox3.Text & "'"

        '        tbl = Selects3(StrSql:=StrSql6)
        '        Grid1.DataSource = tbl
        '        Grid1.Columns(12).Visible = False
        '        Grid1.Columns(0).Width = 250
        '        Grid1.Columns(1).Width = 250
        '        Grid1.Columns(2).Width = 150
        '        Grid1.Columns(11).Width = 320
        '        Grid1.SelectionMode = DataGridViewSelectionMode.FullRowSelect
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

        If MessageBox.Show("Удалить перевозчика?", Рик, MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.No Then Exit Sub
        If удал IsNot Nothing Then
            CheckBox1.Checked = False
            DellAsync(удал)
            Dim m = удал.Организация
            com3all.Remove(com3all.Where(Function(x) x.Naz = m).FirstOrDefault())
            bs3com.ResetBindings(False)
            Grid1all.Remove(удал)
        Else
            MessageBox.Show("Для удаления выберите перевозчика в таблице!", Рик)
        End If

    End Sub
    Private Async Sub DellAsync(ByVal d As Grid1Class)
        Await Task.Run(Sub() Dell(d))
    End Sub
    Private Sub Dell(ByVal d As Grid1Class)
        Using db As New dbAllDataContext()
            Dim f = db.ПеревозчикиБаза.Where(Function(x) x.ID = d.IDПеревоз).FirstOrDefault()
            If f IsNot Nothing Then
                db.ПеревозчикиБаза.DeleteOnSubmit(f)
                db.SubmitChanges()
                Dim mo As New AllUpd
                mo.ПеревозчикиБазаAll()
            End If
        End Using

    End Sub

    'Private Sub Grid1_ColumnHeaderMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs)
    '    For Each column As DataGridViewColumn In Grid1.Columns
    '        column.SortMode = DataGridViewColumnSortMode.NotSortable
    '    Next
    'End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        ДобПер = 1
        ОбнПер = 1
        CheckBox1.Checked = False
        Dim f As New ДобавитьПеревозчика
        f.WindowState = 0
        f.StartPosition = FormStartPosition.CenterScreen
        f.ShowDialog()
        Dim mo As New AllUpd
        mo.ПеревозчикиБазаAll()
        ComboBox3.BeginInvoke(New MethodInvoker(Sub() com3Metod()))
    End Sub

    Private Sub Grid1_ColumnHeaderMouseClick_1(sender As Object, e As DataGridViewCellMouseEventArgs) Handles Grid1.ColumnHeaderMouseClick
        For Each column As DataGridViewColumn In Grid1.Columns
            column.SortMode = DataGridViewColumnSortMode.NotSortable
        Next
    End Sub

    Private Sub Grid1_CellClick_1(sender As Object, e As DataGridViewCellEventArgs) Handles Grid1.CellClick
        If e.RowIndex = -1 Then Return
        удал = Grid1all.ElementAt(e.RowIndex)

    End Sub

    Private Sub Grid1_CellDoubleClick_1(sender As Object, e As DataGridViewCellEventArgs) Handles Grid1.CellDoubleClick
        If e.RowIndex = -1 Then Return
        idtabl = Grid1all.ElementAt(e.RowIndex).IDПеревоз

        Dim f As New ИзменПеревоз
        f.ShowDialog()
    End Sub

    Private Sub ComboBox1_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles ComboBox1.SelectionChangeCommitted
        Cursor = Cursors.WaitCursor
        CheckBox1.Checked = False
        удал = Nothing
        ОбнКом1()

        Cursor = Cursors.Default
    End Sub

    Private Sub ComboBox2_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles ComboBox2.SelectionChangeCommitted

        CheckBox1.Checked = False
        refreshgrid()

    End Sub

    Private Sub ComboBox3_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles ComboBox3.SelectionChangeCommitted
        CheckBox1.Checked = False
        удал = Nothing
        comb3()
    End Sub

    Private Sub ComboBox3_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox3.KeyDown
        If e.KeyCode = Keys.Enter Then
            CheckBox1.Checked = False
            удал = Nothing
            comb3()
        End If
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked = True Then
            Cursor = Cursors.WaitCursor
            удал = Nothing
            Dim mo As New AllUpd

            Do While AllClass.ПеревозчикиБаза Is Nothing
                mo.ПеревозчикиБазаAll()
            Loop

            If Grid1all IsNot Nothing Then
                Grid1all.Clear()
            End If

            ComboBox1.Text = String.Empty
            ComboBox2.Text = String.Empty

            Dim f = (From x In AllClass.ПеревозчикиБаза
                     Order By x.Наименование_фирмы
                     Where x?.ADR?.Length > 0).ToList()
            If f IsNot Nothing Then
                If f.Count > 0 Then
                    Dim i As Integer = 1
                    For Each b In f
                        Dim b1 As New Grid1Class With {.ADR = b.ADR, .IDПеревоз = b.ID, .КоличАвто = b.Кол_во_авто, .Контакт = b.Контактное_лицо, .Объем = b.Объем,
                       .Организация = b.Наименование_фирмы, .Примечание = b.Примечание, .Регионы = b.Регионы, .Ставка = b.Ставка, .Страны = b.Страны_перевозок,
                       .Телефон = b.Телефоны, .ТипАвто = b.Вид_авто, .Тоннаж = b.Тоннаж, .Номер = i}

                        Grid1all.Add(b1)
                        i += 1
                    Next

                End If


            End If
            Cursor = Cursors.Default

        End If
    End Sub
End Class