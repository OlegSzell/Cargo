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
    Private Sub ПоискПеревозчиков_Load(sender As Object, e As EventArgs) Handles MyBase.Load
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

    Private Sub Grid1_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles Grid1.CellDoubleClick
        idtabl = Nothing
        idtabl = Grid1.CurrentRow.Cells(12).Value
        ИзменПеревоз.Show()

    End Sub
    Private Sub Grid1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles Grid1.CellClick


        If IsDBNull(e) Then
            Exit Sub
        End If
        удал = Nothing
        удал = Grid1.CurrentRow.Cells(12).Value
    End Sub

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

    Private Sub Grid1_ColumnHeaderMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles Grid1.ColumnHeaderMouseClick
        For Each column As DataGridViewColumn In Grid1.Columns
            column.SortMode = DataGridViewColumnSortMode.NotSortable
        Next
    End Sub

    Private Sub ComboBox1_SelectedValueChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedValueChanged
        ОбнКом1()

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        ДобПер = 1
        ОбнПер = 1
        ДобавитьПеревозчика.WindowState = 0
        ДобавитьПеревозчика.StartPosition = FormStartPosition.CenterScreen
        ДобавитьПеревозчика.ShowDialog()
    End Sub
End Class