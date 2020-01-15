Option Explicit On
Imports System.Data.OleDb
Public Class ПоискПеревозчиков
    Public Da As New OleDbDataAdapter 'Адаптер
    Public Ds As New DataSet 'Пустой набор записей
    Dim tbl As New DataTable
    Dim cb As OleDb.OleDbCommandBuilder
    Dim Рик As String = "ООО Рикманс"
    Dim загр As String
    Dim ind As Boolean = False
    Dim удал As Integer

    Private Sub ПоискПеревозчиков_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.MdiParent = MDIParent1
        Me.WindowState = FormWindowState.Maximized
        conn = New OleDbConnection
        conn.ConnectionString = ConString
        Try
            conn.Open()
        Catch ex As Exception
            MessageBox.Show("Не подключен диск U")
        End Try
        Dim Год As Integer = Year(Now)

        'If Me.Прием_Load = vbTrue Then Form1.Load = False

        TextBox1.Text = DateTime.Now.ToString("dd.MM.yyyy")

        Dim StrSql As String
        StrSql = "SELECT DISTINCT Страна FROM Страна ORDER BY Страна"
        Dim c As New OleDbCommand With {
            .Connection = conn,
            .CommandText = StrSql
        }
        Dim ds As New DataTable
        Dim da As New OleDbDataAdapter(c)
        da.Fill(ds)
        Me.ComboBox1.AutoCompleteCustomSource.Clear()
        Me.ComboBox1.Items.Clear()
        For Each r As DataRow In ds.Rows
            Me.ComboBox1.AutoCompleteCustomSource.Add(r.Item(0).ToString())
            Me.ComboBox1.Items.Add(r(0).ToString)
        Next

        Dim StrSql2 As String
        StrSql2 = "SELECT DISTINCT [Наименование фирмы] FROM ПеревозчикиБаза ORDER BY [Наименование фирмы]"
        Dim c2 As New OleDbCommand With {
            .Connection = conn,
            .CommandText = StrSql2
        }
        Dim ds2 As New DataTable
        Dim da2 As New OleDbDataAdapter(c2)
        da2.Fill(ds2)
        Me.ComboBox3.AutoCompleteCustomSource.Clear()
        Me.ComboBox3.Items.Clear()
        For Each r As DataRow In ds2.Rows
            Me.ComboBox3.AutoCompleteCustomSource.Add(r.Item(0).ToString())
            Me.ComboBox3.Items.Add(r(0).ToString)
        Next







    End Sub
    Public Sub ОбнКом1()
        tbl.Clear()
        ComboBox2.Text = ""

        CheckBox1.Checked = False
        Dim StrSql3 As String
        StrSql3 = "SELECT DISTINCT Регионы
FROM РегионыРоссии INNER JOIN Страна ON РегионыРоссии.Страны = Страна.Код
Where Страна.Страна = '" & ComboBox1.Text & "'"
        Dim ds2 As DataTable = Selects(StrSql3)

        ComboBox2.Items.Clear()
        ComboBox2.AutoCompleteCustomSource.Clear()

        For Each r As DataRow In ds2.Rows
            Me.ComboBox2.AutoCompleteCustomSource.Add(r.Item(0).ToString())
            Me.ComboBox2.Items.Add(r(0).ToString)
        Next

        refreshgrid()
    End Sub
    Private Sub ComboBox1_SelectedValueChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedValueChanged
        ОбнКом1()

        'ComboBox4.Text = ""
    End Sub
    Private Sub refreshgrid()

        tbl.Clear()

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
        'Страна INNER Join ПеревозчикиNew On Страна.Код = ПеревозчикиNew.[Страны перевозок].Value
        Dim c1 As New OleDbCommand
        c1.Connection = conn
        c1.CommandText = StrSql2
        'Dim ds As New DataSet
        Dim da1 As New OleDbDataAdapter(c1)
        'da.Fill(ds, "Сотрудники")
        da1.Fill(tbl)

        Grid1.DataSource = tbl
        'Grid1.Item(0, 0).Style.Font = New Font("Calibri", 12)
        'cb = New OleDb.OleDbCommandBuilder(da1)
        Grid1.Columns(12).Visible = False
        Grid1.SelectionMode = DataGridViewSelectionMode.FullRowSelect

        Grid1.Columns(0).Width = 250
        Grid1.Columns(1).Width = 250
        Grid1.Columns(2).Width = 150
        Grid1.Columns(11).Width = 320
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

            tbl = Selects(StrSql2)

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

        'Страна INNER Join ПеревозчикиNew On Страна.Код = ПеревозчикиNew.[Страны перевозок].Value
        Dim c6 As New OleDbCommand
        c6.Connection = conn
        c6.CommandText = StrSql6
        'Dim ds As New DataSet
        Dim da6 As New OleDbDataAdapter(c6)
        'da.Fill(ds, "Сотрудники")
        da6.Fill(tbl)
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
            StrSql6 = "DELETE * FROM ПеревозчикиБаза Where ID =" & удал & ""
            Dim c6 As New OleDbCommand
            c6.Connection = conn
            c6.CommandText = StrSql6
            c6.ExecuteNonQuery()
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
End Class