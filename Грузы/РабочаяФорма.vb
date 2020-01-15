Option Explicit On
Imports System.Data.OleDb

Public Class РабочаяФорма
    Dim tbl As New DataTable
    Dim tbl2 As New DataTable
    Dim strsql, strsql2, strsql3 As String



    Private Sub Подборка()

        Dim strsql As String = "SELECT Примечание FROM ИтогГрузПеревоз WHERE IDПеревоз=" & IDПеревоза & " AND IDГруз=" & IDГруз & ""
        Dim ds As DataTable = Selects(strsql)
        If errds = 1 Then
            Exit Sub
        Else
            РабочаяФормаМодал.TextBox1.Text = ds.Rows(0).Item(0).ToString
            РабочаяФормаМодал.Label1.Text = "1"
        End If

    End Sub

    Private Sub Grid2_CellMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles Grid2.CellMouseClick

        If e.Button = MouseButtons.Right Then
            РабочаяФормаМодал.Label1.Text = "0"
            IDПеревоза = Nothing
            NameПеревоза = ""
            IDПеревоза = Grid2.CurrentRow.Cells("ID").Value
            NameПеревоза = Grid2.CurrentRow.Cells("Наименование фирмы").Value
            Подборка()
            РабочаяФормаМодал.ShowDialog()
        End If

    End Sub

    Private Sub Grid1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles Grid1.CellClick

    End Sub
    Public Sub Клик()
        IDПеревоза = Grid2.CurrentRow.Cells("ID").Value
        Grid2.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        strsql3 = "SELECT Примечание, Дата FROM ИтогГрузПеревоз WHERE IDПеревоз=" & IDПеревоза & " AND IDГруз=" & IDГруз & ""
        Dim c3 As New OleDbCommand With {
            .Connection = conn,
            .CommandText = strsql3
        }
        Dim ds3 As New DataTable
        Dim da3 As New OleDbDataAdapter(c3)
        Try
            da3.Fill(ds3)
            Me.ListBox1.Items.Clear()
            For Each r As DataRow In ds3.Rows
                Me.ListBox1.Items.Add(r(0).ToString)
            Next
            Me.ListBox1.Items.Add("Дата переговоров - " & ds3.Rows(0).Item(1).ToString)

        Catch ex As Exception
            'MessageBox.Show("С перевозчиком не велись переговоры по этому грузу")
        End Try
        'Grid2Refresh(1)
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        ДобПер = 1
        ДобавитьПеревозчика.ShowDialog()
        Grid2Refresh(1)
        ДобПер = 0
    End Sub

    Private Sub Grid2_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles Grid2.CellClick

        Клик()


    End Sub

    Private Sub РабочаяФорма_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.MdiParent = MDIParent1
        Me.WindowState = FormWindowState.Maximized
        conn = New OleDbConnection
        conn.ConnectionString = ConString
        Try
            conn.Open()
        Catch ex As Exception
            MessageBox.Show("Не подключен диск U")
        End Try

        strsql = "SELECT Организация, Дата, Груз, СтранаЗагрузки as [Страна загр], СтранаВыгрузки as [Страна выгр], ГородЗагрузки as [Город загр],
ГородВыгрузки as [Город выгр], Ставка, ОрганизКонтакт
From ГрузыКлиентов Where Код=" & IDГруз & ""

        Dim c As New OleDbCommand
        c.Connection = conn
        c.CommandText = strsql
        'Dim ds As New DataSet
        Dim da As New OleDbDataAdapter(c)
        'da.Fill(ds, "Сотрудники")
        da.Fill(tbl)
        RichTextBox1.Text = tbl.Rows(0).Item(8).ToString
        Grid1.DataSource = tbl
        'Grid1.Columns(0).FillWeight = 70
        'Grid1.Columns(1).FillWeight = 50
        Grid1.Columns(2).FillWeight = 750
        Grid1.Columns(8).Visible = False
        Grid1.SelectionMode = DataGridViewSelectionMode.FullRowSelect

        Grid2Refresh(0)
    End Sub

    Private Sub Grid2_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles Grid2.CellDoubleClick
        pr2 = 1
        idtabl = IDПеревоза
        ИзменПеревоз.ShowDialog()
        Клик()
        'Grid2Refresh(1)
        pr2 = 0
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        ГрузРассылка.ShowDialog()
    End Sub

    Private Sub Отбор()
        Dim strsql As String = "SELECT IDПеревоз, Примечание FROM ИтогГрузПеревоз WHERE IDГруз=" & IDГруз & ""
        Dim ds As DataTable = Selects(strsql)
        If errds = 1 Then
            Exit Sub
        End If

        For j As Integer = 0 To ds.Rows.Count - 1
            For i As Integer = 0 To tbl2.Rows.Count - 1
                If ds.Rows(j).Item(0) = tbl2.Rows(i).Item(12) Then
                    tbl2.Rows(i).Item(13) = ds.Rows(j).Item(1).ToString
                End If
            Next
        Next
    End Sub
    Public Sub Grid2Refresh(ByVal d As Integer)
        Dim df As String
        If d = 0 Then
            df = Выборка.СтрПоиска
        Else
            df = Grid1.Rows(0).Cells(3).Value
        End If

        strsql2 = "SELECT ПеревозчикиБаза.[Наименование фирмы], ПеревозчикиБаза.[Контактное лицо],
        ПеревозчикиБаза.Телефоны, ПеревозчикиБаза.[Страны перевозок], ПеревозчикиБаза.Регионы, ПеревозчикиБаза.ADR,
        ПеревозчикиБаза.[Кол-во авто], ПеревозчикиБаза.Вид_авто, ПеревозчикиБаза.Тоннаж, ПеревозчикиБаза.Объем, ПеревозчикиБаза.Ставка,
        ПеревозчикиБаза.Примечание, ПеревозчикиБаза.ID
        From ПеревозчикиБаза Where ПеревозчикиБаза.[Страны перевозок] LIKE '%" & df & "%' ORDER BY [Наименование фирмы]"


        Dim c1 As New OleDbCommand
        c1.Connection = conn
        c1.CommandText = strsql2
        'Dim ds As New DataSet
        Dim da1 As New OleDbDataAdapter(c1)
        'da.Fill(ds, "Сотрудники")
        tbl2.Clear()
        da1.Fill(tbl2)
        If d = 0 Then
            tbl2.Columns.Add("Переговоры", Type.GetType("System.String"))
        End If
        Отбор()
        Grid2.DataSource = tbl2
        Grid2.Columns(12).Visible = False
        Grid2.Columns(0).FillWeight = 200
        Grid2.Columns(1).FillWeight = 200
        Grid2.Columns(2).FillWeight = 200
        Grid2.Columns(3).FillWeight = 200
        Grid2.Columns(11).FillWeight = 200
        Grid2.Columns(13).FillWeight = 250
        Grid2.SelectionMode = DataGridViewSelectionMode.FullRowSelect

        Grid2.ClearSelection() 'Поиск в Grid1
        For Each row As DataGridViewRow In Grid2.Rows
            For Each cell As DataGridViewCell In row.Cells
                If (cell.FormattedValue).Contains(IDПеревоза) Then
                    row.Selected = True
                    Grid2.FirstDisplayedScrollingRowIndex = row.Index
                End If
            Next
        Next


    End Sub


End Class