Option Explicit On
Imports System.Data.OleDb
Imports System.Threading
Public Class РабочаяФорма
    Dim tbl As New DataTable
    Dim tbl2 As New DataTable
    Dim strsql, strsql2, strsql3 As String
    Private Delegate Sub comb38(ByVal ds As DataGridView)
    Private Delegate Sub com4()
    Private Delegate Sub com41(ByVal ds As DataGridView)




    Private Sub Подборка()

        Dim strsql As String = "SELECT Примечание FROM ИтогГрузПеревоз WHERE IDПеревоз=" & IDПеревоза & " AND IDГруз=" & IDГруз & ""
        Dim ds As DataTable = Selects3(strsql)
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
        Dim ds3 As DataTable = Selects3(StrSql:="SELECT Примечание, Дата FROM ИтогГрузПеревоз WHERE IDПеревоз=" & IDПеревоза & " AND IDГруз=" & IDГруз & "")


        Me.ListBox1.Items.Clear()
        If errds = 0 Then

            For Each r As DataRow In ds3.Rows
                Me.ListBox1.Items.Add(r(0).ToString)
            Next
            Me.ListBox1.Items.Add("Дата переговоров - " & ds3.Rows(0).Item(1).ToString)
        End If

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
    Private Sub Загр1()
        If RichTextBox1.InvokeRequired Then
            Me.Invoke(New com4(AddressOf Загр1))
        Else
            strsql = "SELECT Организация, Дата, Груз, СтранаЗагрузки as [Страна загр], СтранаВыгрузки as [Страна выгр], ГородЗагрузки as [Город загр],
ГородВыгрузки as [Город выгр], Ставка, ОрганизКонтакт, Код
From ГрузыКлиентов Where Код=" & IDГруз & ""
            Dim tbl As DataTable = Selects3(strsql)
            RichTextBox1.Text = tbl.Rows(0).Item(8).ToString
            Grid1.DataSource = tbl
            'Grid1.Columns(0).FillWeight = 70
            'Grid1.Columns(1).FillWeight = 50
            Grid1.Columns(2).FillWeight = 750
            Grid1.Columns(8).Visible = False
            Grid1.Columns(9).Visible = False
            Grid1.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        End If
    End Sub
    Private Sub РабочаяФорма_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.MdiParent = MDIParent1
        Me.WindowState = FormWindowState.Maximized
        'conn = New OleDbConnection
        'conn.ConnectionString = ConString
        'Try
        '    conn.Open()
        'Catch ex As Exception
        '    MessageBox.Show("Не подключен диск U")
        'End Try
        Dim k As New Thread(AddressOf Загр1)
        k.IsBackground = True
        'k.SetApartmentState(ApartmentState.STA)
        k.Start()



        'Dim k1 As New Thread(AddressOf Grid2Refresh)
        'k1.IsBackground = True
        ''k1.SetApartmentState(ApartmentState.STA)
        'k1.Start(0)

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
        Dim ds As DataTable = Selects3(strsql)
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

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        Grid2.ClearSelection()
        For Each r As DataGridViewRow In Grid2.Rows
            For Each cell As DataGridViewCell In r.Cells
                If (cell.FormattedValue).Contains(Me.ComboBox1.Text) Then
                    r.Selected = True
                    Grid2.FirstDisplayedScrollingRowIndex = r.Index
                End If
            Next
        Next
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        ДанныеДляВставкиСкайпа = New Tuple(Of String, String, String, String, Integer)(Grid1.Rows(0).Cells(0).Value, Grid1.Rows(0).Cells(6).Value, Grid1.Rows(0).Cells(2).Value, Grid1.Rows(0).Cells(3).Value, Grid1.Rows(0).Cells(9).Value)
        ДляСкайпаРабочаяФорма.ShowDialog()

    End Sub

    Public Sub Grid2Refresh(ByVal d As Integer)
        Dim df As String
        If d = 0 Then
            df = Выборка.СтрПоиска
        Else
            df = Grid1.Rows(0).Cells(3).Value
        End If
        'If Not IsDBNull(tbl2) Then
        '    tbl2.Clear()
        'End If

        'Select Case[Наименование фирмы], [Контактное лицо], Телефоны,[Страны перевозок],Регионы, ADR,[Кол-во авто], Вид_авто, Тоннаж, Объем, Ставка,
        'Примечание, ID

        Dim strsql2 As String = "SELECT [Наименование фирмы],[Контактное лицо],Телефоны,[Страны перевозок],Регионы, ADR,[Кол-во авто],Вид_авто,Тоннаж,Объем,Ставка,Примечание,ID
FROM ПеревозчикиБаза
Where [Страны перевозок] LIKE '%" & df & "%' 
ORDER BY [Наименование фирмы]"
        Dim tbl2 As DataTable = Selects3(StrSql:=strsql2)

        'If d = 0 Then
        '    tbl2.Columns.Add("Переговоры", Type.GetType("System.String"))
        'End If


        tbl2.Columns.Add("Переговоры", Type.GetType("System.String"))

        For x As Integer = 0 To tbl2.Rows.Count - 1
            Dim f As Integer = tbl2.Rows(x).Item(7).ToString.Length
            If f = 0 Then Continue For
            tbl2.Rows(x).Item(7) = Strings.UCase(Strings.Left(tbl2.Rows(x).Item(7).ToString, 1)) & Strings.Right(tbl2.Rows(x).Item(7).ToString, f - 1)
        Next

        Dim dst As DataTable = Selects3(StrSql:="SELECT DISTINCT Вид_авто From ПеревозчикиБаза Where [Страны перевозок] LIKE '%" & df & "%'")

        For y As Integer = 0 To dst.Rows.Count - 1
            Dim f As Integer = dst.Rows(y).Item(0).ToString.Length
            If f = 0 Then Continue For
            dst.Rows(y).Item(0) = Strings.UCase(Strings.Left(dst.Rows(y).Item(0).ToString, 1)) & Strings.Right(dst.Rows(y).Item(0).ToString, f - 1)
        Next

        Dim row As DataRow = dst.NewRow
        row("Вид_авто") = "Все авто"
        dst.Rows.Add(row)


        ComboBox2.AutoCompleteCustomSource.Clear()
        ComboBox2.Items.Clear()
        For i As Integer = 0 To dst.Rows.Count - 1
            Me.ComboBox2.AutoCompleteCustomSource.Add(dst.Rows(i).Item(0).ToString)
            Me.ComboBox2.Items.Add(dst.Rows(i).Item(0).ToString)
        Next




        Отбор()
        grid2start(tbl2)



        Dim g As New Thread(AddressOf com1)
            g.IsBackground = True
            g.Start(Grid2)

    End Sub
    Private Sub grid2start(ByVal ds As DataTable)


        Grid2.DataSource = ds
        Dim f As New Thread(Sub() grid2selected(Grid2))
        f.IsBackground = True
        f.Start()

        'Grid2.Columns(12).Visible = False
        Grid2.Columns(0).FillWeight = 200
            Grid2.Columns(1).FillWeight = 200
            Grid2.Columns(2).FillWeight = 200
            Grid2.Columns(3).FillWeight = 200
            Grid2.Columns(11).FillWeight = 200
            Grid2.Columns(13).FillWeight = 250
            Grid2.SelectionMode = DataGridViewSelectionMode.FullRowSelect

        'Grid2.ClearSelection() 'Поиск в Grid1
        'For Each row As DataGridViewRow In Grid2.Rows
        '    For Each cell As DataGridViewCell In row.Cells
        '        If (cell.FormattedValue).Contains(IDПеревоза) Then
        '            row.Selected = True
        '            Grid2.FirstDisplayedScrollingRowIndex = row.Index
        '        End If
        '    Next
        'Next



    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        If ComboBox2.Text = "Все авто" Then
            Отбор()
            grid2start(tbl2)
            Exit Sub
        End If

        Dim ds As New DataTable
        ds = tbl2.Copy()

        For f As Integer = 0 To ds.Rows.Count - 1
            If Not ds.Rows(f).Item(7).ToString.Contains(ComboBox2.Text) Then
                ds.Rows(f).Delete()
            End If
        Next
        grid2start(ds)


    End Sub

    Private Sub grid2selected(ByVal ds As DataGridView)
        If Grid2.InvokeRequired Then
            Me.Invoke(New com41(AddressOf grid2selected), ds)
        Else
            Grid2.ClearSelection()
            For Each row As DataGridViewRow In Grid2.Rows
                For Each cell As DataGridViewCell In row.Cells
                    If (cell.FormattedValue).Contains(IDПеревоза) Then
                        row.Selected = True
                        Grid2.FirstDisplayedScrollingRowIndex = row.Index
                    End If
                Next
            Next
        End If

    End Sub
    Private Sub com1(ByVal ds As DataGridView)

        If ComboBox1.InvokeRequired Then
            Me.Invoke(New comb38(AddressOf com1), ds)
        Else
            ComboBox1.AutoCompleteCustomSource.Clear()
            ComboBox1.Items.Clear()
            For i As Integer = 0 To ds.Rows.Count - 1
                Me.ComboBox1.AutoCompleteCustomSource.Add(ds.Rows(i).Cells(0).Value.ToString)
                Me.ComboBox1.Items.Add(ds.Rows(i).Cells(0).Value.ToString)
            Next
        End If




    End Sub


End Class