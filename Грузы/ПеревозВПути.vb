Option Explicit On
Imports System.Data.OleDb
Public Class ПеревозВПути
    Dim strsql As String
    Dim ds, ds2 As New DataTable
    Private Sub ПеревозВПути_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        запГрид()
        strsql = "SELECT [Наименование фирмы], ID
FROM ПеревозчикиБаза ORDER BY [Наименование фирмы]"
        Dim c As New OleDbCommand With {
            .Connection = conn,
            .CommandText = strsql
        }

        Dim da As New OleDbDataAdapter(c)
        da.Fill(ds)
        Me.ComboBox1.AutoCompleteCustomSource.Clear()
        Me.ComboBox1.Items.Clear()
        Me.ComboBox2.Items.Clear()
        For Each r As DataRow In ds.Rows
            Me.ComboBox1.AutoCompleteCustomSource.Add(r.Item(0).ToString())
            Me.ComboBox1.Items.Add(r(0).ToString)
            Me.ComboBox2.Items.Add(r(1).ToString)
        Next
        Label2.Text = ""
    End Sub
    Public Sub запГрид()
        Dim времянач As String
        Dim d As Date = Now
        d = d.AddDays(15)
        времянач = Format(d, "MM\/dd\/yyyy")
        Dim времякон As String = DateTime.Now.ToString("MM\/dd\/yyyy")


        ds2.Clear()

        If CheckBox1.Checked = True And TextBox1.Text = "" Then

            strsql = "SELECT ПеревозчикиВПути.Перевозчик, ПеревозчикиБаза.[Контактное лицо], ПеревозчикиБаза.Телефоны, ПеревозчикиВПути.ДатаВыгр as [Дата выгрузки], ПеревозчикиВПути.ГдеВыгр as [Место выгрузки],ПеревозчикиВПути.Примечание
FROM ПеревозчикиБаза INNER Join ПеревозчикиВПути On ПеревозчикиБаза.ID = ПеревозчикиВПути.IDПеревозчика ORDER BY ДатаВыгр"
        ElseIf CheckBox1.Checked = False And TextBox1.Text = "" Then
            strsql = "SELECT ПеревозчикиВПути.Перевозчик, ПеревозчикиБаза.[Контактное лицо], ПеревозчикиБаза.Телефоны, ПеревозчикиВПути.ДатаВыгр as [Дата выгрузки], ПеревозчикиВПути.ГдеВыгр as [Место выгрузки],ПеревозчикиВПути.Примечание
FROM ПеревозчикиБаза INNER Join ПеревозчикиВПути On ПеревозчикиБаза.ID = ПеревозчикиВПути.IDПеревозчика
WHERE ДатаВыгр >= #" & времякон & "# And ДатаВыгр <= #" & времянач & "# ORDER BY ДатаВыгр"
        ElseIf CheckBox1.Checked = True And TextBox1.Text <> "" Then
            strsql = "SELECT ПеревозчикиВПути.Перевозчик, ПеревозчикиБаза.[Контактное лицо], ПеревозчикиБаза.Телефоны, ПеревозчикиВПути.ДатаВыгр as [Дата выгрузки], ПеревозчикиВПути.ГдеВыгр as [Место выгрузки],ПеревозчикиВПути.Примечание
FROM ПеревозчикиБаза INNER Join ПеревозчикиВПути On ПеревозчикиБаза.ID = ПеревозчикиВПути.IDПеревозчика 
WHERE ПеревозчикиВПути.ГдеВыгр LIKE '%" & TextBox1.Text & "%' ORDER BY ДатаВыгр"
        ElseIf CheckBox1.Checked = False And TextBox1.Text <> "" Then
            strsql = "SELECT ПеревозчикиВПути.Перевозчик, ПеревозчикиБаза.[Контактное лицо], ПеревозчикиБаза.Телефоны, ПеревозчикиВПути.ДатаВыгр as [Дата выгрузки], ПеревозчикиВПути.ГдеВыгр as [Место выгрузки],ПеревозчикиВПути.Примечание
FROM ПеревозчикиБаза INNER Join ПеревозчикиВПути On ПеревозчикиБаза.ID = ПеревозчикиВПути.IDПеревозчика
WHERE ДатаВыгр Between #" & времянач & "# And #" & времякон & "# AND ПеревозчикиВПути.ГдеВыгр LIKE '%" & TextBox1.Text & "%' ORDER BY ДатаВыгр"


        End If

        Dim c As New OleDbCommand With {
            .Connection = conn,
            .CommandText = strsql
        }

        Dim da As New OleDbDataAdapter(c)
        da.Fill(ds2)
        Grid1.DataSource = ds2

        Grid1.Columns(0).FillWeight = 100
        Grid1.Columns(4).FillWeight = 300
        Grid1.SelectionMode = DataGridViewSelectionMode.FullRowSelect
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked = True Then
            запГрид()
        Else
            запГрид()
        End If
    End Sub
    Private Sub ПеревозВПути_Closed(sender As Object, e As EventArgs) Handles Me.Closed
        CheckBox1.Checked = False
        Me.ComboBox1.AutoCompleteCustomSource.Clear()
        Me.ComboBox1.Items.Clear()
        Me.ComboBox2.Items.Clear()
        Me.ComboBox1.Text = ""
        TextBox1.Text = ""
        ds.Clear()
        ds2.Clear()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If ComboBox1.Text = "" Then
            MessageBox.Show("Выберите перевозчика!")
            Exit Sub
        End If
        IDПеревоза = Nothing
        NameПеревоза = ""
        IDПеревоза = CType(Label2.Text, Integer)
        NameПеревоза = ComboBox1.Text
        ПеревозВПутиДоб.ShowDialog()
        ds2.Clear()
        запГрид()
    End Sub

    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            запГрид()


        End If
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        Label2.Text = ComboBox2.Items.Item(ComboBox1.SelectedIndex)
    End Sub
End Class