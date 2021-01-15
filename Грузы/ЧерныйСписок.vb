Option Explicit On
'Imports System.Data.OleDb

Public Class ЧерныйСписокForm
    Private comb1 As New List(Of String)
    Private bscom1 As New BindingSource
    Private Grid1all As New List(Of ЧерныйСписок)
    Private bsGrid1 As New BindingSource
    Public g As Integer = 0

    Private Async Sub PredzagAsync()
        Await Task.Run(Sub() Predzag())
    End Sub
    Private Sub Predzag()
        Dim mo As New AllUpd
        Do While AllClass.ЧерныйСписок Is Nothing
            mo.ЧерныйСписокAll()
        Loop
    End Sub
    Private Sub ЧерныйСписок_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        PredzagAsync()
        refgrid1()
    End Sub

    Public Sub refgrid1()
        Dim mo As New AllUpd
        Do While AllClass.ЧерныйСписок Is Nothing
            mo.ЧерныйСписокAll()
        Loop
        bscom1.DataSource = comb1
        ComboBox1.DataSource = bscom1

        bsGrid1.DataSource = Grid1all
        Grid1.DataSource = bsGrid1
        GridView(Grid1)
        Grid1.Columns(0).Visible = False
        Grid1.Columns(1).Width = 200


        Dim f = AllClass.ЧерныйСписок.OrderBy(Function(x) x.Организация).Select(Function(x) x.Организация).Distinct().ToList()
        If f IsNot Nothing Then
            If comb1 IsNot Nothing Then
                comb1.Clear()
                comb1.AddRange(f)
                bscom1.ResetBindings(False)
            End If
        End If
        ComboBox1.Text = String.Empty
        'Dim strsql As String = "SELECT DISTINCT Организация FROM ЧерныйСписок"
        'Dim ds As DataTable = Selects3(strsql)

        'Me.ComboBox1.AutoCompleteCustomSource.Clear()
        'Me.ComboBox1.Items.Clear()
        'For Each r As DataRow In ds.Rows
        '    Me.ComboBox1.AutoCompleteCustomSource.Add(r.Item(0).ToString())
        '    Me.ComboBox1.Items.Add(r(0).ToString)
        'Next
    End Sub

    'Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
    '    Dim strsql As String = "SELECT * FROM ЧерныйСписок WHERE Организация='" & ComboBox1.Text & "'"
    '    Dim ds As DataTable = Selects3(strsql)

    '    Grid1.DataSource = ds
    '    Grid1.Columns(0).Visible = False
    '    Grid1.Columns(1).Width = 200
    'End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If TextBox1.Text = "" Then
            MessageBox.Show("Введите цифры для поиска")
            Exit Sub
        End If

        Dim mo As New AllUpd
        Do While AllClass.ЧерныйСписок Is Nothing
            mo.ЧерныйСписокAll()
        Loop
        Dim txt = Trim(TextBox1.Text)

        Dim f = (From x In AllClass.ЧерныйСписок
                 Where x?.Примечание?.Contains(txt)
                 Select x).ToList()

        If Grid1all IsNot Nothing Then
            Grid1all.Clear()
        End If

        If f IsNot Nothing Then
            If f.Count > 0 Then
                Grid1all.AddRange(f)
                bsGrid1.ResetBindings(False)
            Else
                MessageBox.Show("Нет такого номреа в Черном списке!", Рик)
            End If
        Else
            MessageBox.Show("Нет такого номреа в Черном списке!", Рик)
        End If

        'Dim strsql As String = "SELECT * FROM ЧерныйСписок WHERE Примечание Like '%" & TextBox1.Text & "%'"
        'Dim ds As DataTable = Selects3(strsql)

        'If errds = 1 Then
        '    MessageBox.Show("Нет такого номреа в Черном списке!", Рик)
        '    ds.Clear()
        'Else
        '    Grid1.DataSource = ds
        '    Grid1.Columns(0).Visible = False
        '    Grid1.Columns(1).Width = 200
        '    Grid1.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        'End If

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim f As New ДобЧернСписок(False)
        f.ShowDialog()
        Dim mo As New AllUpd
        mo.ЧерныйСписокAll()
    End Sub

    Private Sub Grid1_CellMouseDoubleClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles Grid1.CellMouseDoubleClick
        Dim d As ЧерныйСписок = Grid1all.ElementAt(Grid1.CurrentRow.Index)
        'Dim strsql As String = "SELECT * FROM ЧерныйСписок WHERE Код=" & d & ""
        'Dim ds As DataTable = Selects3(strsql)
        Dim f As New ДобЧернСписок(True, d)
        f.ShowDialog()
        Dim mo As New AllUpd
        mo.ЧерныйСписокAll()
        d.Примечание = f.Rez
        bsGrid1.ResetBindings(False)
    End Sub

    Private Sub ComboBox1_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles ComboBox1.SelectionChangeCommitted
        Dim mo As New AllUpd
        Do While AllClass.ЧерныйСписок Is Nothing
            mo.ЧерныйСписокAll()
        Loop

        Dim m = ComboBox1.SelectedItem


        Dim f = AllClass.ЧерныйСписок.Where(Function(x) x.Организация = m).FirstOrDefault()
        If Grid1all IsNot Nothing Then
            Grid1all.Clear()
        End If

        Grid1all.Add(f)
        bsGrid1.ResetBindings(False)
        'Dim strsql As String = "SELECT * FROM ЧерныйСписок WHERE Организация='" & ComboBox1.Text & "'"
        'Dim ds As DataTable = Selects3(strsql)

        'Grid1.DataSource = ds
        'Grid1.Columns(0).Visible = False
        'Grid1.Columns(1).Width = 200
    End Sub


End Class