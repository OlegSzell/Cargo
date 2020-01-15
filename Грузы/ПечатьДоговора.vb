Option Explicit On
Imports System.Data.OleDb
Public Class ПечатьДоговора
    Dim strsql, strsql1 As String
    Dim ds, ds1 As DataTable
    Dim Клиент, Перевозчик As String
    Dim FilesList() As String, FilesList1() As String

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        refreshList2()
    End Sub

    Private Sub ListBox1_DoubleClick(sender As Object, e As EventArgs) Handles ListBox1.DoubleClick
        If ListBox1.SelectedIndex = -1 Then
            MessageBox.Show("Выберите документ для просмотра!", Рик, MessageBoxButtons.OK)
            Exit Sub
        End If
        ОткрытиеФайловБезПути(ListBox1.SelectedItem)

    End Sub

    Private Sub ListBox2_DoubleClick(sender As Object, e As EventArgs) Handles ListBox2.DoubleClick
        If ListBox2.SelectedIndex = -1 Then
            MessageBox.Show("Выберите документ для просмотра!", Рик, MessageBoxButtons.OK)
            Exit Sub
        End If
        ОткрытиеФайловБезПути(ListBox2.SelectedItem)


    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If ListBox1.SelectedIndex = -1 Then
            MessageBox.Show("Выберите документ для печати!", Рик, MessageBoxButtons.OK)
            Exit Sub
        End If

        Dim currentdirectory = "Z:\RICKMANS"
        Dim filename = ListBox1.SelectedItem
        Dim path = System.IO.Directory.GetFiles(currentdirectory, "*" & filename, IO.SearchOption.AllDirectories)(0)

        If IO.Path.GetExtension(ListBox1.SelectedItem) = ".doc" Then
            If MessageBox.Show("Распечатать " & ListBox1.SelectedItem & "?", Рик, MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
                ПечатьДоков(path, CType(ComboBox3.Text, Integer))
            End If
        End If

        If IO.Path.GetExtension(ListBox1.SelectedItem) = ".xls" Or IO.Path.GetExtension(ListBox1.SelectedItem) = ".xlsm" Then
            If MessageBox.Show("Распечатать " & ListBox1.SelectedItem & "?", Рик, MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
                ПечатьДоковЭксель(path, CType(ComboBox3.Text, Integer))
            End If
        End If


        'IO.Path.GetPathRoot(Application.ExecutablePath)

        'Dim ff As ListBox.SelectedIndexCollection = ListBox1.SelectedIndices

        'For Each ListBox1.SelectedItem In FilesList

        'Next
        'For i As Integer = 0 To ListBox1.Items.Count - 1
        '    If FilesList.Contains(ListBox1.Items(i)) Then
        '        Dim d As String = Array.Find(FilesList, ListBox1.SelectedItem)
        '    End If
        'Next

        'If Not ListBox1.SelectedIndex = -1 Then
        '    For Each p As Integer In ff
        '        If IO.Path.GetExtension(FilesList1(p)) = ".doc*" Then
        '            ПечатьДоков(FilesList1(p), CType(ComboBox3.Text, Integer))
        '        End If
        '    Next
        'End If

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If ListBox2.SelectedIndex = -1 Then
            MessageBox.Show("Выберите документ для печати!", Рик, MessageBoxButtons.OK)
            Exit Sub
        End If

        Dim currentdirectory = "Z:\RICKMANS"
        Dim filename = ListBox2.SelectedItem
        Dim path = System.IO.Directory.GetFiles(currentdirectory, "*" & filename, IO.SearchOption.AllDirectories)(0)

        If IO.Path.GetExtension(ListBox2.SelectedItem) = ".doc" Then
            If MessageBox.Show("Распечатать договор с " & ListBox2.SelectedItem & "?", Рик, MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
                ПечатьДоков(path, CType(ComboBox4.Text, Integer))
            End If
        End If

        If IO.Path.GetExtension(ListBox2.SelectedItem) = ".xls" Or IO.Path.GetExtension(ListBox2.SelectedItem) = ".xlsm" Then
            If MessageBox.Show("Распечатать " & ListBox1.SelectedItem & "?", Рик, MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
                ПечатьДоковЭксель(path, CType(ComboBox4.Text, Integer))
            End If
        End If
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        refreshList1()
    End Sub


    Private Sub ПечатьДоговора_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        conn = New OleDbConnection
        conn.ConnectionString = ConString
        Try
            conn.Open()
        Catch ex As Exception
            MessageBox.Show("Не подключен диск U")

        End Try


        strsql = "SELECT НазваниеОрганизации FROM Клиент ORDER BY НазваниеОрганизации"
        ds = Selects(strsql)

        Me.ComboBox1.AutoCompleteCustomSource.Clear()
        Me.ComboBox1.Items.Clear()
        For Each r As DataRow In ds.Rows
            Me.ComboBox1.AutoCompleteCustomSource.Add(r.Item(0).ToString())
            Me.ComboBox1.Items.Add(r(0).ToString)
        Next

        strsql1 = "SELECT Названиеорганизации FROM Перевозчики ORDER BY Названиеорганизации"
        ds1 = Selects(strsql1)

        Me.ComboBox2.AutoCompleteCustomSource.Clear()
        Me.ComboBox2.Items.Clear()
        For Each r As DataRow In ds1.Rows
            Me.ComboBox2.AutoCompleteCustomSource.Add(r.Item(0).ToString())
            Me.ComboBox2.Items.Add(r(0).ToString)
        Next
        ComboBox3.Text = "1"
        ComboBox4.Text = "1"

    End Sub
    Private Sub refreshList1()
        Me.Cursor = Cursors.WaitCursor

        FilesList = IO.Directory.GetFiles("Z:\RICKMANS\", "*" & ComboBox1.Text & "*", IO.SearchOption.AllDirectories)
        Dim gth4 As String

        Dim file2() As String = IO.Directory.GetFiles("Z:\RICKMANS\", "*" & ComboBox1.Text & "*", IO.SearchOption.AllDirectories)

        For n As Integer = 0 To FilesList.Length - 1
            gth4 = ""
            gth4 = IO.Path.GetFileName(file2(n))
            file2(n) = gth4
            'TextBox44.Text &= gth + vbCrLf
        Next

        'file2.Reverse

        Array.Sort(file2)
        'Array.Sort(FilesList)

        'ListBox2.Items.Add(Files2)


        ListBox1.Items.Clear()

        For i = 0 To file2.Length - 1 ' Распечатываем весь получившийся массив
            ListBox1.Items.Add(file2(i)) ' На ListBox1
        Next
        Me.Cursor = Cursors.Default
    End Sub
    Private Sub refreshList2()
        Me.Cursor = Cursors.WaitCursor

        FilesList1 = IO.Directory.GetFiles("Z:\RICKMANS\", "*" & ComboBox2.Text & "*", IO.SearchOption.AllDirectories)
        Dim gth4 As String

        Dim file2() As String = IO.Directory.GetFiles("Z:\RICKMANS\", "*" & ComboBox2.Text & "*", IO.SearchOption.AllDirectories)


        For n As Integer = 0 To FilesList1.Length - 1
            gth4 = ""
            gth4 = IO.Path.GetFileName(file2(n))
            file2(n) = gth4
            'TextBox44.Text &= gth + vbCrLf
        Next


        Array.Sort(file2)

        'ListBox2.Items.Add(Files2)


        ListBox2.Items.Clear()

        For i = 0 To file2.Length - 1 ' Распечатываем весь получившийся массив
            ListBox2.Items.Add(file2(i)) ' На ListBox1
        Next
        Me.Cursor = Cursors.Default
    End Sub
End Class