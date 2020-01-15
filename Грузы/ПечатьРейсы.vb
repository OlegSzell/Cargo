Option Explicit On
Imports System.Data.OleDb

Public Class ПечатьРейсы
    Dim strsql As String
    Dim ds As DataTable
    Dim Клиент, Перевозчик As String
    Dim file2(), file3() As String

    Private Sub ПечатьРейсы_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        conn = New OleDbConnection
        conn.ConnectionString = ConString
        Try
            conn.Open()
        Catch ex As Exception
            MessageBox.Show("Не подключен диск U")
        End Try
        ComboBox1.Items.Clear()
        Dim d() As String = {Now.Year, Now.Year - 1, Now.Year - 2, Now.Year - 3, Now.Year - 4}
        ComboBox1.Items.AddRange(d)

        Dim d4() As String = {"", "1", "2", "3"}
        ComboBox3.Items.Clear()
        ComboBox4.Items.Clear()
        ComboBox5.Items.Clear()
        ComboBox6.Items.Clear()
        ComboBox7.Items.Clear()
        ComboBox8.Items.Clear()

        ComboBox3.Items.AddRange(d4)
        ComboBox4.Items.AddRange(d4)
        ComboBox5.Items.AddRange(d4)
        ComboBox6.Items.AddRange(d4)
        ComboBox7.Items.AddRange(d4)
        ComboBox8.Items.AddRange(d4)



    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        ComboBox2.Items.Clear()
        Справки(ComboBox1.SelectedItem)
        ComboBox2.Items.AddRange(Files3)
        ComboBox2.Text = Files3.Last
        Me.ComboBox2.AutoCompleteCustomSource.Clear()
        ComboBox4.AutoCompleteCustomSource.AddRange(Files3)
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked = True Then
            ComboBox3.Text = "1"
        Else
            ComboBox3.Text = ""
        End If
    End Sub

    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox2.CheckedChanged
        If CheckBox2.Checked = True Then
            ComboBox4.Text = "1"
        Else
            ComboBox4.Text = ""
        End If
    End Sub

    Private Sub CheckBox3_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox3.CheckedChanged
        If CheckBox3.Checked = True Then
            ComboBox5.Text = "1"
        Else
            ComboBox5.Text = ""
        End If
    End Sub

    Private Sub CheckBox5_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox5.CheckedChanged
        If CheckBox5.Checked = True Then
            ComboBox7.Text = "1"
        Else
            ComboBox7.Text = ""
        End If
    End Sub

    Private Sub CheckBox8_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox7.CheckedChanged
        If CheckBox7.Checked = True Then
            ComboBox6.Text = "1"
        Else
            ComboBox6.Text = ""
        End If
    End Sub

    Private Sub CheckBox6_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox6.CheckedChanged
        If CheckBox6.Checked = True Then
            ComboBox8.Text = "1"
        Else
            ComboBox8.Text = ""
        End If
    End Sub
    Private Function проверка()
        If ComboBox1.Text = "" Then
            MessageBox.Show("Выберите год!", Рик)
            Return 1
        End If

        If ComboBox2.Text = "" Then
            MessageBox.Show("Выберите рейс!", Рик)
            Return 1
        End If

        If CheckBox1.Checked = False And CheckBox2.Checked = False And CheckBox3.Checked = False And CheckBox4.Checked = False And CheckBox5.Checked = False And CheckBox6.Checked = False And CheckBox7.Checked = False Then
            MessageBox.Show("Выберите документ для печати!", Рик)
            Return 1
        End If

        Return 0
    End Function
    Private Sub ДоговорЗак()
        Dim f As String = IO.Path.GetExtension(file3(0))
        If f = ".doc" Then
            ПечатьДоков(file3(0), CType(ComboBox7.Text, Integer))
        Else
            ПечатьДоковЭксель(file3(0), CType(ComboBox7.Text, Integer))
        End If
    End Sub
    Private Sub ДоговорПер()
        Dim f As String = IO.Path.GetExtension(file2(0))
        If f = ".doc" Then
            ПечатьДоков(file2(0), CType(ComboBox8.Text, Integer))
        Else
            ПечатьДоковЭксель(file2(0), CType(ComboBox8.Text, Integer))
        End If
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Cursor = Cursors.WaitCursor


        If проверка() = 1 Then Exit Sub

        If CheckBox1.Checked = True Or CheckBox2.Checked = True Or CheckBox3.Checked = True Or CheckBox4.Checked = True Or CheckBox7.Checked = True Then
            Dim xlapp As Microsoft.Office.Interop.Excel.Application
            Dim xlworkbook As Microsoft.Office.Interop.Excel.Workbook
            Dim xlworksheet As Microsoft.Office.Interop.Excel.Worksheet
            'Dim misvalue As Object = Reflection.Missing.Value
            xlapp = New Microsoft.Office.Interop.Excel.Application With {
                .Visible = False
            }
            'xlworkbook = xlapp.Workbooks.Add(misvalue)
            xlworkbook = xlapp.Workbooks.Open("Z:\RICKMANS\" & ComboBox1.Text & "\" & ComboBox2.Text,, True)


            If CheckBox1.Checked = True Then
                xlworksheet = xlworkbook.Sheets("ЗАК")
                xlworksheet.PrintOutEx(,, CType(ComboBox3.Text, Integer))
            End If

            If CheckBox2.Checked = True Then
                xlworksheet = xlworkbook.Sheets("АКТ")
                xlworksheet.PrintOutEx(,, CType(ComboBox4.Text, Integer))
            End If

            If CheckBox3.Checked = True Then
                xlworksheet = xlworkbook.Sheets("СФ")
                xlworksheet.PrintOutEx(,, CType(ComboBox5.Text, Integer))
            End If


            If CheckBox4.Checked = True Then
                xlworksheet = xlworkbook.Sheets("РЕЙС")
                xlworksheet.PrintOutEx()
            End If

            If CheckBox7.Checked = True Then
                xlworksheet = xlworkbook.Sheets("ПЕР")
                xlworksheet.PrintOutEx(,, CType(ComboBox6.Text, Integer))
            End If



            xlworkbook.Close(False)
            xlapp.Quit()

            releaseobject(xlapp)
            releaseobject(xlworkbook)
            releaseobject(xlworksheet)
        End If


        If CheckBox5.Checked = True Then
            ДоговорЗак()

        End If

        If CheckBox6.Checked = True Then
            ДоговорПер()
        End If

        MessageBox.Show("Документы распечатаны!", Рик)
        Me.Cursor = Cursors.Default
    End Sub
    Private Sub refreshlist()
        Dim strsql As String
        Try
            strsql = "SELECT НазвОрганизации FROM РейсыПеревозчика WHERE НомерРейса=" & CType(Strings.Left(ComboBox2.Text, 3), Integer) & ""
        Catch ex As Exception
            Exit Sub
        End Try

        Dim ds As DataTable = Selects(strsql)

        'Dim FilesList() As String = IO.Directory.GetFiles("Z:\RICKMANS\", "*" & ds.Rows(0).Item(0).ToString & "*", IO.SearchOption.AllDirectories)
        Dim FilesList() As String = IO.Directory.GetFiles("Z:\RICKMANS\", "*" & "ДОГОВОР" & "*" & ds.Rows(0).Item(0).ToString & "*", IO.SearchOption.AllDirectories)
        Dim gth4 As String




        For n As Integer = 0 To FilesList.Length - 1
            gth4 = ""
            gth4 = IO.Path.GetFileName(FilesList(n))
            FilesList(n) = gth4
            'TextBox44.Text &= gth + vbCrLf
        Next

        'ListBox2.Items.Add(Files2)


        ListBox2.Items.Clear()

        For i = 0 To FilesList.Length - 1 ' Распечатываем весь получившийся массив
            ListBox2.Items.Add(FilesList(i)) ' На ListBox2
        Next

        file2 = IO.Directory.GetFiles("Z:\RICKMANS\", "*" & "ДОГОВОР" & "*" & ds.Rows(0).Item(0).ToString & "*", IO.SearchOption.AllDirectories)

        грид1()



    End Sub
    Private Sub грид1()
        Dim strsql As String
        Try
            strsql = "SELECT НазвОрганизации FROM РейсыКлиента WHERE НомерРейса=" & CType(Strings.Left(ComboBox2.Text, 3), Integer) & ""
        Catch ex As Exception
            Exit Sub
        End Try

        Dim ds As DataTable = Selects(strsql)

        'Dim FilesList() As String = IO.Directory.GetFiles("Z:\RICKMANS\", "*" & ds.Rows(0).Item(0).ToString & "*", IO.SearchOption.AllDirectories)
        Dim FilesList() As String = IO.Directory.GetFiles("Z:\RICKMANS\", "*" & "ДОГОВОР" & "*" & ds.Rows(0).Item(0).ToString & "*", IO.SearchOption.AllDirectories)
        Dim gth4 As String




        For n As Integer = 0 To FilesList.Length - 1
            gth4 = ""
            gth4 = IO.Path.GetFileName(FilesList(n))
            FilesList(n) = gth4
            'TextBox44.Text &= gth + vbCrLf
        Next

        'ListBox2.Items.Add(Files2)


        ListBox1.Items.Clear()

        For i = 0 To FilesList.Length - 1 ' Распечатываем весь получившийся массив
            ListBox1.Items.Add(FilesList(i)) ' На ListBox2
        Next

        file3 = IO.Directory.GetFiles("Z:\RICKMANS\", "*" & "ДОГОВОР" & "*" & ds.Rows(0).Item(0).ToString & "*", IO.SearchOption.AllDirectories)
    End Sub
    Private Sub ListBox2_DoubleClick(sender As Object, e As EventArgs) Handles ListBox2.DoubleClick

        If ListBox2.SelectedIndex = -1 Then
            MessageBox.Show("Выберите документ для просмотра!", Рик, MessageBoxButtons.OK)
            Exit Sub
        End If

        Dim ff As ListBox.SelectedIndexCollection = ListBox2.SelectedIndices


        If Not ListBox2.SelectedIndex = -1 Then

            For Each p As Integer In ff
                Process.Start(file2(p))
            Next

        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        CheckBox1.Checked = False
        CheckBox2.Checked = False
        CheckBox3.Checked = False
        CheckBox4.Checked = False
        CheckBox5.Checked = False
        CheckBox6.Checked = False
        CheckBox7.Checked = False


    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()
    End Sub

    Private Sub PrintDocument1_PrintPage(sender As Object, e As Printing.PrintPageEventArgs) Handles PrintDocument1.PrintPage

    End Sub

    Private Sub ListBox1_DoubleClick(sender As Object, e As EventArgs) Handles ListBox1.DoubleClick

        If ListBox1.SelectedIndex = -1 Then
            MessageBox.Show("Выберите документ для просмотра!", Рик, MessageBoxButtons.OK)
            Exit Sub
        End If

        Dim ff As ListBox.SelectedIndexCollection = ListBox1.SelectedIndices


        If Not ListBox1.SelectedIndex = -1 Then

            For Each p As Integer In ff
                Process.Start(file3(p))
            Next

        End If
    End Sub
    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        refreshlist()


    End Sub
End Class