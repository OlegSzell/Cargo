Option Explicit On
Imports System.Data.OleDb
Imports System.Threading

Public Class ПечатьРейсы
    Dim strsql As String
    Dim ds As DataTable
    Dim Клиент, Перевозчик As String
    Dim file2(), file3() As String
    Dim FilesList() As String
    'Private Declare Function MessageBoxTimeOut Lib "user32" Alias "MessageBoxTimeoutA" (ByVal hWnd As Integer, ByVal lpText As String,
    '                  ByVal uType As MsgBoxStyle, ByVal wLanguageId As Integer, ByVal dwMilliseconds As Integer) As Integer
    Private Declare Function MessageBoxTimeOut Lib "user32" Alias "MessageBoxTimeoutA" (ByVal hWnd As Integer, ByVal lpText As String, ByVal lpCaption As String,
                      ByVal uType As MsgBoxStyle, ByVal wLanguageId As Integer, ByVal dwMilliseconds As Integer) As Integer
    'Private Delegate Sub listbx1(ByVal fillist As List(Of String), ByVal ib As Integer, ByVal НомерРейса As String)
    'Private Delegate Sub listbx2(ByVal fillist As List(Of String), ByVal ib As Integer, ByVal НомерРейса As String)



    Private Sub ПечатьРейсы_Load(sender As Object, e As EventArgs) Handles MyBase.Load


        ComboBox1.Items.Clear()
        Dim d() As String = {Now.Year, Now.Year - 1, Now.Year - 2, Now.Year - 3, Now.Year - 4}
        d.Reverse
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

        Dim hg As New Thread(AddressOf договораобщая)
        hg.IsBackground = True
        hg.Start()

    End Sub
    Private Sub договораобщая()
        FilesList = IO.Directory.GetFiles("Z:\RICKMANS\", "*" & "ДОГОВОР" & "*", IO.SearchOption.AllDirectories)
        Dim gth4 As String

        For n As Integer = 0 To FilesList.Length - 1
            gth4 = IO.Path.GetFileName(FilesList(n))
            FilesList(n) = gth4
            'TextBox44.Text &= gth + vbCrLf
        Next
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        ComboBox2.Items.Clear()
        Справки(ComboBox1.SelectedItem)
        ComboBox2.Items.AddRange(Files3)
        ComboBox2.Text = Files3.Last
        Me.ComboBox2.AutoCompleteCustomSource.Clear()
        ComboBox2.AutoCompleteCustomSource.AddRange(Files3)
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

        ПечДоки()

        Me.Cursor = Cursors.Default
    End Sub

    Private Sub ПечДоки()
        'MessageBoxTimeOut(Me.Handle, "Документы отправлены на принтер!", MsgBoxStyle.Information, 0&, 1500)
        MessageBoxTimeOut(Me.Handle, "Документы отправлены на принтер!", "Печать документа.", MsgBoxStyle.Information, 0&, 3000)
    End Sub
    Private Function numberRe(ByVal tu As Integer) As String
        Dim f() As String = {"РейсыПеревозчика", "РейсыКлиента"}
        Dim dr As New Thread(AddressOf Сбор)
        dr.IsBackground = True
        Dim dict As New Dictionary(Of String, String)


        Dim strsql As String
        Try
            strsql = "SELECT НазвОрганизации FROM " & f(tu) & " WHERE НомерРейса=" & CType(Strings.Left(ComboBox2.Text, 3), Integer) & ""
            Dim ds As DataTable = Selects3(strsql)
            dict.Add(f(tu), ds.Rows(0).Item(0).ToString)
            dr.Start(dict)
            Return ds.Rows(0).Item(0).ToString
        Catch ex As Exception
            Return 0
        End Try

    End Function
    Private Sub Сбор(ByVal Ном As Dictionary(Of String, String))
        'Dim mas() = {file3, file2}



        If Not Ном Is Nothing Then
            If Ном.ContainsKey("РейсыПеревозчика") Then
                file2 = IO.Directory.GetFiles("Z:\RICKMANS\", "*" & "ДОГОВОР" & "*" & Ном("РейсыПеревозчика") & "*", IO.SearchOption.AllDirectories)
            Else
                file3 = IO.Directory.GetFiles("Z:\RICKMANS\", "*" & "ДОГОВОР" & "*" & Ном("РейсыКлиента") & "*", IO.SearchOption.AllDirectories)
            End If


        End If


    End Sub
    'Private Sub refreshlist(ByVal НомерРейса As String)

    '    If НомерРейса = "" Then
    '        MessageBox.Show("Не найден рейс!")
    '        Exit Sub
    '    End If

    '    'Dim FilesList() As String = IO.Directory.GetFiles("Z:\RICKMANS\", "*" & "ДОГОВОР" & "*" & НомерРейса & "*", IO.SearchOption.AllDirectories)
    '    'Dim gth4 As String

    '    'For n As Integer = 0 To FilesList.Length - 1
    '    '    gth4 = ""
    '    '    gth4 = IO.Path.GetFileName(FilesList(n))
    '    '    FilesList(n) = gth4
    '    'Next

    '    Dim fileslistus As New List(Of String)

    '    For Each r In FilesList
    '        If r.Contains(НомерРейса) Then
    '            fileslistus.Add(r.ToString)
    '        End If
    '    Next

    '    ListBDounload(fileslistus, 0, НомерРейса)

    'End Sub
    Private Sub грид1(ByVal НомерРейса As String, ByVal j As Integer)


        'Dim FilesList() As String = IO.Directory.GetFiles("Z:\RICKMANS\", "*" & ds.Rows(0).Item(0).ToString & "*", IO.SearchOption.AllDirectories)
        'Dim FilesList() As String = IO.Directory.GetFiles("Z:\RICKMANS\", "*" & "ДОГОВОР" & "*" & НомерРейса & "*", IO.SearchOption.AllDirectories)
        'Dim gth4 As String

        'For n As Integer = 0 To FilesList.Length - 1
        '    gth4 = ""
        '    gth4 = IO.Path.GetFileName(FilesList(n))
        '    FilesList(n) = gth4
        '    'TextBox44.Text &= gth + vbCrLf
        'Next
        'Dim fileslistus As New List(Of String)
        If j = 1 Then
            ListBox1.Items.Clear()
            For Each r In FilesList
                If r.Contains(НомерРейса) Then
                    'fileslistus.Add(r.ToString)
                    ListBox1.Items.Add(r.ToString)
                End If
            Next

        Else
            ListBox2.Items.Clear()
            For Each r In FilesList
                If r.Contains(НомерРейса) Then
                    'fileslistus.Add(r.ToString)
                    ListBox2.Items.Add(r.ToString)
                End If
            Next
            'file2 = IO.Directory.GetFiles("Z:\RICKMANS\", "*" & "ДОГОВОР" & "*" & НомерРейса & "*", IO.SearchOption.AllDirectories)
        End If


        'ListBox1.Items.Clear()
        '' Распечатываем весь получившийся массив
        'ListBox1.Items.Add(fillist) ' На ListBox2



        'ListBox2.Items.Add(Files2)

        'ListBDounload(fileslistus, 1, НомерРейса)



    End Sub
    'Private Sub ListBDounload(ByVal fillist As List(Of String), ByVal ib As Integer, ByVal НомерРейса As String)

    '    If ib = 1 Then
    '        If ListBox1.InvokeRequired Then
    '            Me.Invoke(New listbx1(AddressOf ListBDounload), fillist, ib, НомерРейса)
    '        Else
    '            ListBox1.Items.Clear()
    '            ' Распечатываем весь получившийся массив
    '            ListBox1.Items.Add(fillist) ' На ListBox2
    '            file3 = IO.Directory.GetFiles("Z:\RICKMANS\", "*" & "ДОГОВОР" & "*" & НомерРейса & "*", IO.SearchOption.AllDirectories)
    '        End If
    '    Else
    '        If ListBox2.InvokeRequired Then
    '            Me.Invoke(New listbx2(AddressOf ListBDounload), fillist, ib, НомерРейса)
    '        Else

    '            ListBox2.Items.Clear()
    '            ' Распечатываем весь получившийся массив
    '            ListBox2.Items.Add(fillist) ' На ListBox2


    '            file2 = IO.Directory.GetFiles("Z:\RICKMANS\", "*" & "ДОГОВОР" & "*" & НомерРейса & "*", IO.SearchOption.AllDirectories)
    '        End If

    '    End If
    'End Sub
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
        'Dim d As New Thread(AddressOf refreshlist)
        'Dim d1 As New Thread(AddressOf грид1)
        'd.IsBackground = True
        'd1.IsBackground = True
        'd.Start(numberRe(0))
        'd1.Start(numberRe(1))
        грид1(numberRe(1), 1)
        грид1(numberRe(0), 0)

    End Sub
End Class