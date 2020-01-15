Imports System.Threading
'Imports Microsoft.Office.Interop.Excel
Imports System.Data.OleDb
Public Class ПоискВРейсах
    Private Delegate Sub comb11()
    Private Delegate Sub comb12()
    Dim ГотовыеДанные As New List(Of String)
    Dim СписокПапок() As String
    Dim fl4 As New List(Of String)()
    Dim flPath4 As New List(Of String)()
    Dim FilesПолнПуть2() As String
    Dim Files45() As String

    Private Sub ПоискВРейсах_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Cursor = Cursors.WaitCursor
        'Dim d As New Thread(AddressOf List1)
        'd.IsBackground = True
        'd.SetApartmentState(ApartmentState.STA)
        'd.Start()
        list1sbor()
        Me.Cursor = Cursors.Default
    End Sub
    Private Sub list1sbor()


        Dim gth3 As String
        Dim f As String = Рейс.comdjx1

        FilesПолнПуть2 = (IO.Directory.GetFiles("Z:\RICKMANS\" & Рейс.comdjx1, "*.xls*", IO.SearchOption.TopDirectoryOnly))
        Files45 = (IO.Directory.GetFiles("Z:\RICKMANS\" & Рейс.comdjx1, "*.xls*", IO.SearchOption.TopDirectoryOnly))
        For n As Integer = 0 To Files45.Length - 1
            gth3 = ""
            gth3 = IO.Path.GetFileName(Files45(n))
            Files45(n) = gth3
        Next
        ListBox1.Items.AddRange(Files45)
        Label14.Text = Files45.Length

    End Sub
    Private Sub List1()
        Dim gth As New List(Of Integer)()
        Dim listpath As New List(Of String)()
        Dim listPathName As New List(Of String)()
        Dim gth3 As String
        Dim Files4() As String
        Dim FilesПолнПуть2() As String


        СписокПапок = (IO.Directory.GetDirectories("Z:\RICKMANS", "*", IO.SearchOption.TopDirectoryOnly)) '"*", IO.SearchOption.TopDirectoryOnly

        For m As Integer = 0 To СписокПапок.Length - 1
            listPathName.Add(Strings.Right(IO.Path.GetFullPath(СписокПапок(m)), 4)) 'отбираем из списка папок только года
        Next

        For y As Integer = 0 To listPathName.Count - 1
            If IsNumeric(listPathName(y)) = True Then
                listpath.Add(listPathName(y)) 'собираем только года
            End If
        Next


        For l As Integer = 0 To listpath.Count - 1
            Try
                FilesПолнПуть2 = (IO.Directory.GetFiles("Z:\RICKMANS\" & Рейс.comdjx1, "*.xls*", IO.SearchOption.TopDirectoryOnly))
                For p As Integer = 0 To FilesПолнПуть2.Count - 1
                    flPath4.Add(FilesПолнПуть2(p))
                Next

                'flPath4.AddRange(IO.Directory.GetFiles("Z:\RICKMANS\" & listPathName(l), "*.xls*", IO.SearchOption.TopDirectoryOnly))
                Files4 = (IO.Directory.GetFiles("Z:\RICKMANS\" & Рейс.comdjx1, "*.xls*", IO.SearchOption.TopDirectoryOnly))
                For n As Integer = 0 To Files4.Length - 1
                    gth3 = ""
                    gth3 = IO.Path.GetFileName(Files4(n))
                    fl4.Add(gth3)
                Next

            Catch ex As Exception

            End Try
        Next
        lab14(fl4.Count)
        Listdelegate(fl4)
    End Sub
    Private Sub lab14(ByVal d As Integer)
        If Label14.InvokeRequired Then
            Me.Invoke(New comb12(Sub() lab14(d)))
        Else
            Label14.Text = d
        End If

    End Sub
    Private Sub Listdelegate(ByVal d As List(Of String))
        If ListBox1.InvokeRequired Then
            Me.Invoke(New comb11(Sub() Listdelegate(d)))
        Else
            ListBox1.Items.Clear()
            For m As Integer = 0 To d.Count - 1
                If Not d(m) = Nothing Then
                    ListBox1.Items.Add(d(m))
                End If
            Next
            'ListBox1.Text = Files3.Last
        End If
    End Sub
    Private Sub ПоискВБазе(ByVal d As String, ByVal m As Integer)
        Dim t, p As String
        If m = 1 Then
            t = TextBox1.Text
            p = "РейсыКлиента"
        ElseIf m = 2 Then
            t = TextBox2.Text
            p = "РейсыКлиента"
        ElseIf m = 3 Then
            t = TextBox3.Text
            p = "РейсыКлиента"
        Else
            t = TextBox4.Text
            p = "РейсыПеревозчика"
        End If

        Dim ds As DataTable = Selects3(StrSql:="SELECT * FROM " & p & " WHERE " & d & " LIKE '%" & t & "%'")
        Dim f As Integer = ds.Rows.Count
        ListBox1.Items.Clear()
        For x As Integer = 0 To Files45.Count - 1
            For v As Integer = 0 To ds.Rows.Count - 1
                If Strings.Left(Files45(x).ToString, 3) = ds.Rows(v).Item(2).ToString Then
                    ListBox1.Items.Add(Files45(x))
                End If
            Next
        Next
        Label14.Text = ListBox1.Items.Count
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        TextBox3.Text = ""
        TextBox2.Text = ""
        list1sbor()
        ПоискВБазе("Маршрут", 1)
        ЧистимРичи()
    End Sub
    Private Sub ЧистимРичи()
        For Each tx1 In Controls.OfType(Of GroupBox)
            For Each tx In tx1.Controls.OfType(Of RichTextBox)
                tx.Text = ""
            Next
        Next
        For Each tx1 In Controls.OfType(Of GroupBox)
            For Each tx2 In tx1.Controls.OfType(Of GroupBox)
                For Each tx In tx2.Controls.OfType(Of RichTextBox)
                    tx.Text = ""
                Next
            Next
        Next
    End Sub
    Private Sub ПоискПомаршрутуВариант1()
        Dim _Excel As New Microsoft.Office.Interop.Excel.Application
        Dim Лист As Microsoft.Office.Interop.Excel.Worksheet
        Dim Книга As Microsoft.Office.Interop.Excel.Workbook
        _Excel.Visible = False

        For x As Integer = 0 To flPath4.Count - 1
            Try
                Книга = _Excel.Workbooks.Open(flPath4(x)) 'Открываем книгу
                Лист = CType(Книга.Worksheets("ЗАК"), Microsoft.Office.Interop.Excel.Worksheet)
                If Лист.Range("C7").Value.ToString.Contains(TextBox1.Text) Then
                    ГотовыеДанные.Add(flPath4(x))
                End If
                _Excel.Workbooks(1).Close(False)
                _Excel.Quit()
            Catch ex As Exception

            End Try

        Next

        ' сохранение и закрытие
        _Excel.Quit()

        'Первый лист книги
        TextBox1.Text = Лист.Range("B5").Value
    End Sub
    Private Sub ПоискПоМаршруту()
        Dim xlapp1 As Application
        Dim xlworkbook1 As Microsoft.Office.Interop.Excel.Workbook
        Dim xlworksheet1 As Microsoft.Office.Interop.Excel.Worksheet
        'Dim misvalue As Object = Reflection.Missing.Value





        'For x As Integer = 0 To flPath4.Count - 1
        '    Try
        '        xlapp1 = New Microsoft.Office.Interop.Excel.Application With {
        '    .Visible = False
        '}
        '        xlworkbook1 = xlapp1.Microsoft.Office.Interop.Excel.Workbooks.Open(flPath4(x),, True) 'Открываем книгу
        '        xlworksheet1 = CType(xlapp1.Worksheets("ЗАК"), Worksheet)
        '        If xlworksheet1.Range("C7").Value.ToString.Contains(TextBox1.Text) Then
        '            ГотовыеДанные.Add(flPath4(x))
        '        End If
        '        xlapp1.Quit()
        '        releaseobject(xlapp1)
        '    Catch ex As Exception

        '    End Try

        'Next
        ListBox1.Items.Clear()
        For n As Integer = 0 To ГотовыеДанные.Count - 1
            ListBox1.Items.Add(ГотовыеДанные(n))
        Next
    End Sub
    Private Sub ВесьСписок()
        Dim f As Integer = CType(Strings.Left(ListBox1.SelectedItem.ToString, 3), Integer)
        Dim ds As DataTable = Selects3(StrSql:="SELECT * FROM РейсыКлиента WHERE НомерРейса=" & f & "")
        Dim ds1 As DataTable = Selects3(StrSql:="SELECT * FROM РейсыПеревозчика WHERE НомерРейса=" & f & "")
        Dim ds2 As DataTable = Selects3(StrSql:="SELECT [Контактное лицо],Телефон FROM Клиент WHERE НазваниеОрганизации='" & ds.Rows(0).Item(1).ToString & "'")
        Dim ds3 As DataTable = Selects3(StrSql:="SELECT [Контактное лицо],Телефон FROM Перевозчики WHERE Названиеорганизации='" & ds1.Rows(0).Item(1).ToString & "'")

        RichTextBox1.Text = ds.Rows(0).Item(1).ToString
        RichTextBox2.Text = ds1.Rows(0).Item(1).ToString
        RichTextBox3.Text = ds.Rows(0).Item(4).ToString
        RichTextBox4.Text = ds.Rows(0).Item(5).ToString
        RichTextBox5.Text = ds.Rows(0).Item(9).ToString
        RichTextBox6.Text = ds.Rows(0).Item(12).ToString
        RichTextBox7.Text = ds.Rows(0).Item(14).ToString
        RichTextBox8.Text = ds.Rows(0).Item(7).ToString
        RichTextBox9.Text = ds.Rows(0).Item(16).ToString
        RichTextBox10.Text = ds2.Rows(0).Item(0).ToString & " (" & ds2.Rows(0).Item(1).ToString & ")"
        RichTextBox11.Text = ds3.Rows(0).Item(0).ToString & " (" & ds3.Rows(0).Item(1).ToString & ")"
        RichTextBox12.Text = ds.Rows(0).Item(17).ToString
        RichTextBox13.Text = ds1.Rows(0).Item(17).ToString
    End Sub

    Private Sub ListBox1_MouseClick(sender As Object, e As MouseEventArgs) Handles ListBox1.MouseClick
        For Each tx1 In Controls.OfType(Of GroupBox)
            For Each tx In tx1.Controls.OfType(Of RichTextBox)
                tx.Text = ""
            Next
        Next

        If TextBox1.Text = "" And TextBox2.Text = "" And TextBox3.Text = "" And TextBox4.Text = "" Then
            ВесьСписок()
        End If


        Dim f As Integer = CType(Strings.Left(ListBox1.SelectedItem.ToString, 3), Integer)
        Dim ds As DataTable = Selects3(StrSql:="SELECT * FROM РейсыКлиента WHERE НомерРейса=" & f & "")
        Dim ds1 As DataTable = Selects3(StrSql:="SELECT * FROM РейсыПеревозчика WHERE НомерРейса=" & f & "")
        Dim ds2 As DataTable = Selects3(StrSql:="SELECT [Контактное лицо],Телефон FROM Клиент WHERE НазваниеОрганизации='" & ds.Rows(0).Item(1).ToString & "'")
        Dim ds3 As DataTable = Selects3(StrSql:="SELECT [Контактное лицо],Телефон FROM Перевозчики WHERE Названиеорганизации='" & ds1.Rows(0).Item(1).ToString & "'")

        RichTextBox1.Text = ds.Rows(0).Item(1).ToString
        RichTextBox2.Text = ds1.Rows(0).Item(1).ToString
        RichTextBox3.Text = ds.Rows(0).Item(4).ToString
        RichTextBox4.Text = ds.Rows(0).Item(5).ToString
        RichTextBox5.Text = ds.Rows(0).Item(9).ToString
        RichTextBox6.Text = ds.Rows(0).Item(12).ToString
        RichTextBox7.Text = ds.Rows(0).Item(14).ToString
        RichTextBox8.Text = ds.Rows(0).Item(7).ToString
        RichTextBox9.Text = ds.Rows(0).Item(16).ToString
        RichTextBox10.Text = ds2.Rows(0).Item(0).ToString & " (" & ds2.Rows(0).Item(1).ToString & ")"
        RichTextBox11.Text = ds3.Rows(0).Item(0).ToString & " (" & ds3.Rows(0).Item(1).ToString & ")"
        RichTextBox12.Text = ds.Rows(0).Item(17).ToString
        RichTextBox13.Text = ds1.Rows(0).Item(17).ToString
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        TextBox1.Text = ""
        TextBox2.Text = ""

        list1sbor()
        ПоискВБазе("ТочныйАдресЗагрузки", 3)
        ЧистимРичи()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        TextBox1.Text = ""
        TextBox3.Text = ""

        list1sbor()
        ПоискВБазе("НазвОрганизации", 2)
        ЧистимРичи()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        TextBox1.Text = ""
        TextBox3.Text = ""
        TextBox2.Text = ""
        list1sbor()
        ПоискВБазе("НазвОрганизации", 4)
        ЧистимРичи()
    End Sub

    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown
        If e.KeyCode = Keys.Enter Then
            Button8.PerformClick()
        End If
    End Sub

    Private Sub TextBox3_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox3.KeyDown
        If e.KeyCode = Keys.Enter Then
            Button2.PerformClick()
        End If
    End Sub

    Private Sub TextBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox2.KeyDown
        If e.KeyCode = Keys.Enter Then
            Button1.PerformClick()
        End If
    End Sub

    Private Sub TextBox4_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox4.KeyDown
        If e.KeyCode = Keys.Enter Then
            Button3.PerformClick()
        End If
    End Sub
End Class