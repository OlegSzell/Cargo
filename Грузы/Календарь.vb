


Imports System.ComponentModel
Imports System.Globalization
Imports System.Reflection
Imports System.Runtime.CompilerServices

Public Class Календарь
    Private Delegate Sub Grid2Delegate(ByVal f As DataTable, ByVal d As String)
    Private Delegate Sub Grid1Delegate(ByVal d As String)
    Private Property eColumn As Integer
    Private Property eRow As Integer
    Private Property NewColor As Color
    Private grid2all As List(Of Grid2Class)
    Private bsgrid2 As BindingSource
    Private SelWeek As String
    Private Sub Календарь_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        NewColor = Color.FromArgb(203, 153, 81)

        grid2all = New List(Of Grid2Class)
        bsgrid2 = New BindingSource
        bsgrid2.DataSource = grid2all
        Grid2.DataSource = bsgrid2
        GridView(Grid2)
        Grid2.Columns(0).Width = 60
        Grid2.SelectionMode = DataGridViewSelectionMode.CellSelect
    End Sub
    Public Class Grid2Class
        Public Property Время As String
        Public Property Понедельник As String
        Public Property Вторник As String
        Public Property Среда As String
        Public Property Четверг As String
        Public Property Пятница As String
        Public Property Суббота As String
        Public Property Воскресенье As String


    End Class
    Private Sub grid2sel(ByVal d As String)
        Dim mo As New AllUpd
        Do While AllClass.Календарь_Даты Is Nothing
            mo.Календарь_ДатыAll()
        Loop

        'номер недели по дате

        'Dim cal As New GregorianCalendar()
        'Dim weekNumber = cal.GetWeekOfYear(d, CalendarWeekRule.FirstFullWeek, DayOfWeek.Monday)

        Dim m As FEndDateWeekClass = FEndDateWeek(d)
        If m Is Nothing Then Return
        If grid2all IsNot Nothing Then
            grid2all.Clear()
        End If
        Dim f1 As New Grid2Class With {.Понедельник = m.FierstDay.ToShortDateString, .Вторник = m.FierstDay.AddDays(1).ToShortDateString,
            .Среда = m.FierstDay.AddDays(2).ToShortDateString, .Четверг = m.FierstDay.AddDays(3).ToShortDateString, .Пятница = m.FierstDay.AddDays(4).ToShortDateString,
            .Суббота = m.FierstDay.AddDays(5).ToShortDateString, .Воскресенье = m.FierstDay.AddDays(6).ToShortDateString}

        grid2all.Add(f1)

        For i As Integer = 0 To 23
            grid2all.Add(New Grid2Class With {.Время = i & ".00"})
        Next
        bsgrid2.ResetBindings(False)
        Dim f = (From x In AllClass.Календарь_Даты
                 Order By x?.Дата
                 Where x?.Дата >= m?.FierstDay And x?.Дата <= m?.EndDay
                 Select x).ToList

        If f IsNot Nothing Then

            ''название столбцов в классе
            'Dim mh1 = f(0).GetType
            'Dim mkl1 As New List(Of Object)
            'For Each b2 In mh.GetFields(BindingFlags.Instance Or BindingFlags.NonPublic)
            '    mkl1.Add(b2.Name)
            'Next
            'mkl1.Remove("_ID")
            'mkl1.Remove("_Дата")
            'mkl1.Remove("_Выполнение")
            'mkl1.Remove("_Неделя")
            'mkl1.Remove("PropertyChangingEvent")
            'mkl1.Remove("PropertyChangedEvent")



            For Each b In f

                Dim dta = CDate(b.Дата).DayOfWeek  'день недели

                'название столбцов в классе
                Dim mh = b.GetType
                Dim mkl As New List(Of Сборка)
                For Each b2 In mh.GetFields(BindingFlags.Instance Or BindingFlags.NonPublic)
                    mkl.Add(New Сборка With {.Время = b2.Name, .Значение = b2?.GetValue(b)?.ToString})
                Next
                mkl.RemoveRange(0, 2)
                mkl.RemoveRange(24, 4)



                SortFunc(dta, mkl, grid2all)
            Next

        End If


    End Sub
    Public Class Сборка
        Public Property Время As String
        Public Property Значение As String
    End Class
    Private Sub SortFunc(ByVal Column As String, ByVal Dann As List(Of Сборка), ByVal p As List(Of Grid2Class))
        Select Case Column 'день недели
            Case 1
                Dim i As Integer = 1
                For Each b1 In Dann
                    p(i).Понедельник = b1.Значение
                    i += 1
                Next

            Case 2
                Dim i As Integer = 1
                For Each b1 In Dann
                    p(i).Вторник = b1.Значение
                    i += 1
                Next
            Case 3
                Dim i As Integer = 1
                For Each b1 In Dann
                    p(i).Среда = b1.Значение
                    i += 1
                Next
            Case 4
                Dim i As Integer = 1
                For Each b1 In Dann
                    p(i).Четверг = b1.Значение
                    i += 1
                Next
            Case 5
                Dim i As Integer = 1
                For Each b1 In Dann
                    p(i).Пятница = b1.Значение
                    i += 1
                Next
            Case 6
                Dim i As Integer = 1
                For Each b1 In Dann
                    p(i).Суббота = b1.Значение
                    i += 1
                Next
            Case 7
                Dim i As Integer = 1
                For Each b1 In Dann
                    p(i).Воскресенье = b1.Значение
                    i += 1
                Next

        End Select
        bsgrid2.ResetBindings(False)
    End Sub
    Private Function FEndDateWeek(ByVal d As String) As FEndDateWeekClass
        Dim f = CDate(d).DayOfWeek
        Dim f1 = CDate(d)
        Select Case f
            Case 1
                Return New FEndDateWeekClass With {.FierstDay = f1, .EndDay = f1.AddDays(6)}
            Case 2
                Return New FEndDateWeekClass With {.FierstDay = f1.AddDays(-1), .EndDay = f1.AddDays(5)}
            Case 3
                Return New FEndDateWeekClass With {.FierstDay = f1.AddDays(-2), .EndDay = f1.AddDays(4)}
            Case 4
                Return New FEndDateWeekClass With {.FierstDay = f1.AddDays(-3), .EndDay = f1.AddDays(3)}
            Case 5
                Return New FEndDateWeekClass With {.FierstDay = f1.AddDays(-4), .EndDay = f1.AddDays(2)}
            Case 6
                Return New FEndDateWeekClass With {.FierstDay = f1.AddDays(-5), .EndDay = f1.AddDays(1)}
            Case 7
                Return New FEndDateWeekClass With {.FierstDay = f1.AddDays(-6), .EndDay = f1}
            Case Else
                Return Nothing
        End Select

    End Function
    Public Class FEndDateWeekClass
        Public Property FierstDay As DateTime
        Public Property EndDay As DateTime
    End Class

    Private Sub GridUpdate(ByVal d As String)

        If Grid1.InvokeRequired Then
            Invoke(New Grid1Delegate(AddressOf GridUpdate), d)
        Else


            Using db As New dbAllDataContext()
                Dim var = db.Календарь_Даты.Where(Function(x) x.Дата >= d).OrderBy(Function(x) x.Дата).Select(Function(x) x).ToList
                'Dim var2 = db.Календарь_Даты.Where(Function(x) x.Дата = d).Select(Function(x) x).FirstOrDefault()
                If var.Count > 0 Then
                    Grid1.DataSource = var
                    Grid1.Columns(2).HeaderText = "0.00"
                    Grid1.Columns(3).HeaderText = "1.00"
                    Grid1.Columns(4).HeaderText = "2.00"
                    Grid1.Columns(5).HeaderText = "3.00"
                    Grid1.Columns(6).HeaderText = "4.00"
                    Grid1.Columns(7).HeaderText = "5.00"
                    Grid1.Columns(8).HeaderText = "6.00"
                    Grid1.Columns(9).HeaderText = "7.00"
                    Grid1.Columns(10).HeaderText = "8.00"
                    Grid1.Columns(11).HeaderText = "9.00"
                    Grid1.Columns(12).HeaderText = "10.00"
                    Grid1.Columns(13).HeaderText = "11.00"
                    Grid1.Columns(14).HeaderText = "12.00"
                    Grid1.Columns(15).HeaderText = "13.00"
                    Grid1.Columns(16).HeaderText = "14.00"
                    Grid1.Columns(17).HeaderText = "15.00"
                    Grid1.Columns(18).HeaderText = "16.00"
                    Grid1.Columns(19).HeaderText = "17.00"
                    Grid1.Columns(20).HeaderText = "18.00"
                    Grid1.Columns(21).HeaderText = "19.00"
                    Grid1.Columns(22).HeaderText = "20.00"
                    Grid1.Columns(23).HeaderText = "21.00"
                    Grid1.Columns(24).HeaderText = "22.00"
                    Grid1.Columns(25).HeaderText = "23.00"



                    GridView(Grid1)
                    Grid1.SelectionMode = DataGridViewSelectionMode.CellSelect
                    ВыделениеСтолбца()
                    'Grid1.AlternatingRowsDefaultCellStyle.BackColor = Color.LightSeaGreen
                Else
                    Grid1.DataSource = Nothing

                End If
            End Using
        End If
    End Sub
    Private Sub ВыделениеСтолбца()
        Dim f As DateTime = DateTime.Now
        Dim g = f.ToString("t")
        If g.Length = 5 Then
            Dim g1 As String = Strings.Left(g, 2)
            Select Case g1
                Case "10"
                    Grid1.Columns(12).DefaultCellStyle.BackColor = Color.MediumVioletRed
                    Grid1.Rows(0).Cells("_10_00").Selected = True
                Case "11"
                    Grid1.Columns(13).DefaultCellStyle.BackColor = Color.MediumVioletRed
                Case "12"
                    Grid1.Columns(14).DefaultCellStyle.BackColor = Color.MediumVioletRed
                Case "13"
                    Grid1.Columns(15).DefaultCellStyle.BackColor = Color.MediumVioletRed
                Case "14"
                    Grid1.Columns(16).DefaultCellStyle.BackColor = Color.MediumVioletRed
                Case "15"
                    Grid1.Columns(17).DefaultCellStyle.BackColor = Color.MediumVioletRed
                Case "16"
                    Grid1.Columns(18).DefaultCellStyle.BackColor = Color.MediumVioletRed
                Case "17"
                    Grid1.Columns(19).DefaultCellStyle.BackColor = Color.MediumVioletRed
                Case "18"
                    Grid1.Columns(20).DefaultCellStyle.BackColor = Color.MediumVioletRed
                Case "19"
                    Grid1.Columns(21).DefaultCellStyle.BackColor = Color.MediumVioletRed
                Case "20"
                    Grid1.Columns(22).DefaultCellStyle.BackColor = Color.MediumVioletRed
                Case "21"
                    Grid1.Columns(23).DefaultCellStyle.BackColor = Color.MediumVioletRed
                Case "22"
                    Grid1.Columns(24).DefaultCellStyle.BackColor = Color.MediumVioletRed
                Case "23"
                    Grid1.Columns(25).DefaultCellStyle.BackColor = Color.MediumVioletRed
            End Select

        ElseIf g.Length = 4 Then

            Dim g2 As String = Strings.Left(g, 1)
            Select Case g2
                Case "0"
                    Grid1.Columns(2).DefaultCellStyle.BackColor = Color.MediumVioletRed
                Case "1"
                    Grid1.Columns(3).DefaultCellStyle.BackColor = Color.MediumVioletRed
                Case "2"
                    Grid1.Columns(4).DefaultCellStyle.BackColor = Color.MediumVioletRed
                Case "3"
                    Grid1.Columns(5).DefaultCellStyle.BackColor = Color.MediumVioletRed
                Case "4"
                    Grid1.Columns(6).DefaultCellStyle.BackColor = Color.MediumVioletRed
                Case "5"
                    Grid1.Columns(7).DefaultCellStyle.BackColor = Color.MediumVioletRed
                Case "6"
                    Grid1.Columns(8).DefaultCellStyle.BackColor = Color.MediumVioletRed
                Case "7"
                    Grid1.Columns(9).DefaultCellStyle.BackColor = Color.MediumVioletRed
                Case "8"
                    Grid1.Columns(10).DefaultCellStyle.BackColor = Color.MediumVioletRed
                Case "9"
                    Grid1.Columns(11).DefaultCellStyle.BackColor = Color.MediumVioletRed

            End Select

        End If
    End Sub

    Private Sub Grid2Update(ByVal d As String)

        Dim dn1 = DatePart("m", CDate(d))
        Dim dn3 = DatePart("yyyy", CDate(d))
        Dim dn2 As Integer = Date.DaysInMonth(DatePart("yyyy", CDate(d)), dn1)

        Dim fdr As Date = CDate("01." & dn1 & "." & dn3)
        Dim fdr1 As Date = fdr.AddDays(dn2 - 1)

        Dim f As New DataTable
        f.Columns.Add("Время")  'вставляем даты
        For i As Integer = 0 To dn2
            f.Columns.Add(fdr.AddDays(i).ToShortDateString & vbCrLf & Format(fdr.AddDays(i), "dddd"))
        Next

        For i2 As Integer = 1 To 24 'вставляем время
            Dim row As DataRow = f.NewRow
            row("Время") = i2 & ".00"
            f.Rows.Add(row)
        Next


        Using db As New dbAllDataContext()
            Dim var = db.Календарь_Даты.Where(Function(x) x.Дата >= fdr And x.Дата <= fdr1).Select(Function(x) x).ToList()
            If var.Count > 0 Then
                For Each col As DataColumn In f.Columns
                    If col.ColumnName = "Время" Then Continue For
                    For Each h In var
                        If Strings.Left(col.ColumnName, 10) = Strings.Left(h.Дата.Value, 10) Then
                            Dim md As String = col.ColumnName
                            f.Rows(23).Item(md) = h._0_00
                            f.Rows(0).Item(md) = h._1_00
                            f.Rows(1).Item(md) = h._2_00
                            f.Rows(2).Item(md) = h._3_00
                            f.Rows(3).Item(md) = h._4_00
                            f.Rows(4).Item(md) = h._5_00
                            f.Rows(5).Item(md) = h._6_00
                            f.Rows(6).Item(md) = h._7_00
                            f.Rows(7).Item(md) = h._8_00
                            f.Rows(8).Item(md) = h._9_00
                            f.Rows(9).Item(md) = h._10_00
                            f.Rows(10).Item(md) = h._11_00
                            f.Rows(11).Item(md) = h._12_00
                            f.Rows(12).Item(md) = h._13_00
                            f.Rows(13).Item(md) = h._14_00
                            f.Rows(14).Item(md) = h._15_00
                            f.Rows(15).Item(md) = h._16_00
                            f.Rows(16).Item(md) = h._17_00
                            f.Rows(17).Item(md) = h._18_00
                            f.Rows(18).Item(md) = h._19_00
                            f.Rows(19).Item(md) = h._20_00
                            f.Rows(20).Item(md) = h._21_00
                            f.Rows(21).Item(md) = h._22_00
                            f.Rows(22).Item(md) = h._23_00
                        End If
                    Next
                Next
            End If
        End Using

        For Each colu As DataColumn In f.Columns
            If colu.ColumnName.Contains(d) = True Then
                d = colu.ColumnName
            End If
        Next



        'Dim d1 As String = f.Columns.Contains(d).ToString
        grid2DelegAsync(f, d)


        'Grid2.DataSource = Nothing
        '    Grid2.DataSource = f
        '    GridView(Grid2)
        '    Grid2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        '    Grid2.ScrollBars = ScrollBars.Both
        '    Grid2.Columns("Время").Frozen = True
        '    Grid2.Columns("Время").DefaultCellStyle.BackColor = Color.LightBlue

        '    Dim vn As DateTime = DateTime.Now
        '    'Dim g = vn.ToString("t")
        '    Dim hl = CType(DatePart("h", vn.ToString("t")), String)
        '    hl &= ".00"


        '    Dim ind As Integer 'строка в цвет по времени
        '    For Each row As DataGridViewRow In Grid2.Rows
        '        If row.Cells(0).Value = hl Then
        '            ind = row.Index
        '            row.DefaultCellStyle.BackColor = Color.LightBlue
        '            row.Cells.Item(d).Style.BackColor = Color.Yellow
        '            row.Cells(0).Style.ForeColor = Color.Red
        '        End If
        '    Next

        '    For Each colum As DataGridViewColumn In Grid2.Columns  'столбец в цвет по сегодняшней дате
        '        If colum.Name = Now.Date.ToShortDateString Then
        '            'colum.DefaultCellStyle.BackColor = Color.LightYellow
        '            colum.HeaderCell.Style.ForeColor = Color.Red
        '            Grid2(colum.Index, ind).Selected = True
        '            'Grid2.CurrentCell = Grid2(colum.Index, ind)
        '        End If
        '    Next

        '    For Each colum As DataGridViewColumn In Grid2.Columns  'столбец в цвет по дате
        '        If colum.Name = d Then
        '            colum.DefaultCellStyle.BackColor = Color.LightGray
        '            Grid2(colum.Index, ind).Selected = True
        '            Grid2.CurrentCell = Grid2(colum.Index, ind)
        '        End If
        '    Next


        'Grid2.SelectionMode = DataGridViewSelectionMode.CellSelect

    End Sub
    Private Async Sub grid2DelegAsync(ByVal f As DataTable, ByVal d As String)
        Await Task.Run(Sub() grid2Deleg(f, d))
    End Sub
    Private Async Sub GridUpdateAsync(ByVal d As String)
        Await Task.Run(Sub() GridUpdate(d))
    End Sub


    Private Sub grid2Deleg(ByVal f As DataTable, ByVal d As String)

        If Grid2.InvokeRequired Then
            Invoke(New Grid2Delegate(AddressOf grid2Deleg), f, d)
        Else
            Grid2.DataSource = Nothing
            Grid2.DataSource = f
            GridView(Grid2)
            Grid2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            Grid2.ScrollBars = ScrollBars.Both
            Grid2.Columns("Время").Frozen = True
            Grid2.Columns("Время").DefaultCellStyle.BackColor = Color.LightBlue

            Dim vn As DateTime = DateTime.Now
            'Dim g = vn.ToString("t")
            Dim hl = CType(DatePart("h", vn.ToString("t")), String)
            hl &= ".00"


            Dim ind As Integer 'строка в цвет по времени
            For Each row As DataGridViewRow In Grid2.Rows
                If row.Cells(0).Value = hl Then
                    ind = row.Index
                    row.DefaultCellStyle.BackColor = Color.LightBlue
                    row.Cells.Item(d).Style.BackColor = Color.Yellow
                    row.Cells(0).Style.ForeColor = Color.Red
                End If
            Next

            For Each colum As DataGridViewColumn In Grid2.Columns  'столбец в цвет по сегодняшней дате

                If Strings.Left(colum.Name, 10) = Now.Date.ToShortDateString Then
                    'colum.DefaultCellStyle.BackColor = Color.LightYellow
                    colum.HeaderCell.Style.ForeColor = Color.Red
                    colum.HeaderCell.Style.Font = New Font("Calibri", 11.0F, FontStyle.Bold)
                    Grid2(colum.Index, ind).Selected = True
                    'Grid2.CurrentCell = Grid2(colum.Index, ind)
                End If

                If colum.Name = d Then
                    colum.DefaultCellStyle.BackColor = Color.LightGray
                    Grid2(colum.Index, ind).Selected = True
                    Grid2.CurrentCell = Grid2(colum.Index, ind)
                End If

            Next

            For Each rw As DataGridViewRow In Grid2.Rows
                For Each cl As DataGridViewColumn In Grid2.Columns
                    If rw.Cells(cl.Index).Value IsNot Nothing Then
                        If rw.Cells(cl.Index).Value.ToString.Contains("(True)") Then
                            Dim lit As Integer = rw.Cells(cl.Index).Value.ToString.Length - 6
                            rw.Cells(cl.Index).Value = Strings.Left(rw.Cells(cl.Index).Value.ToString, lit)
                            rw.Cells(cl.Index).Style.BackColor = NewColor
                            rw.Cells(cl.Index).Style.ForeColor = Color.White
                        End If
                    End If

                Next
            Next


            'For Each coluv As DataGridViewColumn In Grid2.Columns
            '    coluv.SortMode = DataGridViewColumnSortMode.NotSortable
            'Next

            'Grid2.Columns.Cast(Of DataGridViewColumn)().ToList().ForEach(Function(f4) f4.SortMode = DataGridViewColumnSortMode.NotSortable)






            'For Each cl As DataGridViewColumn In Grid2.Columns
            '    cl.Width = 80
            'Next
            'Grid2.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells)
            Grid2.SelectionMode = DataGridViewSelectionMode.CellSelect
        End If

    End Sub

    Private Sub Calendar1_DateSelected(sender As Object, e As DateRangeEventArgs) Handles Calendar1.DateSelected
        Label2.Text = e.Start.ToShortDateString()
        MaskedTextBox1.Text = e.Start.ToShortDateString()

        'Dim myCI As New CultureInfo("ru-RU")
        'Dim myCal As Calendar = myCI.Calendar

        '' Gets the DTFI properties required by GetWeekOfYear.
        'Dim myCWR As CalendarWeekRule = myCI.DateTimeFormat.CalendarWeekRule
        'Dim myFirstDOW As DayOfWeek = myCI.DateTimeFormat.FirstDayOfWeek
        'Dim SelWeek = myCal.GetWeekOfYear(e.Start, myCWR, myFirstDOW)

        grid2sel(e.Start.ToShortDateString())

        'по недели ищем начало недли число
        'Dim firstDay As New DateTime(e.Start.Year, 1, 1)
        'While firstDay.DayOfWeek <> DayOfWeek.Monday
        '    firstDay = firstDay.AddDays(-1)
        'End While
        'Dim k = firstDay.AddDays(7 * (f - 1))

        'GridUpdateAsync(e.Start.ToShortDateString())

        'Grid2Update(e.Start.ToShortDateString())

    End Sub
    Private Sub InsertData()
        Using db As New dbAllDataContext()
            Dim f = db.Календарь_Даты.Where(Function(x) x.Дата = Label2.Text).Select(Function(x) x).FirstOrDefault()
            If f IsNot Nothing Then
                Select Case ComboBox1.Text
                    Case "0.00"
                        f._0_00 = RichTextBox1.Text
                    Case "1.00"
                        f._1_00 = RichTextBox1.Text
                    Case "2.00"
                        f._2_00 = RichTextBox1.Text
                    Case "3.00"
                        f._3_00 = RichTextBox1.Text
                    Case "4.00"
                        f._4_00 = RichTextBox1.Text
                    Case "5.00"
                        f._5_00 = RichTextBox1.Text
                    Case "6.00"
                        f._6_00 = RichTextBox1.Text
                    Case "7.00"
                        f._7_00 = RichTextBox1.Text
                    Case "8.00"
                        f._8_00 = RichTextBox1.Text
                    Case "9.00"
                        f._9_00 = RichTextBox1.Text
                    Case "10.00"
                        f._10_00 = RichTextBox1.Text
                    Case "11.00"
                        f._11_00 = RichTextBox1.Text
                    Case "12.00"
                        f._12_00 = RichTextBox1.Text
                    Case "13.00"
                        f._13_00 = RichTextBox1.Text
                    Case "14.00"
                        f._14_00 = RichTextBox1.Text

                    Case "15.00"
                        f._15_00 = RichTextBox1.Text
                    Case "16.00"
                        f._16_00 = RichTextBox1.Text
                    Case "17.00"
                        f._17_00 = RichTextBox1.Text
                    Case "18.00"
                        f._18_00 = RichTextBox1.Text
                    Case "19.00"
                        f._19_00 = RichTextBox1.Text
                    Case "20.00"
                        f._20_00 = RichTextBox1.Text
                    Case "21.00"
                        f._21_00 = RichTextBox1.Text
                    Case "22.00"
                        f._22_00 = RichTextBox1.Text
                    Case "23.00"
                        f._23_00 = RichTextBox1.Text

                End Select
                db.SubmitChanges()

            End If
        End Using
    End Sub
    Private Sub Deletedata()
        Using db As New dbAllDataContext()
            Dim f = db.Календарь_Даты.Where(Function(x) x.Дата = Label2.Text).Select(Function(x) x).FirstOrDefault()
            If f IsNot Nothing Then
                Select Case ComboBox1.Text
                    Case "0.00"
                        f._0_00 = ""
                    Case "1.00"
                        f._1_00 = ""
                    Case "2.00"
                        f._2_00 = ""
                    Case "3.00"
                        f._3_00 = ""
                    Case "4.00"
                        f._4_00 = ""
                    Case "5.00"
                        f._5_00 = ""
                    Case "6.00"
                        f._6_00 = ""
                    Case "7.00"
                        f._7_00 = ""
                    Case "8.00"
                        f._8_00 = ""
                    Case "9.00"
                        f._9_00 = ""
                    Case "10.00"
                        f._10_00 = ""
                    Case "11.00"
                        f._11_00 = ""
                    Case "12.00"
                        f._12_00 = ""
                    Case "13.00"
                        f._13_00 = ""
                    Case "14.00"
                        f._14_00 = ""

                    Case "15.00"
                        f._15_00 = ""
                    Case "16.00"
                        f._16_00 = ""
                    Case "17.00"
                        f._17_00 = ""
                    Case "18.00"
                        f._18_00 = ""
                    Case "19.00"
                        f._19_00 = ""
                    Case "20.00"
                        f._20_00 = ""
                    Case "21.00"
                        f._21_00 = ""
                    Case "22.00"
                        f._22_00 = ""
                    Case "23.00"
                        f._23_00 = ""

                End Select
                db.SubmitChanges()

            End If
        End Using
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If MessageBox.Show("Сохранить событие?", Рик, MessageBoxButtons.OKCancel, MessageBoxIcon.Question) = DialogResult.Cancel Then
            Exit Sub
        End If

        If Label2.Text = "Дата" Then
            MessageBox.Show("Выберите дату!", Рик)
            Exit Sub
        End If

        If ComboBox1.Text = "" Then
            MessageBox.Show("Выберите время!", Рик)
            Exit Sub
        End If

        Me.Cursor = Cursors.WaitCursor
        Using db As New dbAllDataContext()
            Dim var = db.Календарь_Даты.Where(Function(x) x.Дата = Label2.Text).Select(Function(x) x).FirstOrDefault()
            If var IsNot Nothing Then
                InsertData()
            Else
                Dim f As New Календарь_Даты
                f.Дата = Label2.Text
                db.Календарь_Даты.InsertOnSubmit(f)
                db.SubmitChanges()

                InsertData()
            End If
        End Using


        GridUpdateAsync(Label2.Text)
        Grid2Update(Label2.Text)
        НапоминаниеAdd()

        Dim vbd As New MDIParent1
        vbd.КалендарьНапоминаниеAsync()

        Me.Cursor = Cursors.Default
        очистка()



    End Sub
    Private Sub НапоминаниеAdd()
        If MaskedTextBox1.MaskCompleted = True And ComboBox2.Text <> "" Then

            Using db As New dbAllDataContext()
                Dim f As New КалендарьНапоминание
                f.ДатаНапоминания = MaskedTextBox1.Text
                f.ВремяНапоминания = ComboBox2.Text
                f.ТекстНапоминания = RichTextBox1.Text
                f.Пользователь = Экспедитор
                db.КалендарьНапоминание.InsertOnSubmit(f)
                db.SubmitChanges()
            End Using
        End If
    End Sub
    Private Sub очистка()
        RichTextBox1.Text = ""
        MaskedTextBox1.Text = ""
        ComboBox2.Text = ""
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If MessageBox.Show("Изменить событие?", Рик, MessageBoxButtons.OKCancel, MessageBoxIcon.Question) = DialogResult.Cancel Then
            Exit Sub
        End If

        If Label2.Text = "Дата" Then
            MessageBox.Show("Выберите дату!", Рик)
            Exit Sub
        End If

        If ComboBox1.Text = "" Then
            MessageBox.Show("Выберите время!", Рик)
            Exit Sub
        End If

        Me.Cursor = Cursors.WaitCursor
        InsertData()
        GridUpdateAsync(Label2.Text)
        Grid2Update(Label2.Text)
        НапоминаниеUpd()


        Dim vbd As New MDIParent1 ' запуск напоминания
        vbd.КалендарьНапоминаниеAsync()

        Me.Cursor = Cursors.Default
        очистка()
    End Sub
    Private Sub НапоминаниеUpd()
        Using db As New dbAllDataContext()
            Dim var = db.КалендарьНапоминание.Where(Function(x) x.ДатаНапоминания = CDate(Label2.Text) And x.ВремяНапоминания = ComboBox1.Text And x.Пользователь = Экспедитор).Select(Function(x) x).FirstOrDefault()

            If var IsNot Nothing Then
                var.ВремяНапоминания = ComboBox2.Text
                var.ДатаНапоминания = MaskedTextBox1.Text
                var.ТекстНапоминания = RichTextBox1.Text
                db.SubmitChanges()
            End If
        End Using
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If MessageBox.Show("Удалить событие?", Рик, MessageBoxButtons.OKCancel, MessageBoxIcon.Question) = DialogResult.Cancel Then
            Exit Sub
        End If

        If Label2.Text = "Дата" Then
            MessageBox.Show("Выберите дату!", Рик)
            Exit Sub
        End If

        If ComboBox1.Text = "" Then
            MessageBox.Show("Выберите время!", Рик)
            Exit Sub
        End If


        Me.Cursor = Cursors.WaitCursor
        Deletedata()
        GridUpdateAsync(Label2.Text)
        Grid2Update(Label2.Text)
        Me.Cursor = Cursors.Default
        очистка()
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        If Label2.Text = "Дата" Or Label2.Text = "Время" Then Exit Sub
        Using db As New dbAllDataContext()
            Dim var = db.Календарь_Даты.Where(Function(x) x.Дата = Label2.Text).Select(Function(x) x).FirstOrDefault()
            If var IsNot Nothing Then
                Select Case ComboBox1.Text
                    Case "0.00"
                        RichTextBox1.Text = var._0_00
                    Case "1.00"
                        RichTextBox1.Text = var._1_00
                    Case "2.00"
                        RichTextBox1.Text = var._2_00
                    Case "3.00"
                        RichTextBox1.Text = var._3_00
                    Case "4.00"
                        RichTextBox1.Text = var._4_00
                    Case "5.00"
                        RichTextBox1.Text = var._5_00
                    Case "6.00"
                        RichTextBox1.Text = var._6_00
                    Case "7.00"
                        RichTextBox1.Text = var._7_00
                    Case "8.00"
                        RichTextBox1.Text = var._8_00
                    Case "9.00"
                        RichTextBox1.Text = var._9_00
                    Case "10.00"
                        RichTextBox1.Text = var._10_00
                    Case "11.00"
                        RichTextBox1.Text = var._11_00
                    Case "12.00"
                        RichTextBox1.Text = var._12_00
                    Case "13.00"
                        RichTextBox1.Text = var._13_00
                    Case "14.00"
                        RichTextBox1.Text = var._14_00

                    Case "15.00"
                        RichTextBox1.Text = var._15_00
                    Case "16.00"
                        RichTextBox1.Text = var._16_00
                    Case "17.00"
                        RichTextBox1.Text = var._17_00
                    Case "18.00"
                        RichTextBox1.Text = var._18_00
                    Case "19.00"
                        RichTextBox1.Text = var._19_00
                    Case "20.00"
                        RichTextBox1.Text = var._20_00
                    Case "21.00"
                        RichTextBox1.Text = var._21_00
                    Case "22.00"
                        RichTextBox1.Text = var._22_00
                    Case "23.00"
                        RichTextBox1.Text = var._23_00

                End Select

                If RichTextBox1.Text.Contains("(True)") Then
                    Dim lit As Integer = RichTextBox1.Text.Length - 6
                    RichTextBox1.Text = Strings.Left(RichTextBox1.Text, lit)
                End If
            Else
                RichTextBox1.Text = ""

            End If
            ComboBox2.Text = ComboBox1.Text
        End Using
    End Sub


    Private Sub Grid1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles Grid1.CellContentClick

        'RichTextBox1.Text = Grid1.CurrentRow.Cells(e.ColumnIndex).Value
    End Sub

    Private Sub Grid1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles Grid1.CellClick
        If e.ColumnIndex = -1 Then Exit Sub
        RichTextBox1.Text = Grid1.CurrentRow.Cells(e.ColumnIndex).Value
        MaskedTextBox1.Text = Grid1.CurrentRow.Cells(1).Value
        ComboBox2.Text = Grid1.Columns(e.ColumnIndex).HeaderText
        Label2.Text = Grid1.CurrentRow.Cells(1).Value
        ComboBox1.Text = Grid1.Columns(e.ColumnIndex).HeaderText



    End Sub

    Private Sub Grid2_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles Grid2.CellClick
        'If e.ColumnIndex = -1 Then Exit Sub
        'If e.ColumnIndex = 0 Then Exit Sub
        'If Grid2.CurrentRow.Cells(e.ColumnIndex).Value IsNot Nothing And IsDBNull(Grid2.CurrentRow.Cells(e.ColumnIndex).Value) = False Then
        '    Dim ml As String = Trim(Grid2.CurrentRow.Cells(e.ColumnIndex).Value)
        '    'RichTextBox1.Text.Trim(ml)
        '    'Dim kl = RichTextBox1.Text

        '    MaskedTextBox1.Text = Strings.Left(Grid2.Columns(e.ColumnIndex).HeaderText, 10)
        '    ComboBox2.Text = Grid2.CurrentRow.Cells(0).Value

        '    Label2.Text = Strings.Left(Grid2.Columns(e.ColumnIndex).HeaderText, 10)
        '    ComboBox1.Text = Grid2.CurrentRow.Cells(0).Value
        'End If

    End Sub

    Private Sub Grid2_Scroll(sender As Object, e As ScrollEventArgs) Handles Grid2.Scroll
        GetType(DataGridView).InvokeMember("DoubleBuffered", Reflection.BindingFlags.Instance Or Reflection.BindingFlags.NonPublic Or Reflection.BindingFlags.SetProperty, Nothing, Me.Grid2, New Object() {True})
    End Sub

    Private Sub Grid1_Scroll(sender As Object, e As ScrollEventArgs) Handles Grid1.Scroll 'скрол без бликов  datagrid grid1
        GetType(DataGridView).InvokeMember("DoubleBuffered", Reflection.BindingFlags.Instance Or Reflection.BindingFlags.NonPublic Or Reflection.BindingFlags.SetProperty, Nothing, Me.Grid1, New Object() {True})
    End Sub




    'Private Sub Grid2_CellMouseDown(sender As Object, e As DataGridViewCellMouseEventArgs) Handles Grid2.CellMouseDown
    '    If e.RowIndex = -1 Then Exit Sub
    '    If Grid2.Rows(e.RowIndex).Cells(e.ColumnIndex).Value.ToString <> "" Then
    '        If e.Button = MouseButtons.Right Then
    '            ContextMenuStrip1.Show(MousePosition, ToolStripDropDownDirection.Right)
    '            eRow = e.RowIndex
    '            eColumn = e.ColumnIndex
    '        End If
    '    End If

    'End Sub

    Private Sub ВыполненоToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ВыполненоToolStripMenuItem.Click




        If Grid2.Rows(eRow).Cells(eColumn).Style.BackColor = NewColor Then
            Exit Sub
        Else
            Grid2.Rows(eRow).Cells(eColumn).Style.BackColor = NewColor
            Grid2.Rows(eRow).Cells(eColumn).Style.ForeColor = Color.White
            Using db As New dbAllDataContext
                Dim dat As Date = CDate(Strings.Left(Grid2.Columns(eColumn).HeaderText, 10))
                Dim tim As String = Grid2.Rows(eRow).Cells(0).Value.ToString
                Dim var = db.Календарь_Даты.Where(Function(x) x.Дата = dat).Select(Function(x) x).FirstOrDefault()
                Dim sd As String = "(True)"
                If var IsNot Nothing Then

                    Select Case tim
                        Case "1.00"
                            var._1_00 &= sd
                        Case "2.00"
                            var._2_00 &= sd
                        Case "3.00"
                            var._3_00 &= sd
                        Case "4.00"
                            var._4_00 &= sd
                        Case "5.00"
                            var._5_00 &= sd
                        Case "6.00"
                            var._6_00 &= sd
                        Case "7.00"
                            var._7_00 &= sd
                        Case "8.00"
                            var._8_00 &= sd
                        Case "9.00"
                            var._9_00 &= sd
                        Case "10.00"
                            var._10_00 &= sd
                        Case "11.00"
                            var._11_00 &= sd
                        Case "12.00"
                            var._12_00 &= sd
                        Case "13.00"
                            var._13_00 &= sd
                        Case "14.00"
                            var._14_00 &= sd
                        Case "15.00"
                            var._15_00 &= sd
                        Case "16.00"
                            var._16_00 &= sd
                        Case "17.00"
                            var._17_00 &= sd
                        Case "18.00"
                            var._18_00 &= sd
                        Case "19.00"
                            var._19_00 &= sd
                        Case "20.00"
                            var._20_00 &= sd
                        Case "21.00"
                            var._21_00 &= sd
                        Case "22.00"
                            var._22_00 &= sd
                        Case "23.00"
                            var._23_00 &= sd
                        Case "24.00"
                            var._0_00 &= sd

                    End Select

                    db.SubmitChanges()

                End If

            End Using
        End If

    End Sub
    Private Sub AddDatabase(ByVal dat As String, ByVal int As Integer, ByVal Valu As String)
        Dim dat1 As New Date
        If dat IsNot Nothing Then
            dat1 = CDate(dat)
        End If
        Dim mo As New AllUpd
        Using db As New dbAllDataContext
            Dim f = db.Календарь_Даты.Where(Function(x) CDate(x.Дата) = dat1).FirstOrDefault()
            If f IsNot Nothing Then
                Select Case int
                    Case 0
                        f._0_00 = Valu
                    Case 1
                        f._1_00 = Valu
                    Case 2
                        f._2_00 = Valu
                    Case 3
                        f._3_00 = Valu
                    Case 4
                        f._4_00 = Valu
                    Case 5
                        f._5_00 = Valu
                    Case 6
                        f._6_00 = Valu
                    Case 7
                        f._7_00 = Valu
                    Case 8
                        f._8_00 = Valu
                    Case 9
                        f._9_00 = Valu
                    Case 10
                        f._10_00 = Valu
                    Case 11
                        f._11_00 = Valu
                    Case 12
                        f._12_00 = Valu
                    Case 13
                        f._13_00 = Valu
                    Case 14
                        f._14_00 = Valu
                    Case 15
                        f._15_00 = Valu
                    Case 16
                        f._16_00 = Valu
                    Case 17
                        f._17_00 = Valu
                    Case 18
                        f._18_00 = Valu
                    Case 19
                        f._19_00 = Valu
                    Case 20
                        f._20_00 = Valu
                    Case 21
                        f._21_00 = Valu
                    Case 22
                        f._22_00 = Valu
                    Case 23
                        f._23_00 = Valu
                End Select
                db.SubmitChanges()

            Else
                Dim f1 As New Календарь_Даты

                Select Case int
                    Case 0
                        f1._0_00 = Valu
                    Case 1
                        f1._1_00 = Valu
                    Case 2
                        f1._2_00 = Valu
                    Case 3
                        f1._3_00 = Valu
                    Case 4
                        f1._4_00 = Valu
                    Case 5
                        f1._5_00 = Valu
                    Case 6
                        f1._6_00 = Valu
                    Case 7
                        f1._7_00 = Valu
                    Case 8
                        f1._8_00 = Valu
                    Case 9
                        f1._9_00 = Valu
                    Case 10
                        f1._10_00 = Valu
                    Case 11
                        f1._11_00 = Valu
                    Case 12
                        f1._12_00 = Valu
                    Case 13
                        f1._13_00 = Valu
                    Case 14
                        f1._14_00 = Valu
                    Case 15
                        f1._15_00 = Valu
                    Case 16
                        f1._16_00 = Valu
                    Case 17
                        f1._17_00 = Valu
                    Case 18
                        f1._18_00 = Valu
                    Case 19
                        f1._19_00 = Valu
                    Case 20
                        f1._20_00 = Valu
                    Case 21
                        f1._21_00 = Valu
                    Case 22
                        f1._22_00 = Valu
                    Case 23
                        f1._23_00 = Valu
                End Select
                f1.Дата = dat1.ToShortDateString
                db.Календарь_Даты.InsertOnSubmit(f1)
                db.SubmitChanges()

            End If
            mo.Календарь_ДатыAll()
        End Using
    End Sub



    Private Sub Grid2_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles Grid2.CellEndEdit
        If e.RowIndex = -1 Then Return
        If e.RowIndex = 0 Then Return

        Dim f = e.ColumnIndex
        Dim dat = selAdd(f)
        Dim f3 = grid2all.ElementAt(e.RowIndex)
        Dim f1 = Значение2(f, f3)
        AddDatabase(dat, e.RowIndex - 1, f1)
    End Sub
    Private Function selAdd(ByVal NumbColumn As Integer) As String
        Select Case NumbColumn
            Case 1
                Dim m = grid2all.Select(Function(x) x.Понедельник).FirstOrDefault()
                Return m
            Case 2
                Dim m = grid2all.Select(Function(x) x.Вторник).FirstOrDefault()
                Return m
            Case 3
                Dim m = grid2all.Select(Function(x) x.Среда).FirstOrDefault()
                Return m
            Case 4
                Dim m = grid2all.Select(Function(x) x.Четверг).FirstOrDefault()
                Return m
            Case 5
                Dim m = grid2all.Select(Function(x) x.Пятница).FirstOrDefault()
                Return m
            Case 6
                Dim m = grid2all.Select(Function(x) x.Суббота).FirstOrDefault()
                Return m
            Case 7
                Dim m = grid2all.Select(Function(x) x.Воскресенье).FirstOrDefault()
                Return m
        End Select
        Return Nothing
    End Function
    Private Function Значение2(ByVal NumbColumn As Integer, ByVal Grd2 As Grid2Class) As String
        Select Case NumbColumn
            Case 1
                Dim m = Grd2.Понедельник
                Return m
            Case 2
                Dim m = Grd2.Вторник
                Return m
            Case 3
                Dim m = Grd2.Среда
                Return m
            Case 4
                Dim m = Grd2.Четверг
                Return m
            Case 5
                Dim m = Grd2.Пятница
                Return m
            Case 6
                Dim m = Grd2.Суббота
                Return m
            Case 7
                Dim m = Grd2.Воскресенье
                Return m
        End Select
        Return Nothing
    End Function

End Class