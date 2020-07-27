


Imports System.Runtime.CompilerServices

Public Class Календарь
    Private Delegate Sub Grid2Delegate(ByVal f As DataTable, ByVal d As String)
    Private Delegate Sub Grid1Delegate(ByVal d As String)
    Private Property eColumn As Integer
    Private Property eRow As Integer
    Private Property NewColor As Color
    Private Sub Календарь_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.MdiParent = MDIParent1
        NewColor = Color.FromArgb(203, 153, 81)

    End Sub
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
        GridUpdateAsync(e.Start.ToShortDateString())
        Grid2Update(e.Start.ToShortDateString())

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
        If e.ColumnIndex = -1 Then Exit Sub
        If e.ColumnIndex = 0 Then Exit Sub
        If Grid2.CurrentRow.Cells(e.ColumnIndex).Value IsNot Nothing And IsDBNull(Grid2.CurrentRow.Cells(e.ColumnIndex).Value) = False Then
            Dim ml As String = Trim(Grid2.CurrentRow.Cells(e.ColumnIndex).Value)
            'RichTextBox1.Text.Trim(ml)
            'Dim kl = RichTextBox1.Text

            MaskedTextBox1.Text = Strings.Left(Grid2.Columns(e.ColumnIndex).HeaderText, 10)
            ComboBox2.Text = Grid2.CurrentRow.Cells(0).Value

            Label2.Text = Strings.Left(Grid2.Columns(e.ColumnIndex).HeaderText, 10)
            ComboBox1.Text = Grid2.CurrentRow.Cells(0).Value
        End If

    End Sub

    Private Sub Grid2_Scroll(sender As Object, e As ScrollEventArgs) Handles Grid2.Scroll
        GetType(DataGridView).InvokeMember("DoubleBuffered", Reflection.BindingFlags.Instance Or Reflection.BindingFlags.NonPublic Or Reflection.BindingFlags.SetProperty, Nothing, Me.Grid2, New Object() {True})
    End Sub

    Private Sub Grid1_Scroll(sender As Object, e As ScrollEventArgs) Handles Grid1.Scroll
        GetType(DataGridView).InvokeMember("DoubleBuffered", Reflection.BindingFlags.Instance Or Reflection.BindingFlags.NonPublic Or Reflection.BindingFlags.SetProperty, Nothing, Me.Grid1, New Object() {True})
    End Sub




    Private Sub Grid2_CellMouseDown(sender As Object, e As DataGridViewCellMouseEventArgs) Handles Grid2.CellMouseDown
        If e.RowIndex = -1 Then Exit Sub
        If Grid2.Rows(e.RowIndex).Cells(e.ColumnIndex).Value.ToString <> "" Then
            If e.Button = MouseButtons.Right Then
                ContextMenuStrip1.Show(MousePosition, ToolStripDropDownDirection.Right)
                eRow = e.RowIndex
                eColumn = e.ColumnIndex
            End If
        End If

    End Sub

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
End Class