Option Explicit On
Imports System.Data.OleDb
Imports System.Threading
Imports System.Linq
Public Class Сводная_по_рейсам
    Dim dsper As DataTable
    Dim fr As New Thread(AddressOf Перевоз)
    Dim dc As New Dictionary(Of Integer, Integer)
    Dim dp As New Dictionary(Of Integer, Integer)
    Dim dmax As New Dictionary(Of Integer, Integer)
    Dim dsall As DataTable
    Dim dp3, dp4 As New Dictionary(Of Integer, Integer)
    Dim толькоДобавлНомераРейсов As DataTable
    Dim lts As New List(Of Integer)()



    Private Sub Сводная_по_рейсам_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.MdiParent = MDIParent1

        If fall.IsAlive Then
            fall.Join()
            Запуск(dsall)
            Exit Sub
        ElseIf Not dsall Is Nothing Then
            Запуск(dsall)
            Exit Sub
        End If
        'Ref()
        If dsall Is Nothing Then
            ALL()
            Запуск(dsall)
        End If

    End Sub
    Private Sub Ref()
        Dim StrSql As String
        StrSql = "SELECT DISTINCT НазвОрганизации FROM РейсыКлиента ORDER BY НазвОрганизации"
        Dim ds As DataTable = Selects3(StrSql)

        'Me.ComboBox1.AutoCompleteCustomSource.Clear()
        'Me.ComboBox1.Items.Clear()
        'For Each r As DataRow In ds.Rows
        '    Me.ComboBox1.AutoCompleteCustomSource.Add(r.Item(0).ToString())
        '    Me.ComboBox1.Items.Add(r(0).ToString)
        'Next


        Dim StrSql1 As String
        StrSql1 = "SELECT DISTINCT НазвОрганизации FROM РейсыПеревозчика ORDER BY НазвОрганизации"
        Dim ds1 As DataTable = Selects3(StrSql)

        'Me.ComboBox2.Items.Clear()
        'Me.ComboBox2.AutoCompleteCustomSource.Clear()
        'For Each r As DataRow In ds1.Rows
        '    Me.ComboBox2.AutoCompleteCustomSource.Add(r.Item(0).ToString())
        '    Me.ComboBox2.Items.Add(r(0).ToString)
        'Next
        'ComboBox2.Text = ""
        'ComboBox1.Text = ""

    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        'fr.IsBackground = True
        'fr.Start()




    End Sub
    Public Sub ALL()

        Dim strsql As String = "SELECT НомерРейса as [Рейс], НазвОрганизации as [Клиент], Маршрут, СтоимостьФрахта as [Фрахт],
Валюта as [Валюта], ДатаОплаты as [Срок оплаты] FROM РейсыКлиента WHERE НомерРейса IS NOT NULL" 'клиент
        'Dim ds As DataTable = Selects3(strsql)
        Dim ds As DataTable = Selects3(strsql)

        ds.Columns.Add("Дата оплаты")
        ds.Columns.Add("Cумма")
        ds.Columns.Add("Перевозчик") '8
        ds.Columns.Add("Фрахт пер")
        ds.Columns.Add("Валюта пер")
        ds.Columns.Add("Срок оплаты пер")
        ds.Columns.Add("Дата оплаты пер")
        ds.Columns.Add("Cумма пер")


        Dim strsql5 As String = "SELECT НомерРейса as [Рейс], НазвОрганизации as [Организация], СтоимостьФрахта as [Фрахт],
Валюта as [Валюта], ДатаОплаты as [Срок оплаты] FROM РейсыПеревозчика WHERE НомерРейса IS NOT NULL" 'перевозчик
        'Dim ds5 As DataTable = Selects3(strsql5)
        Dim ds5 As DataTable = Selects3(strsql5)

        ds5.Columns.Add("Дата оплаты")
        ds5.Columns.Add("Cумма")

        For i As Integer = 0 To ds.Rows.Count - 1
            If Not IsDBNull(ds.Rows(i).Item(5)) Then
                ds.Rows(i).Item(5) = Strings.Left(ds.Rows(i).Item(5), 10)
            End If
            For y As Integer = 0 To ds5.Rows.Count - 1
                If ds.Rows(i).Item(0) = ds5.Rows(y).Item(0) Then
                    ds.Rows(i).Item(8) = ds5.Rows(y).Item(1)
                    ds.Rows(i).Item(9) = ds5.Rows(y).Item(2)
                    ds.Rows(i).Item(10) = ds5.Rows(y).Item(3)
                    If Not IsDBNull(ds5.Rows(y).Item(4)) Then
                        ds.Rows(i).Item(11) = Strings.Left(ds5.Rows(y).Item(4), 10)
                    End If
                    ds.Rows(i).Item(12) = ds5.Rows(y).Item(5)
                    ds.Rows(i).Item(13) = ds5.Rows(y).Item(6)
                End If
            Next
        Next
        'ds.DefaultView.Sort = "Рейс" & " ASC" 'сортировка datatable


        Dim strsql6 As String = "SELECT * FROM ОплатыКлиент"
        'Dim ds6 As DataTable = Selects3(strsql6)
        Dim ds6 As DataTable = Selects3(strsql6)

        Dim strsql7 As String = "SELECT * FROM ОплатыПер"
        'Dim ds7 As DataTable = Selects3(strsql7)
        Dim ds7 As DataTable = Selects3(strsql7)

        Dim ds8 As New DataTable
        With ds8.Columns
            .Add("Рейс")
            .Add("Клиент")
            .Add("Маршрут")
            .Add("Фрахт")
            .Add("Валюта")
            .Add("Срок оплаты")
            .Add("Дата оплаты")
            .Add("Cумма")
            .Add("Перевозчик") '8
            .Add("Фрахт пер")
            .Add("Валюта пер")
            .Add("Срок оплаты пер")
            .Add("Дата оплаты пер")
            .Add("Cумма пер")
        End With

        For u As Integer = 0 To ds6.Rows.Count - 1

            If Not dc.ContainsKey(ds6.Rows(u).Item(2)) Then 'перебираем dictionary 
                dc.Add(ds6.Rows(u).Item(2), 1)
            Else
                Dim d As Integer = dc(ds6.Rows(u).Item(2))
                d += 1
                dc(ds6.Rows(u).Item(2)) = d

            End If

        Next

        For pf As Integer = 0 To ds7.Rows.Count - 1 'выбираем количество повторений номеров рейсов

            If Not dp.ContainsKey(ds7.Rows(pf).Item(2)) Then 'перебираем dictionary 
                dp.Add(ds7.Rows(pf).Item(2), 1)
            Else
                Dim d As Integer = dp(ds7.Rows(pf).Item(2))
                d += 1
                dp(ds7.Rows(pf).Item(2)) = d

            End If
        Next

        For Each n In dc.Keys
            If dp.ContainsKey(n) Then
                If dc(n) >= dp(n) Then
                    dmax.Add(n, dc(n))
                Else
                    dmax.Add(n, dp(n))
                End If
            Else
                dmax.Add(n, dc(n))
            End If
        Next



        For Each b In dmax.Keys 'вставляем номера рейсов в таблицу
            For l As Integer = 1 To dmax(b)
                Dim row As DataRow = ds8.NewRow
                row("Рейс") = b
                ds8.Rows.Add(row)
            Next
        Next

        Dim dop As New Dictionary(Of Integer, Integer)

        For j As Integer = 0 To ds8.Rows.Count - 1 'клиент сбор данных в промежуточную таблицу

            For c As Integer = 0 To ds6.Rows.Count - 1 '
                If ds8.Rows(j).Item(0) = ds6.Rows(c).Item(2) Then

                    If Not dop.ContainsValue(ds6.Rows(c).Item(2)) Then
                        dop.Add(ds6.Rows(c).Item(0), ds6.Rows(c).Item(2))
                        ds8.Rows(j).Item(6) = Strings.Left(ds6.Rows(c).Item(3), 10)
                        ds8.Rows(j).Item(7) = ds6.Rows(c).Item(4)
                        Exit For
                    Else
                        If dop.ContainsKey(ds6.Rows(c).Item(0)) Then 'проверяем не повотряются ли ключ и значение
                            Continue For
                        Else
                            dop.Add(ds6.Rows(c).Item(0), ds6.Rows(c).Item(2))
                            ds8.Rows(j).Item(6) = Strings.Left(ds6.Rows(c).Item(3), 10)
                            ds8.Rows(j).Item(7) = ds6.Rows(c).Item(4)
                            Exit For
                        End If
                    End If

                End If
            Next
        Next

        Dim dop1 As New Dictionary(Of Integer, Integer)
        For j As Integer = 0 To ds8.Rows.Count - 1 'перевозчик сбор данных в промежуточную таблицу

            For c As Integer = 0 To ds7.Rows.Count - 1 '
                If ds8.Rows(j).Item(0) = ds7.Rows(c).Item(2) Then

                    If Not dop1.ContainsValue(ds7.Rows(c).Item(2)) Then
                        dop1.Add(ds7.Rows(c).Item(0), ds7.Rows(c).Item(2))
                        ds8.Rows(j).Item(12) = Strings.Left(ds7.Rows(c).Item(3), 10)
                        ds8.Rows(j).Item(13) = ds7.Rows(c).Item(4)
                        Exit For
                    Else
                        If dop1.ContainsKey(ds7.Rows(c).Item(0)) Then 'проверяем не повотряются ли ключ и значение
                            Continue For
                        Else
                            dop1.Add(ds7.Rows(c).Item(0), ds7.Rows(c).Item(2))
                            ds8.Rows(j).Item(12) = Strings.Left(ds7.Rows(c).Item(3), 10)
                            ds8.Rows(j).Item(13) = ds7.Rows(c).Item(4)
                            Exit For
                        End If
                    End If

                End If
            Next
        Next

        For f As Integer = 0 To ds8.Rows.Count - 1
            Dim row As DataRow = ds.NewRow
            row("Рейс") = ds8.Rows(f).Item(0)
            row(6) = ds8.Rows(f).Item(6)
            row(7) = ds8.Rows(f).Item(7)
            row(12) = ds8.Rows(f).Item(12)
            row(13) = ds8.Rows(f).Item(13)
            ds.Rows.Add(row)
            If Not dp3.ContainsKey(ds8.Rows(f).Item(0)) Then
                dp3.Add(ds8.Rows(f).Item(0), ds8.Rows(f).Item(7))
                lts.Add(ds8.Rows(f).Item(0))

            End If
        Next

        Dim v As Integer = ds.Rows.Count
        Dim v1 As Integer = ds.Columns.Count


        Dim lir As New List(Of Integer)()
        For m As Integer = 0 To ds.Rows.Count - 1 'вставляем данные первой доп строки в осн строку
            If Not dp4.ContainsKey(ds.Rows(m).Item(0)) Then
                dp4.Add(ds.Rows(m).Item(0), m)
            Else
                'If IsDBNull(ds.Rows(dp4(ds.Rows(m).Item(0))).Item(7)) And IsDBNull(ds.Rows(dp4(ds.Rows(m).Item(0))).Item(13)) Then
                '    ds.Rows(dp4(ds.Rows(m).Item(0))).Item(6) = ds.Rows(m).Item(6).ToString
                '    ds.Rows(dp4(ds.Rows(m).Item(0))).Item(7) = ds.Rows(m).Item(7).ToString
                '    ds.Rows(dp4(ds.Rows(m).Item(0))).Item(12) = ds.Rows(m).Item(12).ToString
                '    ds.Rows(dp4(ds.Rows(m).Item(0))).Item(13) = ds.Rows(m).Item(13).ToString
                '    lir.Add(m)
                'ElseIf Not IsDBNull(ds.Rows(dp4(ds.Rows(m).Item(0))).Item(7)) And IsDBNull(ds.Rows(dp4(ds.Rows(m).Item(0))).Item(13)) Then
                '    If ds.Rows(m).Item(12).ToString <> "" Then
                '        ds.Rows(dp4(ds.Rows(m).Item(0))).Item(12) = ds.Rows(m).Item(12).ToString
                '        ds.Rows(dp4(ds.Rows(m).Item(0))).Item(13) = ds.Rows(m).Item(13).ToString
                '        ds.Rows(m).Item(13) = ""
                '        ds.Rows(m).Item(12) = ""
                '    End If
                'ElseIf IsDBNull(ds.Rows(dp4(ds.Rows(m).Item(0))).Item(7)) And Not IsDBNull(ds.Rows(dp4(ds.Rows(m).Item(0))).Item(13)) Then
                '    If ds.Rows(m).Item(7).ToString <> "" Then
                '        ds.Rows(dp4(ds.Rows(m).Item(0))).Item(6) = ds.Rows(m).Item(6).ToString
                '        ds.Rows(dp4(ds.Rows(m).Item(0))).Item(7) = ds.Rows(m).Item(7).ToString
                '        ds.Rows(m).Item(6) = ""
                '        ds.Rows(m).Item(7) = ""
                '    End If

                ds.Rows(dp4(ds.Rows(m).Item(0))).Item(6) &= ds.Rows(m).Item(6).ToString & vbCrLf
                ds.Rows(dp4(ds.Rows(m).Item(0))).Item(7) &= ds.Rows(m).Item(7).ToString & vbCrLf
                ds.Rows(dp4(ds.Rows(m).Item(0))).Item(12) &= ds.Rows(m).Item(12).ToString & vbCrLf
                ds.Rows(dp4(ds.Rows(m).Item(0))).Item(13) &= ds.Rows(m).Item(13).ToString & vbCrLf
            End If
        Next

        'For Each item As Integer In lir 'удалеям ненужные строки
        '    For k As Integer = 0 To ds.Rows.Count - 1
        '        If k = item Then
        '            ds.Rows(k).Delete()
        '        End If
        '    Next
        'Next
        dp4.Clear()

        ds.DefaultView.Sort = "Клиент" & " ASC"

        Dim dm As New DataTable
        With dm.Columns
            .Add("Рейс")
            .Add("Клиент")
            .Add("Маршрут")
            .Add("Фрахт")
            .Add("Валюта")
            .Add("Срок оплаты")
            .Add("Дата оплаты")
            .Add("Cумма")
            .Add("Перевозчик") '8
            .Add("Фрахт пер")
            .Add("Валюта пер")
            .Add("Срок оплаты пер")
            .Add("Дата оплаты пер")
            .Add("Cумма пер")
            .Add("Дельта")
        End With

        For m1 As Integer = 0 To ds.Rows.Count - 1

            If ds.Rows(m1).Item(1).ToString <> "" Then
                Dim row As DataRow = dm.NewRow
                row("Рейс") = ds.Rows(m1).Item(0)
                row("Клиент") = ds.Rows(m1).Item(1)
                If ds.Rows(m1).Item(2).ToString.Length > 40 Then 'если длинный текст
                    ds.Rows(m1).Item(2) = Strings.Left(ds.Rows(m1).Item(2), 40) & " ..."
                    row("Маршрут") = ds.Rows(m1).Item(2)
                Else
                    row("Маршрут") = ds.Rows(m1).Item(2)
                End If
                row("Фрахт") = ds.Rows(m1).Item(3)
                row("Валюта") = ds.Rows(m1).Item(4)
                If Not IsDBNull(ds.Rows(m1).Item(5)) Then
                    row("Срок оплаты") = Strings.Left(ds.Rows(m1).Item(5), 10)
                End If
                row("Дата оплаты") = ds.Rows(m1).Item(6)
                row("Cумма") = ds.Rows(m1).Item(7)
                row("Перевозчик") = ds.Rows(m1).Item(8)
                row("Фрахт пер") = ds.Rows(m1).Item(9)
                row("Валюта пер") = ds.Rows(m1).Item(10)
                row("Срок оплаты пер") = ds.Rows(m1).Item(11)
                row("Дата оплаты пер") = ds.Rows(m1).Item(12)
                row("Cумма пер") = ds.Rows(m1).Item(13)
                Try
                    row("Дельта") = Math.Round(CType(ds.Rows(m1).Item(3), Double) - CType(ds.Rows(m1).Item(9), Double), 2)
                Catch ex As Exception

                End Try




                dm.Rows.Add(row)
                End If


        Next

        dsall = dm





    End Sub
    Private Sub Запуск(ByVal ds As DataTable)
        Me.Cursor = Cursors.WaitCursor
        'ds.DefaultView.Sort = "Рейс" & " ASC" 'сортировка datatable
        ds.DefaultView.Sort = "Рейс" & " DESC"
        Grid1.DataSource = ds
        GridView(Grid1)

        Grid1.RowsDefaultCellStyle.BackColor = Color.Lavender
        Grid1.AlternatingRowsDefaultCellStyle.BackColor = Color.LightBlue
        'gridselect()




        Grid1.BorderStyle = BorderStyle.Fixed3D
        Dim font As New Font(Grid1.DefaultCellStyle.Font.FontFamily, 11, FontStyle.Bold)
        Grid1.Columns(6).DefaultCellStyle.Font = font
        Grid1.Columns(7).DefaultCellStyle.ForeColor = Color.DarkRed
        Grid1.Columns(7).DefaultCellStyle.Font = font
        Grid1.Columns(12).DefaultCellStyle.Font = font
        Grid1.Columns(13).DefaultCellStyle.ForeColor = Color.DarkRed
        Grid1.Columns(13).DefaultCellStyle.Font = font
        Grid1.Columns(0).DefaultCellStyle.Font = font
        'Grid1.Columns(7).DefaultCellStyle.WrapMode = DataGridViewCellBorderStyle.SunkenVertical


        Grid1.Columns(0).Width = 60
        Grid1.Columns(1).Width = 160
        Grid1.Columns(2).Width = 260
        Grid1.Columns(8).Width = 160
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub gridselect()

        For Each item As Integer In lts
            For i As Integer = 0 To Grid1.Rows.Count - 1
                For Each rf As DataGridViewRow In Grid1.Rows
                    'If DataGridView1.Rows(i).Cells(0).Value = item Then
                    '    DataGridView1.DefaultCellStyle.ForeColor = Color.Gold
                    '    rf.DefaultCellStyle.BackColor = Color.Gray
                    '    Exit For
                    'End If
                Next
            Next

            Next


    End Sub
    Private Sub Перевоз()
        Dim strsql As String = "SELECT НомерРейса as [Рейс], НазвОрганизации as [Организация], СтоимостьФрахта as [Фрахт],
Валюта as [Валюта], ДатаОплаты as [Срок оплаты] FROM РейсыПеревозчика WHERE НомерРейса IS NOT NULL"
        Dim ds As DataTable = Selects3(strsql)

        ds.Columns.Add("Дата оплаты")
        ds.Columns.Add("Cумма")

        Dim strsql1 As String = "SELECT DISTINCT Рейс FROM ОплатыПер"
        Dim ds1 As DataTable = Selects3(strsql1)

        For i As Integer = 0 To ds.Rows.Count - 1
            For y As Integer = 0 To ds1.Rows.Count - 1
                If ds.Rows(i).Item(0) = ds1.Rows(y).Item(0) Then
                    Dim ds3 As DataTable = ДобОплатВТаблицу(ds.Rows(i).Item(0), 0)
                    For z As Integer = 0 To ds3.Rows.Count - 1
                        Dim row As DataRow = ds.NewRow
                        row("Рейс") = ds.Rows(i).Item(0)
                        row("Дата оплаты") = ds3.Rows(z).Item(3)
                        row("Cумма") = ds3.Rows(z).Item(4)
                        ds.Rows.Add(row)
                    Next
                End If

            Next
        Next
        ds.DefaultView.Sort = "Рейс" & " ASC" 'сортировка datatable

        dsper = ds
    End Sub
    Private Function ДобОплатВТаблицу(ByVal d As Integer, ByVal f As Integer) As DataTable
        Dim strsql As String
        If f = 1 Then
            strsql = "SELECT * FROM ОплатыКлиент WHERE Рейс=" & d & ""

        Else
            strsql = "SELECT * FROM ОплатыПер WHERE Рейс=" & d & ""

        End If
        Dim ds As DataTable = Selects3(strsql)
        Return ds
    End Function

    Private Sub Grid1_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles Grid1.CellDoubleClick
        НомРейса = Grid1.SelectedCells(0).Value
        bl = True
        Выбор.ShowDialog()

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        dc.Clear()
        dp.Clear()
        dmax.Clear()
        dsall.Clear()
        ALL()
        Запуск(dsall)
    End Sub
End Class