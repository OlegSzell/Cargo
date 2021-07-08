Imports System.IO
Imports System.Text
Imports iTextSharp.text.pdf
Imports iTextSharp.text.pdf.parser

Public Class ОтчетДляОлега
    Dim WordsScan As New List(Of String)()
    Dim dt2 As New DataTable
    Private com2 As New List(Of IDNaz)
    Private bscom2 As New BindingSource
    Private ВыпискаВыборка As New List(Of Выписка)

    Private Sub Button1_Click(sender As Object, e As EventArgs)

        Dim filePath As String = "c:\Users\Oleg\Desktop\3.xml"
        DataSet.ReadXml(filePath)
        Grid1.DataSource = DataSet.Tables(0)
        Grid1.DataMember = "authors"

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs)


        Dim swXML As New System.IO.StringWriter()
        DataSet.WriteXmlSchema(swXML)
        'TextBox1.Text = swXML.ToString

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs)
        'чтение из файла
        Using fstream As FileStream = File.OpenRead("c:\Users\Oleg\Desktop\сентябрь.txt")
            'преобразуем строку в байты
            Dim array As Byte() = New Byte(fstream.Length) {}
            'чтение данных

            fstream.Read(array, 0, array.Length)
            'For i As Integer = 0 To fstream.Length - 1
            '    DataGridView1.Rows.Add(fstream.BeginRead(DataGridView1))
            'Next

            'декодируем байты в строку
            Dim textFromFile As String = System.Text.Encoding.Default.GetString(array)
            'Console.WriteLine("Текст из файла: {0}", textFromFile)
            'TextBox1.Text = textFromFile

        End Using

    End Sub
    Private Sub КопированиеИзБуфераВЭксель(ByVal pathopfd As String)
        'Dim f As String = IO.File.ReadAllText(pathopfd, System.Text.Encoding.UTF8)
        Dim f As String
        Dim mf = IO.File.OpenText(pathopfd)
        f = mf.ReadToEnd()


        My.Computer.Clipboard.SetText(f)
        Dim xlapp1 As New Microsoft.Office.Interop.Excel.Application
        Dim xlworkbook1 As Microsoft.Office.Interop.Excel.Workbook
        Dim xlworksheet1 As Microsoft.Office.Interop.Excel.Worksheet

        xlapp1.Workbooks.Add()
        xlworkbook1 = xlapp1.ActiveWorkbook

        xlworksheet1 = xlworkbook1.Worksheets(1)
        xlworkbook1.Worksheets(1).Range("A1").Select
        xlworkbook1.Worksheets(1).Paste
        xlworkbook1.SaveAs("C:\Users\Oleg\Desktop\ОтчетОлег\" & Now.Month.ToString & ".xlsx") '& " (" & Now.ToShortTimeString & ")"

        xlworkbook1.Close(False)
        xlapp1.Quit()

        releaseobject(xlapp1)
        releaseobject(xlworkbook1)
        releaseobject(xlworksheet1)



        'Копирование содержимого без заголовков
        '    'Grid1.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableWithoutHeaderText


        '    'Grid1.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText


        '    'Выделение содержимого DGV
        '    'Grid1.SelectAll()

        '    'Помещаем в буфер обмена выделенные ячейки
        '    'Clipboard.SetDataObject(Grid1.GetClipboardContent())

        '    Clipboard.

        '    Dim путь = "C:\Users\" & My.Computer.Name & "\Documents\dgv.html"
        '    Dim путь1 = "C:\Users\" & My.Computer.Name & "\Documents\dgv.txt"
        '    Dim путь2 = "C:\Users\" & My.Computer.Name & "\Documents\dgv.rtf"
        '    Dim путь3 = "C:\Users\" & My.Computer.Name & "\Documents\dgv.xlsx"


        '    'Записываем текст из буфера обмена в файл
        '    Using writer As New StreamWriter(путь1, False, System.Text.Encoding.Unicode)
        '        writer.Write(Clipboard.GetText())

        '        'writer.Write(Clipboard.GetText(TextDataFormat.html))
        '        'writer.Write(Clipboard.GetText(TextDataFormat.Rtf))
        '    End Using

        '    'Dim filePath = My.Computer.FileSystem.OpenTextFileReader(путь2) 'в этой строке нет необходимости, просто так для примера
        '    'filePath.Close()


        '    'Dim currentdirectory = "C:\"
        '    'Dim filename = "dgv"
        '    'Dim path = System.IO.Directory.GetFiles(currentdirectory, filename & "*", IO.SearchOption.AllDirectories)(0)
        '    'IO.File.Open(путь, FileMode.Open)


        '    'Process.Start(путь3, Chr(34) & путь1 & Chr(34))
        '    Process.Start("excel.exe", Chr(34) & путь1 & Chr(34))




    End Sub
    Private Sub AddInBase(d As List(Of String))
        Dim datY As Integer
        Dim datM As Integer
        Dim ManthName As String
        Try
            datY = CDate(Replace(d(0), "/", ".")).Year
            datM = CDate(Replace(d(0), "/", ".")).Month
            ManthName = MonthName(datM)
        Catch ex As Exception
            MessageBox.Show("Что пошло нетак с файлом txt!", Рик)
            Return
        End Try
        Using db As New dbAllDataContext(_cn3)
            Dim f = (From x In db.Выписка
                     Where x.Год = datY And x.МесяцЧислом = datM
                     Select x).ToList()

            If f IsNot Nothing Then
                If f.Count > 0 Then
                    If MessageBox.Show("Заменить данные?", Рик, MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.No Then
                        Return
                    Else
                        db.Выписка.DeleteAllOnSubmit(f)
                        db.SubmitChanges()
                    End If
                End If
            End If


            Dim oborРасход As Double = 0
            Dim obor1Приход As Double = 0
            Dim lst As New List(Of Выписка)
            For mc As Integer = 0 To d.Count - 1 Step 9
                If d(mc).Contains("Обороты") Then
                    oborРасход = CDbl(Replace(Trim(d(mc + 1)), ".", ","))
                    obor1Приход = CDbl(Replace(Trim(d(mc + 2)), ".", ","))
                    Continue For
                End If
                Dim m As New Выписка With {.Год = datY, .Месяц = ManthName, .МесяцЧислом = datM, .ДатаОперации = Replace(d(mc), "/", "."), .Расход = d(mc + 5),
                    .Приход = d(mc + 6), .Организация = d(mc + 7), .НазначениеПлатежа = d(mc + 8)}
                lst.Add(m)
            Next
            Dim dbRashod As Double = 0
            Dim dbPrichod As Double = 0
            For Each b In lst
                dbRashod += Math.Round(CDbl(Replace(Trim(b.Расход), ".", ",")), 2)
                dbPrichod += Math.Round(CDbl(Replace(Trim(b.Приход), ".", ",")), 2)
            Next
            If oborРасход = Not dbRashod Then
                MessageBox.Show("Не сходится расход!", Рик)
                Return
            End If
            If obor1Приход = Not dbPrichod Then
                MessageBox.Show("Не сходится приход!", Рик)
                Return
            End If

            db.Выписка.InsertAllOnSubmit(lst)
            db.SubmitChanges()


        End Using




    End Sub
    Private Sub Txt()






        Dim pathopfd As String
        Dim opfd As OpenFileDialog
        opfd = OpenFileDialog1
        opfd.InitialDirectory = "C:\Users\Oleg\Desktop\ОтчетОлег\"
        opfd.Filter = "Файлы txt|*.txt"
        opfd.Title = "Выберите файл txt"
        opfd.FileName = "txt"
        If opfd.ShowDialog() = DialogResult.OK Then
            pathopfd = opfd.FileName
            If Not IO.Path.GetExtension(pathopfd) = ".txt" Then

                MessageBox.Show("Выберите txt - файл!", Рик)

                Exit Sub

            End If
            Dim Massiv() As String = IO.File.ReadAllLines(pathopfd, System.Text.Encoding.Default)
            'Dim read As New System.IO.StreamReader(pathopfd)
            'Dim Massiv As New ArrayList() From {IO.File.ReadAllLines(pathopfd, System.Text.Encoding.Default)}
            'Massiv.RemoveRange(0, 29)
            Dim _list As New List(Of String)()
            For x As Integer = 0 To 30
                Massiv(x) = ""
            Next
            For x1 As Integer = 0 To Massiv.Count - 1
                If Not Massiv(x1) = "" Then
                    _list.Add(Massiv(x1))
                End If
            Next

            AddInBase(_list)
            Dim f As Integer = _list.Count

            Dim dtable As New DataTable
            dtable.Columns.Add("Дата")
            dtable.Columns.Add("Организация")
            dtable.Columns.Add("Основание платежа")
            dtable.Columns.Add("Приход")
            dtable.Columns.Add("Расход")


            Dim dtable1 As New DataTable
            dtable1.Columns.Add("Дата")
            dtable1.Columns.Add("Организация")
            dtable1.Columns.Add("Основание платежа")
            dtable1.Columns.Add("Приход")
            dtable1.Columns.Add("Расход")

            Dim ИтогоПриход, ИтогоРасход As Double





            Dim row As DataRow
            For h As Integer = 8 To _list.Count - 1 Step 9
                If Not _list(h - 3) = "0.00" Then 'расход
                    If Strings.Left(_list(h - 8).ToString, 10).Contains("Обор") Then Continue For
                    row = dtable.NewRow()
                    row("Дата") = Strings.Left(_list(h - 8).ToString, 10)
                    row("Организация") = LCase(_list(h - 1).ToString)
                    row("Основание платежа") = LCase(_list(h).ToString)
                    row("Приход") = ""
                    If _list(h - 3).ToString.Contains(".") And _list(h - 3).ToString.Contains(",") Then
                        Dim ch As Char = "."
                        Dim i As Integer = _list(h - 3).ToString.IndexOf(ch)
                        _list(h - 3) = _list(h - 3).ToString.Substring(0, _list(h - 3).ToString.Length - (_list(h - 3).ToString.Length - i))
                        _list(h - 3) = Replace(_list(h - 3), ",", "")
                        _list(h - 3) = Trim(_list(h - 3))
                        row("Расход") = _list(h - 3).ToString
                        ИтогоРасход += _list(h - 3).ToString
                    Else
                        row("Расход") = CDbl(Replace(_list(h - 3).ToString, ".", ","))
                        ИтогоРасход += CDbl(Replace(_list(h - 3).ToString, ".", ","))
                    End If

                    dtable.Rows.Add(row)
                Else 'приход
                    row = dtable1.NewRow()
                    row("Дата") = Strings.Left(_list(h - 8).ToString, 10)
                    row("Организация") = LCase(_list(h - 1).ToString)
                    row("Основание платежа") = LCase(_list(h).ToString)
                    If _list(h - 2).ToString.Contains(".") And _list(h - 2).ToString.Contains(",") Then
                        Dim ch As Char = "."
                        Dim i As Integer = _list(h - 2).ToString.IndexOf(ch)
                        _list(h - 2) = _list(h - 2).ToString.Substring(0, _list(h - 2).ToString.Length - (_list(h - 2).ToString.Length - i))
                        _list(h - 2) = Replace(_list(h - 2), ",", "")
                        _list(h - 2) = Trim(_list(h - 2))
                        row("Приход") = _list(h - 2).ToString
                        ИтогоПриход += _list(h - 2).ToString
                    Else
                        row("Приход") = CDbl(Replace(_list(h - 2).ToString, ".", ","))
                        ИтогоПриход += CDbl(Replace(_list(h - 2).ToString, ".", ","))
                    End If
                    row("Расход") = ""
                    dtable1.Rows.Add(row)
                End If
            Next

            dtable.Merge(dtable1)

            row = dtable.NewRow()
            row("Организация") = "ИТОГО"
            row("Приход") = ИтогоПриход
            row("Расход") = ИтогоРасход
            dtable.Rows.Add(row)

            grid1Color(dtable)

            'Dim _string As String = read.ReadToEnd
            'read.Close()
        Else
            MessageBox.Show("Вы не выбрали файл!", Рик)
            Exit Sub
        End If

        Grid2s(Grid1)
    End Sub
    Private Class Grid1Class
        Public Property Номер As Integer
        Public Property Дата As String
        Public Property Организация As String
        Public Property Основание_платежа As String
        Public Property Приход As String
        Public Property Расход As String


    End Class

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Txt()
        Exit Sub



        Dim xlapp1 As New Microsoft.Office.Interop.Excel.Application
        Dim xlworkbook1 As Microsoft.Office.Interop.Excel.Workbook
        Dim xlworksheet1 As Microsoft.Office.Interop.Excel.Worksheet

        Dim pathopfd As String
        Dim opfd As New OpenFileDialog
        opfd = OpenFileDialog1
        opfd.Title = "Выбрать файл Excel"
        opfd.FileName = "Имя файла"
        opfd.InitialDirectory = "C:\Users\Oleg\Desktop\ОтчетОлег\"
        opfd.Filter = "Все файлы|*.*"
        'opfd.Filter = "Эксель файлы|*.xlsx"
        If opfd.ShowDialog() = DialogResult.OK Then
            pathopfd = opfd.FileName
            'Dim Massiv() As String = IO.File.ReadAllLines(pathopfd, System.Text.Encoding.Default)
            КопированиеИзБуфераВЭксель(pathopfd)
        Else
            MessageBox.Show("Вы не выбрали файл!", Рик)
            Exit Sub
        End If

        Me.Cursor = Cursors.WaitCursor

        xlworkbook1 = xlapp1.Workbooks.Open(pathopfd)
        xlworksheet1 = xlworkbook1.ActiveSheet

        Dim df As New DataTable

        Dim rcouunt As Integer = xlworksheet1.UsedRange.Columns.Count 'количество столбцов в экселе
        Dim rcouunt1 As Integer = xlworksheet1.UsedRange.Rows.Count
        'Grid1.ColumnCount = rcouunt
        For v As Integer = 0 To rcouunt - 1
            'Grid1.Columns(v).Name = "Столбец " & v.ToString
            'Grid1.Columns(v).Name = xlworksheet1.Cells(1, v + 1).Value
            df.Columns.Add(v.ToString)
        Next



        For i As Integer = 0 To rcouunt1
            df.Rows.Add(xlworksheet1.Cells(i + 1, 1).Value, xlworksheet1.Cells(i + 1, 2).Value, xlworksheet1.Cells(i + 1, 3).Value,
xlworksheet1.Cells(i + 1, 4).Value, xlworksheet1.Cells(i + 1, 5).Value, xlworksheet1.Cells(i + 1, 6).Value, xlworksheet1.Cells(i + 1, 7).Value)
        Next

        Dim ИтогоПриход, ИтогоРасход As Double

        Dim dtable As New DataTable
        dtable.Columns.Add("Дата")
        dtable.Columns.Add("Организация")
        dtable.Columns.Add("Основание платежа")
        dtable.Columns.Add("Приход")
        dtable.Columns.Add("Расход")




        Dim row As DataRow

        For x As Integer = 1 To df.Rows.Count - 1
            If Not df.Rows(x).Item(6).ToString = "0.00" And DBNull.Value.Equals(df.Rows(x).Item(6)) = False Then 'пустая ячейка datatable проверка
                If df.Rows(x).Item(0).ToString.Contains("Входящее сальдо") Then Continue For
                If df.Rows(x).Item(6).ToString.Contains("Кредит") Then Continue For
                If df.Rows(x).Item(0).ToString.Contains("Обороты") Then Exit For

                Dim sum As Double
                Dim leiht As String = df.Rows(x).Item(6).ToString
                leiht = leiht.Length 'полная длина
                Dim leihtobrez As Integer = CType(leiht, Integer) - 3 'обрезанная длина
                Dim h As String = Strings.Right(df.Rows(x).Item(6).ToString, 2)
                If h = "00" Then
                    h = Strings.Left(df.Rows(x).Item(6).ToString, leihtobrez)
                    If h.Contains(",") Then
                        h = Replace(h, ",", "")
                        h = Trim(h)
                    End If
                    sum = CType(h, Integer)
                Else
                    'h = Strings.Left(df.Rows(x).Item(5).ToString, leihtobrez)
                    'h = Replace(h, ",", ".")
                    Dim l As Integer = Len(df.Rows(x).Item(6).ToString)
                    Dim l2 As String = df.Rows(x).Item(6).ToString
                    l2 = Replace(l2, ".", ",")

                    For f As Integer = 0 To 1
                        l2 = Replace(l2, ",", "")
                    Next

                    Dim l3 As Integer = l - Len(l2)


                    If l3 = 1 Then
                        sum = CDbl(Replace(df.Rows(x).Item(6).ToString, ".", ","))
                    Else
                        'Dim l4 As String = Replace(df.Rows(x).Item(6).ToString, ".", ",")
                        Dim ch As Char = "."
                        Dim i As Integer = df.Rows(x).Item(6).ToString.IndexOf(ch)

                        df.Rows(x).Item(6) = df.Rows(x).Item(6).ToString.Substring(0, df.Rows(x).Item(6).ToString.Length - (df.Rows(x).Item(6).ToString.Length - i))
                        'For f As Integer = 0 To 0
                        '    df.Rows(x).Item(6) = Replace(df.Rows(x).Item(6).ToString, ",", "")
                        'Next
                        sum = CDbl(Replace(df.Rows(x).Item(6).ToString, ".", ","))

                    End If

                End If
                ИтогоПриход += sum
                Try
                    row = dtable.NewRow()
                    row("Дата") = Strings.Left(df.Rows(x).Item(0).ToString, 10)
                    row("Организация") = LCase(df.Rows(x + 1).Item(0).ToString)
                    row("Основание платежа") = LCase(df.Rows(x + 2).Item(0).ToString)
                    row("Приход") = sum
                    row("Расход") = ""
                    dtable.Rows.Add(row)
                Catch ex As Exception

                End Try

            End If

        Next


        For x As Integer = 1 To df.Rows.Count - 1
            If Not df.Rows(x).Item(5).ToString = "0.00" And DBNull.Value.Equals(df.Rows(x).Item(5)) = False Then 'пустая ячейка datatable проверка Not IsDate(df.Rows(x).Item(5))
                If df.Rows(x).Item(0).ToString.Contains("Входящее сальдо") Then Continue For
                If df.Rows(x).Item(0).ToString.Contains("Обороты") Then Exit For
                If df.Rows(x).Item(5).ToString.Contains("Дебет") Then Continue For
                If df.Rows(x).Item(5).ToString.Contains("Номинал") Then Continue For

                Dim sum As Double
                Dim leiht As String = df.Rows(x).Item(5).ToString
                leiht = leiht.Length 'полная длина
                Dim leihtobrez As Integer = CType(leiht, Integer) - 3 'обрезанная длина
                Dim h As String = Strings.Right(df.Rows(x).Item(5).ToString, 2)
                If h = "00" Then
                    h = Strings.Left(df.Rows(x).Item(5).ToString, leihtobrez)
                    If h.Contains(",") Then
                        h = Replace(h, ",", "")
                        h = Trim(h)
                    End If
                    sum = CType(h, Integer)
                Else
                    'h = Strings.Left(df.Rows(x).Item(5).ToString, leihtobrez)
                    'h = Replace(h, ",", ".")

                    Dim l As Integer = Len(df.Rows(x).Item(5).ToString)
                    Dim l2 As String = df.Rows(x).Item(5).ToString
                    l2 = Replace(l2, ".", ",")

                    For f As Integer = 0 To 1
                        l2 = Replace(l2, ",", "")
                    Next

                    Dim l3 As Integer = l - Len(l2)


                    If l3 = 1 Then
                        sum = CDbl(Replace(df.Rows(x).Item(5).ToString, ".", ","))
                    Else
                        'Dim l4 As String = Replace(df.Rows(x).Item(6).ToString, ".", ",")
                        Dim ch As Char = "."
                        Dim i As Integer = df.Rows(x).Item(5).ToString.IndexOf(ch)

                        df.Rows(x).Item(5) = df.Rows(x).Item(5).ToString.Substring(0, df.Rows(x).Item(5).ToString.Length - (df.Rows(x).Item(5).ToString.Length - i))
                        'For f As Integer = 0 To 0
                        '    df.Rows(x).Item(6) = Replace(df.Rows(x).Item(6).ToString, ",", "")
                        'Next
                        sum = CDbl(Replace(df.Rows(x).Item(5).ToString, ".", ","))

                    End If
                End If
                ИтогоРасход += sum
                Try
                    row = dtable.NewRow()
                    row("Дата") = Strings.Left(df.Rows(x).Item(0).ToString, 10)
                    row("Организация") = LCase(df.Rows(x + 1).Item(0).ToString)
                    row("Основание платежа") = LCase(df.Rows(x + 2).Item(0).ToString)
                    row("Приход") = ""
                    row("Расход") = sum
                    dtable.Rows.Add(row)
                Catch ex As Exception

                End Try

            End If
        Next

        row = dtable.NewRow()
        row("Организация") = "ИТОГО"
        row("Приход") = ИтогоПриход
        row("Расход") = ИтогоРасход
        dtable.Rows.Add(row)




        grid1Color(dtable)

        'Grid1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells 'хорошая вещь для подбора размера текста
        xlworkbook1.Close(False)
        xlapp1.Quit()

        releaseobject(xlapp1)
        releaseobject(xlworkbook1)
        releaseobject(xlworksheet1)
        Grid2s(Grid1)

        Me.Cursor = Cursors.Default
    End Sub
    Private Sub grid1Color(ByVal d As DataTable)
        Grid1.DataSource = d
        Dim integ As Integer = Grid1.Rows.Count
        Grid1.Rows(Grid1.Rows.Count - 1).Cells(4).Style.Font = New Font(Grid1.DefaultCellStyle.Font, FontStyle.Bold)
        Grid1.Rows(Grid1.Rows.Count - 1).Cells(3).Style.Font = New Font(Grid1.DefaultCellStyle.Font, FontStyle.Bold)
        GridView(Grid1)
        Grid1.Columns(3).Width = 100
        Grid1.Columns(0).Width = 100
        Grid1.Columns(4).Width = 100
        Dim font As New Font(Grid1.DefaultCellStyle.Font.FontFamily, 12, FontStyle.Italic)
        Grid1.Columns(0).DefaultCellStyle.Font = font
        Grid1.Columns(1).DefaultCellStyle.Font = font
        Grid1.RowsDefaultCellStyle.BackColor = Color.Lavender
        Grid1.AlternatingRowsDefaultCellStyle.BackColor = Color.LightBlue
    End Sub
    Private Sub Grid2s(ByVal grid1 As DataGridView)
        Dim Итого As Double
        Dim row As DataRow

        If DBNull.Value.Equals(dt2) = False Then
            dt2.Clear()
        End If


        For x As Integer = 0 To WordsScan.Count - 1
            For i As Integer = 0 To grid1.Rows.Count - 1
                If DBNull.Value.Equals(grid1.Rows(i).Cells(2).Value) = True Or DBNull.Value.Equals(grid1.Rows(i).Cells(1).Value) = True Then Continue For

                If grid1.Rows(i).Cells(1).Value.Contains(WordsScan.Item(x).ToString) Or grid1.Rows(i).Cells(2).Value.Contains(WordsScan.Item(x).ToString) Then
                    If grid1.Rows(i).Cells(4).Value = "" Then Continue For
                    row = dt2.NewRow()
                    row("Дата") = grid1.Rows(i).Cells(0).Value
                    row("Наименование") = grid1.Rows(i).Cells(1).Value & vbCrLf & grid1.Rows(i).Cells(2).Value
                    row("Сумма") = grid1.Rows(i).Cells(4).Value
                    Итого += CType(grid1.Rows(i).Cells(4).Value, Double)
                    dt2.Rows.Add(row)
                End If
            Next
        Next

        row = dt2.NewRow()
        row("Дата") = "ИТОГО"
        row("Сумма") = Итого

        dt2.Rows.Add(row)



        Grid2.DataSource = dt2
        GridView(Grid2)
        Grid2.Columns(0).Width = 80
        Grid2.Columns(2).Width = 80


    End Sub
    Private Sub ОтчетДляОлега_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.MdiParent = MDIParent1
        WordsScan.AddRange(New String() {"клиент-банк", "шелягович олег", "современные логистические системы", "оплата за интернет", "белтелеком", "технобанк",
            "белпочта", "шелягович вика", "отчисления с фсзн", "филиал белгосстраха", "подоходный налог с заработной платы", "перечисление заработной платы",
            "комиссия банка за банковский перевод", "комиссия за безналичное перечисление денежных средств на текущие счета с оформлением банковских платежных карточек за период",
            "жилтехсервис", "оплата за услуги связи", "ндс по реализации.возврат по решению инспекции мнс", "арендная плата за"})

        dt2.Columns.Add("Дата")
        dt2.Columns.Add("Наименование")
        dt2.Columns.Add("Сумма")

        bscom2.DataSource = com2
        ComboBox2.DataSource = bscom2
        ComboBox2.DisplayMember = "Naz"


        If ВыпискаВыборка IsNot Nothing Then
            ВыпискаВыборка.Clear()
        End If



        Dim com1 As New List(Of String)
        Using db As New dbAllDataContext(_cn3)
            Dim f = (From x In db.Выписка
                     Select x).ToList()
            ВыпискаВыборка.AddRange(f)
            If f IsNot Nothing Then
                If f.Count > 0 Then
                    Dim ea = f.OrderBy(Function(x) x.Год).Select(Function(x) x.Год).Distinct().ToList()


                    ComboBox1.Items.Clear()
                    For Each b In ea
                        ComboBox1.Items.Add(b)
                    Next



                End If
            End If
        End Using
        ComboBox2.Text = String.Empty

    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Dim pathopfd As String
        Dim opfd As OpenFileDialog
        opfd = OpenFileDialog1
        opfd.InitialDirectory = "C:\Users\Oleg\Desktop\ОтчетОлег\"
        opfd.Filter = "PDF Files(*.PDF)|*.PDF|All Files(*.*)|*.*"
        opfd.Title = "Выберите файл pdf"
        opfd.FileName = "pdf"
        If opfd.ShowDialog() = DialogResult.OK Then
            pathopfd = opfd.FileName.ToString
            Dim strText As String = String.Empty
            Dim strText1 As String
            Try
                Dim reader As PdfReader = New PdfReader(pathopfd)

                For page As Integer = 1 To reader.NumberOfPages
                    Dim its As ITextExtractionStrategy = New LocationTextExtractionStrategy()
                    Dim s As String = PdfTextExtractor.GetTextFromPage(reader, page, its)
                    s = Encoding.UTF8.GetString(ASCIIEncoding.Convert(Encoding.[Default], Encoding.UTF8, Encoding.[Default].GetBytes(s)))
                    strText += s
                Next
                strText1 = Replace(strText, "  ", "")
                reader.Close()
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try
            'RichTextBox1.Text = strText1
        End If
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click

        Dim ms As New List(Of String), ms1 As New List(Of String), ms2 As New List(Of String)
        Dim pathopfd As String = String.Empty
        Dim opfd As OpenFileDialog
        opfd = OpenFileDialog1
        opfd.InitialDirectory = "C:\Users\Oleg\Desktop\ОтчетОлег\"
        opfd.Filter = "PDF Files(*.docx)|*.docx|All Files(*.*)|*.*"
        opfd.Title = "Выберите файл docx"
        opfd.FileName = "docx"
        If opfd.ShowDialog() = DialogResult.OK Then
            pathopfd = opfd.FileName.ToString
            Dim strText As String = String.Empty
        End If
        Dim word As New Microsoft.Office.Interop.Word.Application
        Dim missing As Object = Type.Missing
        Dim doc As Microsoft.Office.Interop.Word.Document = word.Documents.Open(pathopfd, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing)

        'doc = word.Documents.Open(pathopfd, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing)
        Dim dt As New DataTable
        Dim dr As DataRow
        dr = dt.NewRow
        Dim sb As New StringBuilder
        sb.Append("<table border='1'><tr>")
        sb.Append("</tr>")
        For Each tableRow As Microsoft.Office.Interop.Word.Row In word.Application.ActiveDocument.Tables(3).Rows
            sb.Append("<tr>")
            For i As Integer = 1 To tableRow.Cells.Count
                sb.Append("<td>" & tableRow.Cells(i).Range.Text.Replace("\r\a", "").Trim() & "</td>")

            Next
            sb.Append("</tr>")
        Next
        sb.Append("</table>")
        'RichTextBox1.Text = sb.ToString
        doc.Close(missing, missing, missing)
        word.Application.Quit()


        'Dim wordApp As Microsoft.Office.Interop.Word.Application = New Microsoft.Office.Interop.Word.Application()
        'Dim filename As Object = pathopfd
        'Dim missing As Object = Type.Missing
        'Dim doc As Microsoft.Office.Interop.Word.Document = wordApp.Documents.Open(filename, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing)
        'Dim table As Microsoft.Office.Interop.Word.Table = doc.Tables(3)


        'For i As Integer = 1 To table.Rows.Count - 1
        '    ms.Add(table.Cell(i, 0).Range.Text)
        '    ms1.Add(table.Cell(i, 1).Range.Text)
        '    ms2.Add(table.Cell(i, 2).Range.Text)
        'Next

        'doc.Close(missing, missing, missing)
        'Runtime.InteropServices.Marshal.ReleaseComObject(wordApp)


    End Sub

    Private Sub ComboBox2_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles ComboBox2.SelectionChangeCommitted
        Dim Год = ComboBox1.Text
        Dim f As IDNaz = ComboBox2.SelectedItem
        If Год = String.Empty Then
            MessageBox.Show("Выберите год!", Рик)
            Return
        End If
        Dim f1 As New List(Of Grid1Class)

        f1 = (From x In ВыпискаВыборка
              Order By x.ДатаОперации
              Where x.Год = Год And x.МесяцЧислом = f.ID
              Select New Grid1Class With {.Дата = x.ДатаОперации, .Организация = x.Организация, .Основание_платежа = x.НазначениеПлатежа, .Приход = x.Приход, .Расход = x.Расход}).ToList()

        Dim prih As Double = Nothing, rash As Double = Nothing
        If f1 IsNot Nothing Then
            If f1.Count > 0 Then
                Dim i As Integer = 1
                For Each b In f1
                    b.Номер = i
                    i += 1
                    prih += CDbl(Replace(b.Приход, ".", ","))
                    rash += CDbl(Replace(b.Расход, ".", ","))
                Next
                Dim f2 As New Grid1Class With {.Организация = "- ИТОГО - ", .Приход = Math.Round(prih, 2), .Расход = Math.Round(rash, 2)}
                f1.Add(f2)
                Grid1.BeginInvoke(New MethodInvoker(Sub() grd1(f1))) 'поток параллельно с контролами


            End If
        End If

    End Sub
    Private Sub grd1(d As List(Of Grid1Class))
        Dim mf As New BindingSource
        mf.DataSource = d
        Grid1.DataSource = mf
        GridView(Grid1)
        Grid1.Columns(0).Width = 60
        Grid1.Columns(1).Width = 80
        Grid1.Columns(4).Width = 90
        Grid1.Columns(5).Width = 90
        mf.ResetBindings(False)
    End Sub

    Private Sub ComboBox1_SelectedValueChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedValueChanged

        If com2 IsNot Nothing Then
            com2.Clear()
        End If


        Dim es1 = (From x In ВыпискаВыборка
                   Where x.Год = ComboBox1.Text
                   Group x By Keys = New With {Key x.Месяц, Key x.МесяцЧислом}
                 Into Group
                   Select New IDNaz With {.ID = Keys.МесяцЧислом, .Naz = Keys.Месяц}).OrderBy(Function(x) x.ID).ToList()

        If es1 IsNot Nothing Then
            com2.AddRange(es1)
            bscom2.ResetBindings(False)
        End If
        ComboBox2.Text = String.Empty

    End Sub
End Class