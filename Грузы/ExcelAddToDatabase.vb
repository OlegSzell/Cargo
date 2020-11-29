
Imports ClosedXML.Excel
Imports Microsoft.Office.Interop.Excel
Imports System.IO
Imports System.Reflection

Public Class ExcelAddToDatabase
    Private Sub ExcelAddToDatabase_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim f1 As New List(Of Integer)

        For i As Integer = 2011 To 2018
            f1.Add(i)
        Next
        Dim f2 As New List(Of ExcelAddToDatabase)
        Dim f3 As New List(Of String)
        Dim f31 As New List(Of String)
        Dim f5 = Directory.GetFiles("c:\Времянка\", "*.xlsx", IO.SearchOption.TopDirectoryOnly).ToList()
        f31.AddRange(f5)
        For Each b In f1

            'Dim f51 = (Directory.GetFiles("Z:\RICKMANS\" & b, "*.xlsx", IO.SearchOption.TopDirectoryOnly).ToList())
            'Dim f52 = (Directory.GetFiles("Z:\RICKMANS\" & b, "*.xlsm", IO.SearchOption.TopDirectoryOnly).ToList())


            'f31.AddRange(f51)
            'f31.AddRange(f52)

        Next

        GeneralAsync(f31)




    End Sub
    Private Async Sub GeneralAsync(ByVal f31 As List(Of String))
        Await Task.Run(Sub() General(f31))
    End Sub
    Private Sub General(ByVal f31 As List(Of String))
        Dim f8 As New List(Of ФайлыExcelВсе)

        Dim f7 As List(Of ФайлыExcelВсе)
        Using db As New dbAllDataContext()
            f7 = db.ФайлыExcelВсе.Select(Function(x) x).ToList()
            'Dim f9 = db.ФайлыExcelВсе.Where(Function(x) x.ID > 147).Select(Function(x) x).ToList()
            'For Each b In f9
            '    b.Перевозчик = Nothing
            '    db.SubmitChanges()
            'Next
        End Using



        Try
            For Each b In f31
                If b.ToString.Contains("ОБРАЗЕЦ") Then Continue For

                Dim flinf As New FileInfo(b)

                Dim mdf As String = f7.Where(Function(x) x.Рейс = flinf.Name).Select(Function(x) x.Перевозчик).FirstOrDefault
                If String.IsNullOrEmpty(mdf) = False Then Continue For

                If b.Contains("101 Победа 1 - Бизунова 1.xlsx") Then Continue For
                If b.Contains("223 ТехноОснова 2 - Опалейчук Авто 1.xlsx") Then Continue For
                If b.Contains("012 Виларда Голд 33 - Куттер 10.xlsx") Then Continue For
                If b.Contains("023 Вудконцепт 3 -СиЛ 1.xlsx") Then Continue For
                If b.Contains("030 Порто-порто 37 - Белсотра 10.xlsx") Then Continue For
                If b.Contains("037 Виталюр 151 - Автороад 1.xlsx") Then Continue For
                If b.Contains("037 Вудконцепт 4 -СиЛ 2.xlsx") Then Continue For
                If b.Contains("039 Интерра Про 8 - Автовнештранс 1.xlsx") Then Continue For
                If b.Contains("048 Порто-порто 59 - Белсотра 14.xlsx") Then Continue For
                If b.Contains("059 Порто-порто 40 - Эркюль 11.xlsx") Then Continue For
                If b.Contains("061 Интерра Про 9 - Белтрансконсалт 4.xlsx") Then Continue For
                If b.Contains("066 Порто-порто 41 - Евротранстранзит 5.xlsx") Then Continue For
                If b.Contains("067 Интерра Про 10 - Слаин Авто 2.xlsx") Then Continue For
                If b.Contains("068 Порто порто 60 - Белсотра 15.xlsx") Then Continue For
                If b.Contains("071 Порто-порто 42 - Лукьянов 1.xlsx") Then Continue For
                If b.Contains("078 ММК-Белснаб 7 + Светоприбор 3 - Белтаможсервис 1.xlsx") Then Continue For
                If b.Contains("081 Интерра Про 1 - Куттер 4.xlsx") Then Continue For
                If b.Contains("081 Интерра Про 11 - Белсотра 11.xlsx") Then Continue For
                If b.Contains("084 Порто-порто 26 - Эркюль 10.xlsx") Then Continue For
                If b.Contains("087 Порто-порто 45 - Евротранстранзит 6.xlsx") Then Continue For
                If b.Contains("119 Kilimalis 2 - Белавтогаз 5.xlsx") Then Continue For
                If b.Contains("123 Светоприбор 4 - ГалТексАвто 2.xlsx") Then Continue For
                If b.Contains("140 ММК Белснаб 8 - АТлайн авто 1.xlsx") Then Continue For

                If b.Contains("188 Евроторг 86 - Кубоктранс 1.xlsx") Then Continue For
                If b.Contains("190 Евроторг 87 - Белтрансконсалт 2+ ПРОСТОЙ.xlsx") Then Continue For
                If b.Contains("191 Евроторг 90 - АМ Экспедиция 1.xlsx") Then Continue For
                If b.Contains("203 Вудконцепт 1 - Евротранстранзит 4.xlsx") Then Continue For
                If b.Contains("211 Евроторг 92 - Галенда 1.xlsx") Then Continue For
                If b.Contains("253 ММК Белснаб 4 + ИП Тимошенко 1 - Пост ТЭК 1.xlsx") Then Continue For
                If b.Contains("304 Порто-порто 21 - Белсотра 9 .xlsx") Then Continue For
                'If b.Contains("211 ММК-Белснаб 2 - Белсотра 8.xlsx") Then Continue For
                If b.Contains("Порто-порто") Then Continue For
                If b.Contains("Интерра Про") Then Continue For
                If b.Contains("ММК-Белснаб") Then Continue For
                If b.Contains("Вудконцепт") Then Continue For
                If b.Contains("Евроторг") Then Continue For

                Dim workbook As New XLWorkbook
                Try
                    workbook = New XLWorkbook(flinf.FullName)
                Catch ex As Exception
                    MessageBox.Show(ex.Message & " " & flinf.FullName)
                End Try
                Dim rows As IXLRangeRows

                Try
                    Dim worksheet = workbook.Worksheet("ПЕР")
                    rows = worksheet.RangeUsed().RowsUsed()

                Catch ex As Exception
                    Continue For
                End Try

                Dim rwNumber As Integer = 0
                Dim ml As New ФайлыExcelВсе
                ml.Рейс = flinf.Name.ToString
                Dim kld As String = Nothing

                If rows.Count > 0 Then
                    For Each b2 In rows

                        Dim gh = b2.Cell(1).Value
                        Try
                            If gh.ToString.ToUpper.Contains("Реквизиты перевозчика:".ToUpper) = True Then
                                rwNumber = b2.RowNumber
                                kld = b2.Cell(2).Value

                            End If

                            If b2.RowNumber = rwNumber + 1 Then
                                kld = kld & ", " & b2.Cell(2).Value

                            End If
                        Catch ex As Exception
                            Continue For
                        End Try
                        'If gh.ToString.ToUpper.Contains("Реквизиты перевозчика:".ToUpper) = True Then
                        '    rwNumber = b2.RowNumber
                        '    kld = b2.Cell(2).Value

                        'End If

                        'If gh.ToString.ToUpper.Contains("Реквизиты клиента:".ToUpper) = True Then
                        '    ml.Клиент = b2.Cell(2).Value
                        'End If

                        '    If b2.RowNumber = rwNumber + 1 Then
                        '    kld = kld & ", " & b2.Cell(2).Value

                        'End If

                        'If gh.ToString.ToUpper.Contains("Дата подачи под загрузку".ToUpper) = True Then
                        '    ml.ДатаЗагрузки = b2.Cell(2).Value

                        'End If

                        'If gh.ToString.ToUpper.Contains("Точный адрес загрузки, контактные лица, тел-ны:".ToUpper) = True Then
                        '    ml.АдресЗагрузки = b2.Cell(2).Value

                        'End If
                        'If gh.ToString.ToUpper.Contains("Адрес затаможки:".ToUpper) = True Then
                        '    ml.АдресЗатаможки = b2.Cell(2).Value

                        'End If

                        'If gh.ToString.ToUpper.Contains("Наименование, вес и количество груза:".ToUpper) = True Then
                        '    ml.Груз = b2.Cell(2).Value

                        'End If

                        'If gh.ToString.ToUpper.Contains("Маршрут:".ToUpper) = True Then
                        '    ml.Маршрут = b2.Cell(2).Value

                        'End If

                        'If gh.ToString.ToUpper.Contains("Дата подачи под растаможку:".ToUpper) = True Then
                        '    ml.ДатаПодРастаможку = b2.Cell(2).Value

                        'End If

                        'If gh.ToString.ToUpper.Contains("Точный адрес растаможки:".ToUpper) = True Then
                        '    ml.АдресРастаможки = b2.Cell(2).Value

                        'End If

                        'If gh.ToString.ToUpper.Contains("Точный адрес разгрузки:".ToUpper) = True Then
                        '    ml.АдресВыгрузки = b2.Cell(2).Value

                        'End If

                        'If gh.ToString.ToUpper.Contains("Дополнительные условия:".ToUpper) = True Then
                        '    ml.ДопУсловия = b2.Cell(2).Value

                        'End If

                        'If gh.ToString.ToUpper.Contains("ПОРУЧЕНИЕ № ".ToUpper) = True Then
                        '    ml.ДатаПоручения = b2.Cell(5).Value

                        'End If

                        ' Dim f7 = rows.Select(Function(x) New ФайлыExcelВсе With {.Авто = IIf(x.ToString.ToUpper.Contains("Тип, объем и тоннаж"), x.Cell("C").Value, Nothing)}).FirstOrDefault()

                    Next

                    Using db As New dbAllDataContext()
                        Dim f = db.ФайлыExcelВсе.Where(Function(x) x.Рейс = flinf.Name).Select(Function(x) x).FirstOrDefault
                        f.Перевозчик = kld
                        db.SubmitChanges()
                        'db.ФайлыExcelВсе.InsertOnSubmit(ml)
                        'db.SubmitChanges()
                    End Using
                End If

            Next
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try



        'If f8 IsNot Nothing Then
        '    Using db As New dbAllDataContext()
        '        For Each b In f8
        '            db.ФайлыExcelВсе.InsertOnSubmit(b)
        '            db.SubmitChanges()
        '        Next

        '    End Using
        'End If



    End Sub
    Private Async Sub gh(ByVal lst As List(Of String))
        Await Task.Run(Sub() OpenExSaveXlsx(lst))

    End Sub
    Private Sub OpenExSaveXlsx(ByVal lst As List(Of String))
        Dim xlapp As Microsoft.Office.Interop.Excel.Application
        xlapp = New Microsoft.Office.Interop.Excel.Application
        For Each b In lst
            Dim b1 = IO.Path.GetFileNameWithoutExtension(b)
            Dim b2 = Path.GetExtension(b)
            If b2 = ".xlsm" Then Continue For
            Dim fln = IO.Path.Combine("c:\Времянка", b1)
            Dim fln2 = fln & ".xlsx"
            If File.Exists(fln2) Then Continue For
            xlapp.DisplayAlerts = 0
            xlapp.Workbooks.Open(b)
            xlapp.ActiveWorkbook.SaveAs(fln, XlFileFormat.xlOpenXMLWorkbook, Type.Missing, Type.Missing, False, False, XlSaveAsAccessMode.xlShared,
                                        XlSaveConflictResolution.xlLocalSessionChanges, Type.Missing, Type.Missing)
            xlapp.DisplayAlerts = 1

        Next
        xlapp.Quit()
        releaseobject(xlapp)
    End Sub
    Public Sub releaseobject(ByVal obj As Object)
        Try
            Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub

End Class