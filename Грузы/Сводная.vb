Imports System.ComponentModel
Imports System.IO
Imports System.Threading
Imports ClosedXML.Excel

Public Class Сводная
    Dim com1 As New List(Of String)
    Private bscom1 As New BindingSource

    Dim Grid1All As List(Of Grid1Class)
    Private bsGrid1 As New BindingSource

    Dim xlapp As Microsoft.Office.Interop.Excel.Application
    Dim xlworkbook As Microsoft.Office.Interop.Excel.Workbook
    Dim xlworksheet As Microsoft.Office.Interop.Excel.Worksheet


    Sub New()

        ' Этот вызов является обязательным для конструктора.
        InitializeComponent()

        bscom1.DataSource = com1
        ComboBox1.DataSource = bscom1

        Grid1All = New List(Of Grid1Class)
        bsGrid1.DataSource = Grid1All
        Grid1.DataSource = bsGrid1
        GridView(Grid1)
        Grid1.Columns(3).HeaderText = "Сумма и дата оплаты"
        Grid1.Columns(4).HeaderText = "Сумма и дата поступления оплаты"
        Grid1.Columns(5).HeaderText = "Остаток"

        Grid1.Columns(7).HeaderText = "Сумма и дата оплаты"
        Grid1.Columns(8).HeaderText = "Сумма и дата поступления оплаты"
        Grid1.Columns(9).HeaderText = "Остаток"




        ' Добавить код инициализации после вызова InitializeComponent().
        predzagruzkaAsync()
    End Sub
    Private Sub Сводная_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim mo As New AllUpd
        Do While AllClass.РейсыКлиента Is Nothing
            mo.РейсыКлиентаAll()
        Loop


        Dim f = (From x In AllClass.РейсыКлиента
                 Select x.ДатаПоручения).Distinct()

        Dim m As New List(Of String)
        For Each b In f
            If b IsNot Nothing Then
                m.Add(CDate(b).Year)
            End If
        Next
        If com1 IsNot Nothing Then
            com1.Clear()
        End If

        com1.AddRange(m.OrderBy(Function(x) x).Distinct().ToList())
        bscom1.ResetBindings(False)
        ComboBox1.Text = String.Empty

        Grid1.Columns(0).Width = 60
        Grid1.Columns(1).Width = 80
        Grid1.Columns(3).Width = 100
        Grid1.Columns(4).Width = 100
        Grid1.Columns(5).Width = 100
        Grid1.Columns(7).Width = 100
        Grid1.Columns(8).Width = 100
        Grid1.Columns(9).Width = 100
        Grid1.Columns(10).Width = 100



    End Sub
    Private Async Sub predzagruzkaAsync()
        Await Task.Run(Sub() predzagruzka())
    End Sub
    Private Sub predzagruzka()
        Dim mo As New AllUpd
        Do While AllClass.РейсыКлиента Is Nothing
            mo.РейсыКлиентаAll()
        Loop

        Do While AllClass.РейсыПеревозчика Is Nothing
            mo.РейсыПеревозчикаAll()
        Loop

        Do While AllClass.ОплатыКлиент Is Nothing
            mo.ОплатыКлиентAll()
        Loop


        Do While AllClass.ОплатыПер Is Nothing
            mo.ОплатыПерAll()
        Loop


        Do While AllClass.СводнаяОплаты Is Nothing
            mo.СводнаяОплатыAll()
        Loop


        Do While AllClass.СводнаяОплатыТаблицы Is Nothing
            mo.СводнаяОплатыТаблицыAll()
        Loop
    End Sub
    Private Async Sub Com1sAsync(f As String, Optional H As Boolean = False)



        Dim mo As New AllUpd
        Do While AllClass.СводнаяОплаты Is Nothing
            mo.СводнаяОплатыAll()
        Loop


        Do While AllClass.СводнаяОплатыТаблицы Is Nothing
            mo.СводнаяОплатыТаблицыAll()
        Loop

        Dim f1 As New List(Of СводнаяОплатыТаблицы)

        Using db As New dbAllDataContext(_cn3)
            f1 = (From x In db.СводнаяОплаты
                  Join y In db.СводнаяОплатыТаблицы On x.ID Equals y.IDСоднаяОлпаты
                  Where x.Год = f And x.Состояние = "Нет изменения"
                  Select y).ToList()
        End Using


        Dim msd As New List(Of Grid1Class)
        'Dim txt4 As New Thread(Sub() UpdC(f, msd))

        If H = True Then
            'If txt4.IsAlive = True Then
            '    txt4.Join()
            'End If

            If f1 IsNot Nothing Then
                If f1.Count > 0 Then
                    MessageBox.Show("Нет изменения!", Рик)
                    Return
                End If
            End If
        End If



        If f1 IsNot Nothing Then
            If f1.Count > 0 Then
                If Grid1All IsNot Nothing Then
                    Grid1All.Clear()
                End If

                For Each b In f1
                    Dim b1 As New Grid1Class With {.Номер = b.Номер, .Рейс = b.Рейс, .Дельта = b.Дельта, .Клиент = b.Клиент, .ОстатокКлиент = b.ОстатокКлиент,
                        .ДатаОплатыИСуммаПоступленияКлиент = b.ДатаОплатыИСуммаПоступленияКлиент, .ДатаОплатыИСуммаПоступленияПеревозчик = b.ДатаОплатыИСуммаПоступленияПеревозчик,
                        .ОстатокПеревозчик = b.ОстатокПеревозчик, .Перевозчик = b.Перевозчик, .СуммаИДатаОплКлиент = b.СуммаИДатаОплКлиент, .СуммаИДатаОплПеревозчик = b.СуммаИДатаОплПеревозчик}
                    Grid1All.Add(b1)
                    'If b.Рейс = 722 Then
                    '    Dim mfg As String = String.Empty
                    'End If
                Next
                Grid1.BeginInvoke(New MethodInvoker(Sub() grid1Upd()))
            Else
                If Grid1All IsNot Nothing Then
                    Grid1All.Clear()
                End If

                msd = Await Task.Run(Function() Com1s(f))
                Grid1All.AddRange(msd)
                Grid1.BeginInvoke(New MethodInvoker(Sub() grid1Upd()))
                Await Task.Run(Sub() UpdC(f, msd))
                'Return
                'UpdCAsync(f, msd)
                'txt4.Start()


            End If
        Else
            If Grid1All IsNot Nothing Then
                Grid1All.Clear()
            End If
            msd = Await Task.Run(Function() Com1s(f))



            'txt4.Start()
            'UpdCAsync(f, msd)
            Grid1All.AddRange(msd)
            Grid1.BeginInvoke(New MethodInvoker(Sub() grid1Upd()))
            Await Task.Run(Sub() UpdC(f, msd))
            'Return
        End If

        'Grid1.BeginInvoke(New MethodInvoker(Sub() grid1Upd()))
        'grid1Upd()
    End Sub

    Private Sub grid1Upd()
        bsGrid1.ResetBindings(False)
        If Grid1.Rows.Count > 0 Then
            For Each b As DataGridViewRow In Grid1.Rows
                If b.Cells(5).Value > "0" Then
                    If b.Cells(3).Value.ToString.Length > 20 Then
                        Dim dt = CDate(Strings.Right(b.Cells(3).Value, 10))
                        If Not dt = Nothing Then
                            If dt > Now Then
                                b.Cells(3).Style.ForeColor = Color.Green
                            Else
                                b.Cells(3).Style.ForeColor = Color.Red
                            End If
                        End If
                    End If
                End If
                If b.Cells(9).Value > "0" Then
                    If b.Cells(7).Value.ToString.Length > 20 Then
                        Dim dt = CDate(Strings.Right(b.Cells(7).Value, 10))
                        If Not dt = Nothing Then
                            If dt > Now Then
                                b.Cells(7).Style.ForeColor = Color.Green
                            Else
                                b.Cells(7).Style.ForeColor = Color.Red
                            End If
                        End If
                    End If
                End If
            Next
        End If

    End Sub

    Private Function Com1s(f As String) As List(Of Grid1Class)
        Dim mo As New AllUpd
        Do While AllClass.РейсыКлиента Is Nothing
            mo.РейсыКлиентаAll()
        Loop

        Do While AllClass.РейсыПеревозчика Is Nothing
            mo.РейсыПеревозчикаAll()
        Loop


        mo.ОплатыКлиентAll()
        mo.ОплатыПерAll()


            Do While AllClass.СводнаяОплаты Is Nothing
            mo.СводнаяОплатыAll()
        Loop


        Do While AllClass.СводнаяОплатыТаблицы Is Nothing
            mo.СводнаяОплатыТаблицыAll()
        Loop

        Dim grd1 As New List(Of Grid1Class)

        Dim client As New List(Of РейсыКлиента)
        'Dim perevoz As New List(Of РейсыПеревозчика)
        'Dim Oplclient As New List(Of ОплатыКлиент)(AllClass.ОплатыКлиент.Where(Function(x) CDate(x.ДатаОплаты.ToString).Year = f).ToList())
        'Dim Oplperevoz As New List(Of ОплатыПер)(AllClass.ОплатыПер.Where(Function(x) CDate(x.ДатаОплаты.ToString).Year = f).ToList())



        For Each b In AllClass.РейсыКлиента
            If b.ДатаПоручения IsNot Nothing Then
                If CDate(b.ДатаПоручения).Year.ToString = f Then
                    client.Add(b)
                End If
            End If
        Next
        'For Each b In AllClass.РейсыПеревозчика
        '    If b.ДатаПоручения IsNot Nothing Then
        '        If CDate(b.ДатаПоручения).Year.ToString = f Then
        '            perevoz.Add(b)
        '        End If
        '    End If
        'Next

        Dim i As Integer = 1
        For Each b In client.OrderByDescending(Function(x) x.НомерРейса).ToList()
            Dim m1 As New Grid1Class

            Dim m = AllClass.ОплатыКлиент.Where(Function(x) x.Рейс = b.НомерРейса).ToList()
            Dim sum As Double = AllClass.ОплатыКлиент.Where(Function(x) x.Рейс = b.НомерРейса).Sum(Function(x) x.Сумма)

            Dim kl As String = Nothing


            If m IsNot Nothing Then
                If m.Count > 0 Then
                    For Each b1 In m

                        If kl Is Nothing Then
                            kl = b1.Сумма & vbCrLf & b1.ДатаОплаты
                        Else
                            kl = kl & vbCrLf & b1.Сумма & vbCrLf & b1.ДатаОплаты
                        End If

                    Next
                End If
            End If
            m1.Рейс = b.НомерРейса
            m1.ДатаОплатыИСуммаПоступленияКлиент = kl
            m1.Клиент = b.НазвОрганизации
            m1.СуммаИДатаОплКлиент = b.СтоимостьФрахта & " - (" & LCase(b.Валюта) & ")" & vbCrLf & IIf(b.ДатаОплаты Is Nothing, String.Empty, b.ДатаОплаты)

            If b.Валюта = "Рубль" Then 'тип валюты платежа
                m1.ОстатокКлиент = CType(b.СтоимостьФрахта, Double) - sum
            Else
                Dim валюта As String = b.Валюта
                Dim summa As Double = Nothing
                For Each b3 In m
                    Dim курс As Double = Nothing
                    Dim mR As New NbRbClassNew()
                    курс = Replace(mR.Курс(b3.ДатаОплаты, валюта), ".", ",")

                    'делим сумму прихода на курс получаем в валюте
                    Dim nk = Math.Round(b3.Сумма / курс, 2)
                    summa += nk
                Next
                If CType(b.СтоимостьФрахта, Double) = summa Then
                    m1.ОстатокКлиент = "0"
                End If
            End If


            'If b.НомерРейса = 586 Then
            '    Dim mas As Integer = 1
            'End If

            Dim p = AllClass.ОплатыПер.Where(Function(x) x.Рейс = b.НомерРейса).ToList()
            Dim sumPer As Double = AllClass.ОплатыПер.Where(Function(x) x.Рейс = b.НомерРейса).Sum(Function(x) x.Сумма)
            Dim perRe As РейсыПеревозчика = AllClass.РейсыПеревозчика.Where(Function(x) x.НомерРейса = b.НомерРейса).FirstOrDefault()
            Dim kl1 As String = Nothing
            If perRe Is Nothing Then Continue For
            If p IsNot Nothing Then
                If p.Count > 0 Then

                    For Each b1 In p
                        If kl1 Is Nothing Then
                            kl1 = b1.Сумма & vbCrLf & b1.ДатаОплаты
                        Else
                            kl1 = kl1 & vbCrLf & b1.Сумма & vbCrLf & b1.ДатаОплаты
                        End If

                    Next
                End If
            End If

            m1.ДатаОплатыИСуммаПоступленияПеревозчик = kl1
            m1.Перевозчик = perRe.НазвОрганизации
            m1.СуммаИДатаОплПеревозчик = perRe.СтоимостьФрахта & " - (" & LCase(perRe.Валюта) & ")" & vbCrLf & IIf(perRe.ДатаОплаты Is Nothing, String.Empty, perRe.ДатаОплаты)

            If b.НомерРейса = 451 Then
                Dim mnh As Integer = 1
            End If

            If perRe.Валюта = "Рубль" Then 'тип валюты платежа
                m1.ОстатокПеревозчик = CType(perRe.СтоимостьФрахта, Double) - sumPer
            Else
                Dim валюта As String = perRe.Валюта
                Dim summa As Double = Nothing
                For Each b4 In p
                    Dim курс As Double = Nothing
                    Dim mR As New NbRbClassNew()
                    курс = Replace(mR.Курс(b4.ДатаОплаты, валюта), ".", ",")

                    'делим сумму прихода на курс получаем в валюте
                    Dim nk = Math.Round(b4.Сумма / курс, 2)
                    summa += nk
                Next
                If CType(perRe.СтоимостьФрахта, Double) = summa Then
                    m1.ОстатокПеревозчик = "0"
                Else
                    Dim mR1 As New NbRbClassNew() 'для примерного расчет остатка по курсу на сегодня
                    Dim курс = Replace(mR1.Курс(Now, валюта), ".", ",")
                    Dim ost = CType(perRe.СтоимостьФрахта, Double) - summa
                    m1.ОстатокПеревозчик = ost & "(" & валюта & ")" & vbCrLf & (Math.Round(ost * курс, 2))
                End If
            End If

            If b.Валюта = "Рубль" And perRe.Валюта = "Рубль" Then
                m1.Дельта = CDbl(b.СтоимостьФрахта) - CDbl(perRe.СтоимостьФрахта)
            ElseIf Not b.Валюта = "Рубль" And perRe.Валюта = "Рубль" Then
                m1.Дельта = sum - CDbl(perRe.СтоимостьФрахта)
            ElseIf Not b.Валюта = "Рубль" And Not perRe.Валюта = "Рубль" Then
                m1.Дельта = CDbl(b.СтоимостьФрахта) - CDbl(perRe.СтоимостьФрахта)
            End If



            m1.Номер = i
            i += 1

            grd1.Add(m1)
        Next



        Return grd1
    End Function

    Private Async Sub UpdCAsync(f As String, grd1 As List(Of Grid1Class))
        Await Task.Run(Sub() UpdC(f, grd1))
    End Sub
    Private Sub UpdC(f As String, grd1 As List(Of Grid1Class))

        Dim mo As New AllUpd
        Using db As New dbAllDataContext(_cn3)
            Dim idYear = db.СводнаяОплаты.Where(Function(x) x.Год = f).Select(Function(x) x).FirstOrDefault()
            If idYear IsNot Nothing Then  'если год  внесен в базу

                idYear.Состояние = "Нет изменения"
                idYear.Дата_Изменения = Now
                db.SubmitChanges()


                Dim m = db.СводнаяОплатыТаблицы.Where(Function(x) x.IDСоднаяОлпаты = idYear.ID).ToList()
                If m IsNot Nothing Then
                    db.СводнаяОплатыТаблицы.DeleteAllOnSubmit(m)
                    db.SubmitChanges()
                End If
                Dim mfColl As New List(Of СводнаяОплатыТаблицы)
                For Each b In grd1

                    Dim mf As New СводнаяОплатыТаблицы With {
                        .IDСоднаяОлпаты = idYear.ID,
                        .ДатаОплатыИСуммаПоступленияКлиент = b.ДатаОплатыИСуммаПоступленияКлиент,
                        .ДатаОплатыИСуммаПоступленияПеревозчик = b.ДатаОплатыИСуммаПоступленияПеревозчик,
                        .Дельта = b.Дельта,
                        .Клиент = b.Клиент,
                        .Номер = b.Номер,
                        .ОстатокКлиент = b.ОстатокКлиент,
                        .ОстатокПеревозчик = b.ОстатокПеревозчик,
                        .Перевозчик = b.Перевозчик,
                        .Рейс = b.Рейс,
                        .СуммаИДатаОплКлиент = b.СуммаИДатаОплКлиент,
                        .СуммаИДатаОплПеревозчик = b.СуммаИДатаОплПеревозчик}
                    mfColl.Add(mf)

                Next
                db.СводнаяОплатыТаблицы.InsertAllOnSubmit(mfColl)
                db.SubmitChanges()



                mo.СводнаяОплатыТаблицыAll()
            Else
                'если год не внесен в базу
                Dim ff As New СводнаяОплаты
                ff.Год = Trim(f)
                ff.Состояние = "Нет изменения"
                ff.Дата_Изменения = Now
                db.СводнаяОплаты.InsertOnSubmit(ff)
                db.SubmitChanges()


                Dim mfColl As New List(Of СводнаяОплатыТаблицы)
                For Each b In grd1

                    Dim mf As New СводнаяОплатыТаблицы With {
                        .IDСоднаяОлпаты = ff.ID,
                        .ДатаОплатыИСуммаПоступленияКлиент = b.ДатаОплатыИСуммаПоступленияКлиент,
                        .ДатаОплатыИСуммаПоступленияПеревозчик = b.ДатаОплатыИСуммаПоступленияПеревозчик,
                        .Дельта = b.Дельта,
                        .Клиент = b.Клиент,
                        .Номер = b.Номер,
                        .ОстатокКлиент = b.ОстатокКлиент,
                        .ОстатокПеревозчик = b.ОстатокПеревозчик,
                        .Перевозчик = b.Перевозчик,
                        .Рейс = b.Рейс,
                        .СуммаИДатаОплКлиент = b.СуммаИДатаОплКлиент,
                        .СуммаИДатаОплПеревозчик = b.СуммаИДатаОплПеревозчик}
                    mfColl.Add(mf)

                Next
                db.СводнаяОплатыТаблицы.InsertAllOnSubmit(mfColl)
                db.SubmitChanges()



                mo.СводнаяОплатыТаблицыAll()
                mo.СводнаяОплатыAll()
            End If
        End Using

    End Sub

    Private Sub ComboBox1_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles ComboBox1.SelectionChangeCommitted
        Dim f As String = ComboBox1.SelectedItem

        Com1sAsync(Trim(f))


    End Sub
    Public Class Grid1Class
        Public Property Номер As String
        Public Property Рейс As String
        Public Property Клиент As String
        Public Property СуммаИДатаОплКлиент As String
        Public Property ДатаОплатыИСуммаПоступленияКлиент As String
        Public Property ОстатокКлиент As String
        Public Property Перевозчик As String
        Public Property СуммаИДатаОплПеревозчик As String
        Public Property ДатаОплатыИСуммаПоступленияПеревозчик As String
        Public Property ОстатокПеревозчик As String
        Public Property Дельта As String


    End Class
    Private ms1 As New List(Of Grid1Class)
    Private lfk As New List(Of Grid1Class)
    Private ms2 As New List(Of Grid1Class)
    Private lfk2 As New List(Of Grid1Class)

    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click
        If TextBox1.Text.Length = 0 Then Return
        Grid1.BeginInvoke(New MethodInvoker(Sub() PicBox1Click()))
    End Sub
    Private Sub PicBox1Click()

        If lfk Is Nothing Then Return
        Dim m

        For Each b In lfk
            If ms1.Contains(b) Then
                If ms1.Count = lfk.Count Then 'если сделали круг и начали заново
                    ms1.Clear()
                    m = Grid1All.IndexOf(b)
                    Grid1.ClearSelection()
                    Grid1.Rows(m).Selected = True
                    Grid1.FirstDisplayedScrollingRowIndex = m
                    ms1.Add(b)
                    Exit For
                Else
                    Continue For
                End If

            Else
                m = Grid1All.IndexOf(b)
                Grid1.ClearSelection()
                Grid1.Rows(m).Selected = True
                Grid1.FirstDisplayedScrollingRowIndex = m
                ms1.Add(b)
                Exit For
            End If

        Next
    End Sub





    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        ms1.Clear()
        lfk.Clear()
        If Grid1All IsNot Nothing Then
            lfk = Grid1All.Where(Function(x) x?.Клиент?.ToUpper.Contains(TextBox1.Text.ToUpper)).ToList()
        End If

    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        ms2.Clear()
        lfk2.Clear()
        If Grid1All IsNot Nothing Then
            lfk2 = Grid1All.Where(Function(x) x?.Перевозчик?.ToUpper.Contains(TextBox2.Text.ToUpper)).ToList()
        End If

    End Sub

    Private Sub PictureBox2_Click(sender As Object, e As EventArgs) Handles PictureBox2.Click
        If TextBox2.Text.Length = 0 Then Return
        Grid1.BeginInvoke(New MethodInvoker(Sub() PicBox2Click()))

    End Sub
    Private Sub PicBox2Click()
        If lfk2 Is Nothing Then Return

        Dim m

        For Each b In lfk2
            If ms2.Contains(b) Then
                If ms2.Count = lfk2.Count Then 'если сделали круг и начали заново
                    ms2.Clear()
                    m = Grid1All.IndexOf(b)
                    Grid1.ClearSelection()
                    Grid1.Rows(m).Selected = True
                    Grid1.FirstDisplayedScrollingRowIndex = m
                    ms2.Add(b)
                    Exit For
                Else
                    Continue For
                End If

            Else
                m = Grid1All.IndexOf(b)
                Grid1.ClearSelection()
                Grid1.Rows(m).Selected = True
                Grid1.FirstDisplayedScrollingRowIndex = m
                ms2.Add(b)
                Exit For
            End If

        Next
    End Sub

    Private Sub PictureBox3_Click(sender As Object, e As EventArgs) Handles PictureBox3.Click
        If TextBox3.Text.Length = 0 Then Return
        If Grid1All Is Nothing Then Return
        Dim m = Grid1All.IndexOf(Grid1All.Where(Function(x) x?.Рейс = TextBox3.Text).FirstOrDefault())
        If m > 0 Then
            Grid1.ClearSelection()
            Grid1.Rows(m).Selected = True
            Grid1.FirstDisplayedScrollingRowIndex = m
        End If

    End Sub

    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            If TextBox1.Text.Length = 0 Then Return
            Grid1.BeginInvoke(New MethodInvoker(Sub() PicBox1Click()))

        End If
    End Sub

    Private Sub TextBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox2.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            If TextBox2.Text.Length = 0 Then Return
            Grid1.BeginInvoke(New MethodInvoker(Sub() PicBox2Click()))

        End If
    End Sub

    Private Sub PictureBox4_Click(sender As Object, e As EventArgs)
        If ComboBox1.SelectedItem Is Nothing Then
            MessageBox.Show("Выберите год для обновления!", Рик)
            Return
        End If

        Dim f As String = ComboBox1.SelectedItem

        Com1sAsync(Trim(f), True)


    End Sub

    Private Sub Grid1_DataError(sender As Object, e As DataGridViewDataErrorEventArgs) Handles Grid1.DataError
        MessageBox.Show("Error happened " _
        & e.Context.ToString())

        If (e.Context = DataGridViewDataErrorContexts.Commit) _
        Then
            MessageBox.Show("Commit error")
        End If
        If (e.Context = DataGridViewDataErrorContexts _
        .CurrentCellChange) Then
            MessageBox.Show("Cell change")
        End If
        If (e.Context = DataGridViewDataErrorContexts.Parsing) _
        Then
            MessageBox.Show("parsing error")
        End If
        If (e.Context =
        DataGridViewDataErrorContexts.LeaveControl) Then
            MessageBox.Show("leave control error")
        End If

        If (TypeOf (e.Exception) Is ConstraintException) Then
            Dim view As DataGridView = CType(sender, DataGridView)
            view.Rows(e.RowIndex).ErrorText = "an error"
            view.Rows(e.RowIndex).Cells(e.ColumnIndex) _
            .ErrorText = "an error"

            e.ThrowException = False
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If ComboBox1.Text = String.Empty Then
            MessageBox.Show("Выберите год для обновления!", Рик)
            Return
        End If
        Cursor = Cursors.WaitCursor
        Dim f As String = ComboBox1.SelectedItem

        Com1sAsync(Trim(f), True)
        Cursor = Cursors.Default
    End Sub
    Private Sub ExtrExcel()
        If Grid1All Is Nothing Then
            Return
        End If

        Me.Cursor = Cursors.WaitCursor

        Dim Pathвремянка2 = "C:\Users\Public\Downloads"
        If File.Exists(Pathвремянка2 & "\" & "Oplaty2.xlsx") Then
            File.Delete(Pathвремянка2 & "\" & "Oplaty2.xlsx")
        End If
        If File.Exists(Pathвремянка2 & "\" & "Oplaty.xlsx") = False Then
            'копируем из ресурсов в папку general
            File.WriteAllBytes(Pathвремянка2 & "\" & "Oplaty.xlsx", My.Resources.Report)
        Else
            File.Delete(Pathвремянка2 & "\" & "Oplaty.xlsx")
            File.WriteAllBytes(Pathвремянка2 & "\" & "Oplaty.xlsx", My.Resources.Report)
        End If
        Dim FileВремянка2 = Pathвремянка2 & "\" & "Oplaty.xlsx"


        Dim wb As New XLWorkbook(FileВремянка2)
        Dim wc = wb.Worksheet("Оплаты")


        wc.Cell(2, 2).Value = Now.ToShortDateString

        Dim il As Integer = 1
        For Each b In Grid1All
            With wc
                .Cell(4 + il, 1).Value = b.Номер
                .Cell(4 + il, 2).Value = b.Рейс
                .Cell(4 + il, 3).Value = b.Клиент
                .Cell(4 + il, 4).Value = b.СуммаИДатаОплКлиент
                .Cell(4 + il, 5).Value = b.ДатаОплатыИСуммаПоступленияКлиент
                .Cell(4 + il, 6).Value = b.ОстатокКлиент
                .Cell(4 + il, 7).Value = b.Перевозчик
                .Cell(4 + il, 8).Value = b.СуммаИДатаОплПеревозчик
                .Cell(4 + il, 9).Value = b.ДатаОплатыИСуммаПоступленияПеревозчик
                .Cell(4 + il, 10).Value = b.ОстатокПеревозчик
                .Cell(4 + il, 11).Value = b.Дельта
                il += 1
            End With
        Next

        Dim rngTable = wc.Range("A4:K" & Grid1All.Count + 5)
        rngTable.Style.Border.RightBorder = XLBorderStyleValues.Thin
        rngTable.Style.Border.BottomBorder = XLBorderStyleValues.Thin


        Try
            Using db As New FileStream("Excel4.xlsx", FileMode.Create)
                wb.SaveAs(db)
            End Using
            Me.Cursor = Cursors.Default
            Process.Start(IO.Path.Combine(Application.StartupPath, "Excel4.xlsx"))
        Catch ex As Exception
            MessageBox.Show("Закройте ранее созданный экземпляр Excel", Рик)
            Me.Cursor = Cursors.Default
        End Try






        'Dim i, j As Integer 'сохранение в эксель

        'Dim misvalue As Object = Reflection.Missing.Value
        'xlapp = New Microsoft.Office.Interop.Excel.Application
        'xlworkbook = xlapp.Workbooks.Add(FileВремянка2)
        'xlworksheet = xlworkbook.Sheets("Оплаты")


        'xlworksheet.Cells(2, 2) = Now


        ''Dim il As Integer = 1
        ''For Each b In Grid1All
        ''    xlworksheet.Cells(4 + il, 1) = b.Номер
        ''    xlworksheet.Cells(4 + il, 2) = b.Рейс
        ''    xlworksheet.Cells(4 + il, 3) = b.Клиент
        ''    xlworksheet.Cells(4 + il, 4) = b.СуммаИДатаОплКлиент
        ''    xlworksheet.Cells(4 + il, 5) = b.ДатаОплатыИСуммаПоступленияКлиент
        ''    xlworksheet.Cells(4 + il, 6) = b.ОстатокКлиент
        ''    xlworksheet.Cells(4 + il, 7) = b.Перевозчик
        ''    xlworksheet.Cells(4 + il, 8) = b.СуммаИДатаОплПеревозчик
        ''    xlworksheet.Cells(4 + il, 9) = b.ДатаОплатыИСуммаПоступленияПеревозчик
        ''    xlworksheet.Cells(4 + il, 10) = b.ОстатокПеревозчик
        ''    xlworksheet.Cells(4 + il, 11) = b.Дельта
        ''    il += 1
        ''Next



        'For i = 0 To Grid1.Rows.Count - 1
        '    For j = 1 To Grid1.ColumnCount

        '        xlworksheet.Cells(i + 5, j) = Grid1(j - 1, i).Value

        '    Next
        'Next
        'xlworksheet.Range(Cell1:="A4", Cell2:="K" & Grid1.Rows.Count + 5).Cells.Borders.LineStyle = 1 'рисуем границы

        'xlworksheet.SaveAs("C:\Users\Public\Downloads\Oplaty2.xlsx")


        'xlworksheet.PrintOutEx()
        'xlworkbook.Close()
        'xlapp.Quit()
        'Me.Cursor = Cursors.Default
        'releaseobjectAsync(xlapp, xlworkbook, xlworksheet)




    End Sub
    Private Async Sub releaseobjectAsync(ByVal obj As Object, obj1 As Object, obj2 As Object)
        Await Task.Run(Sub() releaseobject(obj))
        Await Task.Run(Sub() releaseobject(obj1))
        Await Task.Run(Sub() releaseobject(obj2))
    End Sub
    Private Sub releaseobject(ByVal obj As Object)
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
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        ExtrExcel()
    End Sub

End Class