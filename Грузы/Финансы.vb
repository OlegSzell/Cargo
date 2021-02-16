Public Class Финансы
    Public grid1all As New List(Of AppEFC.Finans)
    Public bsgrid1 As New BindingSource

    Private Grid1all_2 As New List(Of Grid1Class)
    Public bsgrid1_2 As New BindingSource
    Private Sub Label9_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs)
        'Using db As New AppEFC
        '    Dim f As New AppEFC.Finans
        '    f.Dates = Now
        '    f.Клиент = "Виталюр"
        '    f.Перевозчик = "РТЛТранс"
        '    f.ОплаченоКлиентом = "1500"
        '    f.ОплаченоПеревозчику = "1300"
        '    f.Остаток = CType(f.ОплаченоКлиентом, Integer) - CType(f.ОплаченоПеревозчику, Integer)
        '    f.СтавкаПеревозчика = "1300"
        '    f.СуммаПеревозки = "1500"

        '    db.Finanss.Add(f)
        '    db.SaveChanges()

        '    Dim f1 = db.Finanss.ToList()
        '    grid1all.AddRange(f1)
        '    bsgrid1.ResetBindings(False)
        'End Using
    End Sub

    Private Sub Финансы_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'bsgrid1.DataSource = grid1all
        'Grid1.DataSource = bsgrid1


        bsgrid1_2.DataSource = Grid1all_2
        Grid1.DataSource = bsgrid1_2
        GridView(Grid1)

        Grid1.Columns(0).Width = 60
        Grid1.Columns(1).Width = 80
        Grid1.Columns(3).Width = 100
        Grid1.Columns(4).Width = 100
        Grid1.Columns(5).Width = 100
        Grid1.Columns(7).Width = 100
        Grid1.Columns(8).Width = 100
        Grid1.Columns(9).Width = 100
        Grid1.Columns(10).Width = 100
        Grid1.Columns(11).Visible = False
        Grid1.Columns(12).Visible = False

        Grid1.Columns(3).HeaderText = "Сумма и дата планируемой оплаты"
        Grid1.Columns(4).HeaderText = "Сумма и дата оплаты"
        Grid1.Columns(5).HeaderText = "Остаток"

        Grid1.Columns(7).HeaderText = "Сумма и дата планируемой оплаты"
        Grid1.Columns(8).HeaderText = "Сумма и дата оплаты"
        Grid1.Columns(9).HeaderText = "Остаток"
        MaskedTextBox1.Text = Now

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

        Public Property СуммаОплатКлиента As String
        Public Property СуммаОплатПеревозчику As String


    End Class


    Private Function Com1s(Optional f As Integer = 550, Optional Dat As Date = Nothing) As List(Of Grid1Class)
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

        If Dat = Nothing Then
            Dat = Now
        End If


        Dim grd1 As New List(Of Grid1Class)

        Dim client As New List(Of РейсыКлиента)
        Dim ds = (From x In AllClass.РейсыКлиента
                  Join y In AllClass.ОплатыКлиент On x.Код Equals y.IDКлиента
                  Order By x.НомерРейса Descending
                  Where x.НомерРейса > f And y.ДатаОплаты <= CDate(Dat)
                  Select x).ToList()



        Dim ds7 = (From x In AllClass.РейсыПеревозчика
                   Join y In AllClass.ОплатыПер On x.Код Equals y.IDПер
                   Order By x.НомерРейса Descending
                   Where x.НомерРейса > f And y.ДатаОплаты <= CDate(Dat)
                   Select x).ToList()


        Dim ds8 = ds.Select(Function(x) x.НомерРейса).Union(ds7.Select(Function(x) x.НомерРейса)).OrderBy(Function(x) x).ToList()

        For Each b In ds8
            Dim fb = AllClass.РейсыКлиента.Where(Function(x) x.НомерРейса?.ToString = b).Select(Function(x) x).FirstOrDefault()
            client.Add(fb)
        Next






        Dim i As Integer = 1
        For Each b In client
            Dim m1 As New Grid1Class  'Клиент
            Dim m As New List(Of ОплатыКлиент)
            Dim sum As Double

            m = AllClass.ОплатыКлиент.Where(Function(x) x.Рейс = b.НомерРейса And x.ДатаОплаты <= CDate(Dat)).ToList()
            sum = AllClass.ОплатыКлиент.Where(Function(x) x.Рейс = b.НомерРейса).Sum(Function(x) x.Сумма)
            If b.НомерРейса = 591 Then
                Dim mfdv As Integer = 0
            End If


            Dim kl As String = Nothing
            Dim kl2 As String = Nothing

            If m IsNot Nothing Then
                If m.Count > 0 Then
                    For Each b1 In m

                        If kl Is Nothing Then
                            kl = b1.Сумма & vbCrLf & "-" & b1.ДатаОплаты
                            kl2 = b1.Сумма
                        Else
                            kl = kl & vbCrLf & b1.Сумма & vbCrLf & "-" & b1.ДатаОплаты
                            kl2 = CType(Math.Round(CType(kl2, Double) + CType(b1.Сумма, Double)), String)
                        End If

                    Next
                End If
            End If
            m1.Рейс = b.НомерРейса
            m1.ДатаОплатыИСуммаПоступленияКлиент = kl
            m1.Клиент = b.НазвОрганизации
            m1.СуммаИДатаОплКлиент = b.СтоимостьФрахта & " - (" & LCase(b.Валюта) & ")" & vbCrLf & IIf(b.ДатаОплаты Is Nothing, String.Empty, b.ДатаОплаты)
            m1.СуммаОплатКлиента = kl2

            If b.Валюта = "Рубль" Then 'тип валюты платежа
                If m.Count > 0 Then
                    m1.ОстатокКлиент = CType(b.СтоимостьФрахта, Double) - sum
                Else
                    m1.ОстатокКлиент = b.СтоимостьФрахта
                End If

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

            Dim p = (From x In AllClass.РейсыПеревозчика        'Перевозчик
                     Join y In AllClass.ОплатыПер On x.Код Equals y.IDПер
                     Where x.НомерРейса = b.НомерРейса And y.ДатаОплаты <= CDate(Dat)
                     Select x, y).ToList()


            Dim sumPer As Double = p.Select(Function(x) x.y.Сумма).Sum(Function(x) x)
            Dim perRe As РейсыПеревозчика = AllClass.РейсыПеревозчика.Where(Function(x) x.НомерРейса = b.НомерРейса).FirstOrDefault()

            Dim kl1 As String = Nothing
            Dim kl3 As String = Nothing
            If perRe Is Nothing Then Continue For

            If p IsNot Nothing Then
                If p.Count > 0 Then

                    For Each b1 In p
                        If kl1 Is Nothing Then
                            kl1 = b1.y.Сумма & vbCrLf & "-" & b1.y.ДатаОплаты
                            kl3 = b1.y.Сумма
                        Else
                            kl1 = kl1 & vbCrLf & b1.y.Сумма & vbCrLf & "-" & b1.y.ДатаОплаты
                            kl3 = CType(Math.Round(CType(kl3, Double) + CType(b1.y.Сумма, Double)), String)
                        End If

                    Next
                End If
            End If

            m1.ДатаОплатыИСуммаПоступленияПеревозчик = kl1
            m1.Перевозчик = perRe.НазвОрганизации
            m1.СуммаИДатаОплПеревозчик = perRe.СтоимостьФрахта & " - (" & LCase(perRe.Валюта) & ")" & vbCrLf & IIf(perRe.ДатаОплаты Is Nothing, String.Empty, perRe.ДатаОплаты)
            m1.СуммаОплатПеревозчику = kl3

            If perRe.Валюта = "Рубль" Then 'тип валюты платежа
                m1.ОстатокПеревозчик = CType(perRe.СтоимостьФрахта, Double) - sumPer
            Else
                Dim валюта As String = perRe.Валюта
                Dim summa As Double = Nothing
                For Each b4 In p
                    Dim курс As Double = Nothing
                    Dim mR As New NbRbClassNew()
                    курс = Replace(mR.Курс(b4.y.ДатаОплаты, валюта), ".", ",")

                    'делим сумму прихода на курс получаем в валюте
                    Dim nk = Math.Round(b4.y.Сумма / курс, 2)
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

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If MaskedTextBox1.MaskCompleted = False Then
            MessageBox.Show("Заполните поле дата!", Рик)
            Return
        End If

        Cursor = Cursors.WaitCursor

        If Grid1all_2 IsNot Nothing Then
            Grid1all_2.Clear()
        End If
        Dim grd1V As New List(Of Grid1Class)
        grd1V.AddRange(Com1s(, CDate(MaskedTextBox1.Text).ToShortDateString).OrderByDescending(Function(x) x.Рейс).ToList())


        Dim grd2 As New List(Of Grid1Class) 'отбираем при необходимости только без нулевых оплат

        If CheckBox1.Checked = False Then
            Dim mt = grd1V.Where(Function(x) x.ОстатокКлиент = "0" And x.ОстатокПеревозчик = "0").Select(Function(x) x.Рейс).ToList()
            Dim mt1 = grd1V.Select(Function(x) x.Рейс).ToList().Except(mt).ToList()
            For Each b In mt1
                grd2.Add(grd1V.Where(Function(x) x.Рейс = b).FirstOrDefault())
            Next
        Else
            grd2.AddRange(grd1V)

        End If

        Dim mk As Double = 0, ml As Double = 0, txt6 As Double = 0, txt1 As Double = 0
        Dim i As Integer = 1
        'For Each b In grd2
        '    If CType(b.Дельта, Double) > 0 Then
        '        txt6 += CType(b.Дельта, Double)
        '    Else
        '        txt6 += 100
        '    End If

        '    mk = CType(mk, Double) + CType(b.СуммаОплатКлиента, Double)
        '    ml = CType(ml, Double) + CType(b.СуммаОплатПеревозчику, Double)
        '    b.Номер = i
        '    Grid1all_2.Add(b)
        '    i += 1
        'Next
        Dim OplMy As Double = 0
        Dim Ubtk As Double = 0
        Dim RaschOst As Double = 0
        For Each b In grd2
            If CType(b.Дельта, Double) > 0 Then
                txt6 += CType(b.Дельта, Double)
            Else
                txt6 += 100
            End If

            If CDbl(b.ОстатокПеревозчик) > 0 And CDbl(b.ОстатокКлиент) = 0 Then
                Ubtk += CDbl(b.ОстатокПеревозчик)
            ElseIf CDbl(b.ОстатокПеревозчик) > 0 And CDbl(b.ОстатокКлиент) > 0 Then 'если клиент не оплатил и мы перевозчику что то оплатили
                If CDbl(b.СуммаОплатКлиента) = 0 Then 'если клиент вообще не оплатил а мы оплатили частично перевозчику
                    OplMy += CDbl(b.СуммаОплатПеревозчику)
                ElseIf CDbl(b.СуммаОплатКлиента) > 0 And CDbl(b.ОстатокКлиент) > 0 Then 'частично оплатил клиент и частично полатили перевозчику
                    Ubtk += (CDbl(b.СуммаОплатКлиента) - CDbl(b.СуммаОплатПеревозчику))
                End If
            ElseIf CDbl(b.ОстатокПеревозчик) = 0 And CDbl(b.ОстатокКлиент) > 0 Then 'если мы полностью оплатили перевозчику , а клиент либо вообще не оплатил или оплатил частично
                If CDbl(b.СуммаОплатКлиента) = 0 Then 'если клиент вообще не оплатил а мы оплатили частично перевозчику
                    OplMy += CDbl(b.СуммаОплатПеревозчику)
                ElseIf CDbl(b.ОстатокКлиент) > 0 And CDbl(b.СуммаОплатКлиента) > 0 Then 'частично оплатил клиент и частично полатили перевозчику
                    Ubtk += (CDbl(b.СуммаОплатКлиента) - CDbl(b.СуммаОплатПеревозчику))
                End If
            ElseIf CDbl(b.ОстатокПеревозчик) = 0 And b.ОстатокКлиент = String.Empty Then
                OplMy += CDbl(b.СуммаОплатПеревозчику)

            End If



            mk = CType(mk, Double) + CType(b.СуммаОплатКлиента, Double)
            ml = CType(ml, Double) + CType(b.СуммаОплатПеревозчику, Double)
            b.Номер = i
            Grid1all_2.Add(b)
            i += 1
        Next
        RaschOst = Math.Round(Ubtk - OplMy, 0)


        bsgrid1_2.ResetBindings(False)

        If TextBox1.Text.Length = 0 Then
            txt1 = 0
        Else
            txt1 = CType(TextBox1.Text, Double)
        End If

        If TextBox5.Text = String.Empty Then
            TextBox5.Text = 0
        End If

        RaschOst = Math.Round((OplMy + (txt1 - CDbl(TextBox5.Text))) - Ubtk, 0)



        TextBox2.Text = Math.Round(mk)
        TextBox3.Text = Math.Round(ml)
        TextBox4.Text = Math.Round(mk) - (Math.Round(ml) + txt1)
        Label14.Text = Math.Round(txt6)
        Label15.Text = Math.Round(txt6) - ((Math.Round(txt6) * 40) / 100)

        Label13.Text = Grid1all_2.Count
        TextBox6.Text = RaschOst
        Cursor = Cursors.Default


    End Sub
End Class