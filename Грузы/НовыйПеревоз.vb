Option Explicit On
Imports System.Data.OleDb
Imports System.Threading
Public Class НовыйПеревоз

    Dim strsql As String
    Dim ds As DataTable
    Dim КодДляУдал As Перевозчики
    Dim кол As Integer

    Private com4all As List(Of Перевозчики)
    Private bscom4all As BindingSource
    Private listbx1 As List(Of Перевозчики)
    Private bslistbx1 As BindingSource
    Private com1all As List(Of ФормаСобств)
    Private bscom1all As BindingSource
    Private Async Sub ПредзагрузкаAsync()
        Await Task.Run(Sub() Предзагрузка())
    End Sub
    Private Sub Предзагрузка()
        Dim mo As New AllUpd
        mo.ПеревозчикиAll()
        mo.ФормаСобствAll()
    End Sub
    Private Sub НовыйПеревоз_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        ComboBox3.Text = "1"
        ПредзагрузкаAsync()
        'Me.ListBox1.Items.Clear()
        'For Each r As DataRow In ds1.Rows
        '    Me.ListBox1.Items.Add(r(0).ToString)
        'Next
        com4all = New List(Of Перевозчики)
        bscom4all = New BindingSource
        bscom4all.DataSource = com4all
        ComboBox4.DataSource = bscom4all
        ComboBox4.DisplayMember = "Названиеорганизации"
        ComboBox4.Text = String.Empty

        listbx1 = New List(Of Перевозчики)
        bslistbx1 = New BindingSource
        bslistbx1.DataSource = listbx1
        ListBox1.DataSource = bslistbx1
        ListBox1.DisplayMember = "Названиеорганизации"


        com1all = New List(Of ФормаСобств)
        bscom1all = New BindingSource
        bscom1all.DataSource = com1all
        ComboBox1.DataSource = bscom1all
        ComboBox1.DisplayMember = "ПолноеНазвание"
        ComboBox1.Text = String.Empty

        Dim mo As New AllUpd
        Do While AllClass.Перевозчики Is Nothing
            mo.ПеревозчикиAll()
        Loop
        Do While AllClass.ФормаСобств Is Nothing
            mo.ФормаСобствAll()
        Loop

        Dim f = AllClass.Перевозчики.OrderBy(Function(x) x.Названиеорганизации).Select(Function(x) x).ToList()
        If f IsNot Nothing Then
            listbx1.AddRange(f)
            com4all.AddRange(f)
            bscom4all.ResetBindings(False)
            bslistbx1.ResetBindings(False)
        End If

        com1all.AddRange(AllClass.ФормаСобств.OrderBy(Function(x) x.ПолноеНазвание).Select(Function(x) x).ToList())
        bscom1all.ResetBindings(False)

        'ComboBox1.AutoCompleteCustomSource.Clear()
        'ComboBox1.Items.Clear()

        'ComboBox4.Items.Clear()
        'ComboBox4.AutoCompleteCustomSource.Clear()
        'ListBox1.Items.Clear()
        'Using db As New dbAllDataContext(_cn3)
        '    Dim var = db.Перевозчики.OrderBy(Function(x) x.Названиеорганизации).Select(Function(x) x.Названиеорганизации).ToList
        '    If var.Count > 0 Then
        '        For Each k In var
        '            ListBox1.Items.Add(k)
        '            ComboBox4.Items.Add(k)
        '            ComboBox4.AutoCompleteCustomSource.Add(k)
        '        Next
        '    End If

        '    Dim var1 = db.ФормаСобств.OrderBy(Function(x) x.ПолноеНазвание).Select(Function(x) x.ПолноеНазвание).ToList()
        '    If var1.Count > 0 Then
        '        For Each f In var1
        '            ComboBox1.AutoCompleteCustomSource.Add(f)
        '            ComboBox1.Items.Add(f)
        '        Next
        '    End If
        'End Using



        'For Each r As DataRow In dtПеревозчики.Rows
        '    ListBox1.Items.Add(r(0).ToString)
        '    ComboBox4.Items.Add(r(0).ToString)
        '    ComboBox4.AutoCompleteCustomSource.Add(r(0).ToString)
        'Next

        'Dim c = From x In dtФормаСобствAll Order By x.Item("ПолноеНазвание").ToString Select x.Item("ПолноеНазвание")

        'ComboBox1.AutoCompleteCustomSource.Clear()
        'ComboBox1.Items.Clear()
        'For Each r In c
        '    ComboBox1.AutoCompleteCustomSource.Add(r.ToString)
        '    ComboBox1.Items.Add(r.ToString)
        'Next




        ''Dim f1 As New Thread(Sub() COMxt(Me, "SELECT ПолноеНазвание FROM ФормаСобств ORDER BY ПолноеНазвание", ComboBox1))
        ''f1.IsBackground = True
        ''f1.Start()

        MaskedTextBox1.Text = Now.ToShortDateString
    End Sub
    Private Sub richtextfont()
        'Dim Ctrl As Control

        'For Each Ctrl In Me.Controls 'перебираем текстбоксы вне tabcontrol и groupbox
        '    If TypeName(Ctrl) = "RichTextBox" Then
        '        arrtbox.Add(Ctrl.Name, Ctrl.Text)
        '        'Ctrl.Value = "бла-бла-бла"
        '    End If
        'Next


    End Sub

    Private Sub очистка()
        ComboBox1.Text = String.Empty
        RichTextBox1.Text = ""
        RichTextBox2.Text = ""
        RichTextBox3.Text = ""
        RichTextBox4.Text = ""
        RichTextBox5.Text = ""
        RichTextBox6.Text = ""
        RichTextBox7.Text = ""
        RichTextBox8.Text = ""
        RichTextBox9.Text = ""
        RichTextBox10.Text = ""
        RichTextBox11.Text = ""
        RichTextBox12.Text = ""
        RichTextBox13.Text = ""

    End Sub
    Private Sub ВставкаКлиент(ByVal ds As Перевозчики)


        ComboBox1.Text = ds.Форма_собственности
        RichTextBox1.Text = ds.Названиеорганизации
        RichTextBox2.Text = ds.Адрес_организации
        RichTextBox3.Text = ds.Почтовый_адрес
        RichTextBox4.Text = ds.РасчСчетЕвро
        RichTextBox5.Text = ds.РасчСчетДоллар
        RichTextBox6.Text = ds.РасчСчетРубли
        RichTextBox7.Text = ds.Контактное_лицо & " " & ds.Телефон
        RichTextBox8.Text = ds.Адрес_банка
        RichTextBox9.Text = ds.РасчСчетРоссРубли
        RichTextBox10.Text = ds.НаОснЧегоДейств
        RichTextBox12.Text = ds.ФИОРуководителя
        RichTextBox11.Text = ds.ФИОРодпадеж
        RichTextBox13.Text = ds.Должность
        RichTextBox14.Text = ds.ДолжРодПадеж
        If ds.ПерЭкспедитор = "Да" Then
            CheckBox2.Checked = True
        Else
            CheckBox2.Checked = False
        End If
    End Sub
    Private Sub ListBox1_Click(sender As Object, e As EventArgs) Handles ListBox1.Click
        ListBox1R()

    End Sub
    Private Sub ListBox1R()
        If Not ListBox1.SelectedIndex = -1 Then
            Try
                очистка()
                'Dim ds2 As DataTable = Selects3(StrSql:="SELECT * FROM Перевозчики WHERE Названиеорганизации='" & ListBox1.SelectedItem & "'")
                'Dim ds2 = dtПеревозчики.Select("Названиеорганизации='" & ListBox1.SelectedItem & "'")
                Dim ds2 As Перевозчики = ListBox1.SelectedItem
                ВставкаКлиент(ds2)
                КодДляУдал = ListBox1.SelectedItem
                ComboBox4.Text = String.Empty
            Catch ex As Exception

            End Try

        End If
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked = True Then
            очистка()
            CheckBox1.CheckState = False
        End If
    End Sub
    Private Function Проверка()
        If ComboBox1.Text = "" Or RichTextBox1.Text = "" Or RichTextBox2.Text = "" Or RichTextBox3.Text = "" Or RichTextBox8.Text = "" Or RichTextBox7.Text = "" Or RichTextBox13.Text = "" Or RichTextBox14.Text = "" Or RichTextBox12.Text = "" Or RichTextBox11.Text = "" Or RichTextBox10.Text = "" Then
            MessageBox.Show("Заполните поля отмеченные звездочкой!", Рик)
            Return 1
        End If

        If Not RichTextBox6.MaxLength = 28 Then
            MessageBox.Show("Проверьте количество знаков в поле 'Расчетный счет (BYN)'!", Рик)
            Return 1
        End If
        If RichTextBox6.Text = "" And RichTextBox5.Text = "" And RichTextBox4.Text = "" And RichTextBox9.Text = "" Then
            MessageBox.Show("Необходимо наличие хотя-бы одного расчетного счета обязательно!", Рик)
            Return 1
        End If

        If CheckBox4.Checked = True And ComboBox3.Text = "" Then
            MessageBox.Show("Необходимо выбрать количество экземпляров договора!", Рик)
            Return 1
        End If
        If CheckBox3.Checked = True Then
            If MessageBox.Show("Создать новый догововр с действующим перевозчиков?", Рик, MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.No Then
                Return 1
            End If
        End If

        Return 0
    End Function
    Private Sub ДокиПерЭксп()
        Dim расч As New List(Of String)() From {RichTextBox6.Text, RichTextBox5.Text, RichTextBox4.Text, RichTextBox9.Text}
        Dim готрасч As String = String.Empty
        For i As Integer = 0 To расч.Count - 1
            If расч(i) <> "" Then
                готрасч = готрасч & ", IBAN - " & расч(i)
            End If
        Next



        'Dim ds = dtПеревозчики.Select("Названиеорганизации='" & RichTextBox1.Text & "'")

        Dim nz As String = RichTextBox1.Text
        Dim ds = listbx1.Where(Function(x) x.Названиеорганизации = nz).Select(Function(x) x).FirstOrDefault()
        If ds Is Nothing Then Return

        Dim oWord As Microsoft.Office.Interop.Word.Application
        Dim oWordDoc As Microsoft.Office.Interop.Word.Document
        'KillProc()
        oWord = CreateObject("Word.Application")
        oWord.Visible = False

        Dim d As String = ""
        Select Case ComboBox2.Text
            Case "Евро"
                d = "евро"
            Case "Доллар"
                d = "доллар"
            Case "Российский рубль"
                d = "россрубл"
        End Select

        oWordDoc = oWord.Documents.Add("Z:\RICKMANS\ДОГОВОРА\ДОГОВОР ПЭ.doc")
        'Dim Files() As String = IO.Directory.GetFiles("Z:\RICKMANS\ДОГОВОРА", "*.txt", IO.SearchOption.TopDirectoryOnly)



        With oWordDoc.Bookmarks
            .Item("ПЭ1").Range.Text = ds.Договор
            .Item("ПЭ2").Range.Text = ds.Дата & "г."
            If ComboBox1.Text = "Индивидуальный предприниматель" Then
                .Item("ПЭ3").Range.Text = ComboBox1.Text & " " & RichTextBox1.Text
                .Item("ПЭ7").Range.Text = ComboBox1.Text
                .Item("ПЭ8").Range.Text = RichTextBox1.Text
            Else
                .Item("ПЭ3").Range.Text = ComboBox1.Text & " «" & RichTextBox1.Text & "»"
                .Item("ПЭ7").Range.Text = ComboBox1.Text
                .Item("ПЭ8").Range.Text = " «" & RichTextBox1.Text & "»"
            End If

            .Item("ПЭ4").Range.Text = Strings.LCase(RichTextBox14.Text)
            .Item("ПЭ5").Range.Text = RichTextBox11.Text
            .Item("ПЭ6").Range.Text = RichTextBox10.Text
            .Item("ПЭ9").Range.Text = RichTextBox2.Text & ", " & Trim(RichTextBox3.Text) & ", " & Trim(RichTextBox7.Text) & готрасч & ", " & Trim(RichTextBox8.Text)
            .Item("ПЭ10").Range.Text = RichTextBox13.Text
            .Item("ПЭ11").Range.Text = ФИОКорРук(RichTextBox12.Text, False)
            .Item("ПЭ12").Range.Text = РасчетнСчет(d)


            If Now > CDate("10.10.2021") Then
                .Item("ПЭ13").Range.Text = FaceNew
                .Item("ПЭ14").Range.Text = OsnNew
                .Item("ПЭ15").Range.Text = DoljNews
                .Item("ПЭ16").Range.Text = NameNews
                '.Item("ПЭ17").Range.Text = TypeNews

                'Else
                '    .Item("ПЭ13").Range.Text = FaceOld
                '    .Item("ПЭ14").Range.Text = OsnOld
                '    .Item("ПЭ15").Range.Text = DoljOld
                '    .Item("ПЭ16").Range.Text = NameOld
            End If

        End With


        If Not IO.Directory.Exists("Z:\RICKMANS\" & Now.Year & "\ДОГОВОР П\") Then
            IO.Directory.CreateDirectory("Z:\RICKMANS\" & Now.Year & "\ДОГОВОР П\")
        End If



        Try
            oWordDoc.SaveAs2("Z:\RICKMANS\" & Now.Year & "\ДОГОВОР П\ДОГОВОР ПЭ - " & кол & " " & RichTextBox1.Text & ".doc",,,,,, False)
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

        Dim СохрЗак As String = "Z:\RICKMANS\" & Now.Year & "\ДОГОВОР П\ДОГОВОР ПЭ - " & кол & " " & RichTextBox1.Text & ".doc"

        oWordDoc.Close(True)
        oWord.Quit(True)

        If MessageBox.Show("Печатать договор?", Рик, MessageBoxButtons.OKCancel, MessageBoxIcon.Question) = DialogResult.OK Then
            ПечатьДоков(СохрЗак, CType(ComboBox3.Text, Integer))
        End If






    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Cursor = Cursors.WaitCursor
        If Проверка() = 1 Then
            Me.Cursor = Cursors.Default
            Exit Sub
        End If

        Dim mo As New AllUpd

        Dim naz As String = RichTextBox1.Text
        Using db As New dbAllDataContext(_cn3)
            Dim f1 = db.Перевозчики.Where(Function(x) x.Названиеорганизации = naz).Select(Function(x) x).FirstOrDefault()

            Dim f As String
            Dim i As Integer

            If CheckBox2.Checked = True Then
                f = "Да"
                i = 1
            Else
                f = ""
                i = 0
            End If

            If f1 IsNot Nothing Then
                Dim nom2 As String = Nothing
                With f1
                    .Форма_собственности = ComboBox1.Text
                    .Адрес_организации = Trim(RichTextBox2.Text)
                    .Почтовый_адрес = Trim(RichTextBox3.Text)
                    .РасчСчетРубли = Trim(RichTextBox6.Text)
                    .РасчСчетРоссРубли = Trim(RichTextBox9.Text)
                    .РасчСчетДоллар = Trim(RichTextBox5.Text)
                    .РасчСчетЕвро = Trim(RichTextBox4.Text)
                    .Адрес_банка = Trim(RichTextBox8.Text)
                    .Контактное_лицо = Trim(RichTextBox7.Text)
                    .Должность = Trim(RichTextBox13.Text)
                    .НаОснЧегоДейств = Trim(RichTextBox10.Text)
                    .ФИОРуководителя = Trim(RichTextBox12.Text)
                    .ФИОРодпадеж = Trim(RichTextBox11.Text)
                    .ДолжРодПадеж = Trim(RichTextBox14.Text)
                    .ПерЭкспедитор = f

                    If CheckBox3.Checked = True Then
                        nom2 = НомерДог(i)
                        .Договор = nom2
                        .Дата = MaskedTextBox1.Text
                    End If
                End With
                db.SubmitChanges()



                'обновляем листбокс и комбобокс
                Dim f2 = listbx1.ElementAt(ListBox1.SelectedIndex)
                If f2 IsNot Nothing Then
                    With f2
                        .Форма_собственности = ComboBox1.Text
                        .Адрес_организации = Trim(RichTextBox2.Text)
                        .Почтовый_адрес = Trim(RichTextBox3.Text)
                        .РасчСчетРубли = Trim(RichTextBox6.Text)
                        .РасчСчетРоссРубли = Trim(RichTextBox9.Text)
                        .РасчСчетДоллар = Trim(RichTextBox5.Text)
                        .РасчСчетЕвро = Trim(RichTextBox4.Text)
                        .Адрес_банка = Trim(RichTextBox8.Text)
                        .Контактное_лицо = Trim(RichTextBox7.Text)
                        .Должность = Trim(RichTextBox13.Text)
                        .НаОснЧегоДейств = Trim(RichTextBox10.Text)
                        .ФИОРуководителя = Trim(RichTextBox12.Text)
                        .ФИОРодпадеж = Trim(RichTextBox11.Text)
                        .ДолжРодПадеж = Trim(RichTextBox14.Text)
                        .ПерЭкспедитор = f

                        If CheckBox3.Checked = True Then
                            .Договор = nom2
                            .Дата = MaskedTextBox1.Text
                        End If
                    End With

                    bslistbx1.ResetBindings(False)

                End If
                Dim f3 = com4all.ElementAt(ListBox1.SelectedIndex)
                With f3

                    .Форма_собственности = ComboBox1.Text
                    .Адрес_организации = Trim(RichTextBox2.Text)
                    .Почтовый_адрес = Trim(RichTextBox3.Text)
                    .РасчСчетРубли = Trim(RichTextBox6.Text)
                    .РасчСчетРоссРубли = Trim(RichTextBox9.Text)
                    .РасчСчетДоллар = Trim(RichTextBox5.Text)
                    .РасчСчетЕвро = Trim(RichTextBox4.Text)
                    .Адрес_банка = Trim(RichTextBox8.Text)
                    .Контактное_лицо = Trim(RichTextBox7.Text)
                    .Должность = Trim(RichTextBox13.Text)
                    .НаОснЧегоДейств = Trim(RichTextBox10.Text)
                    .ФИОРуководителя = Trim(RichTextBox12.Text)
                    .ФИОРодпадеж = Trim(RichTextBox11.Text)
                    .ДолжРодПадеж = Trim(RichTextBox14.Text)
                    .ПерЭкспедитор = f

                    If CheckBox3.Checked = True Then
                        .Договор = nom2
                        .Дата = MaskedTextBox1.Text
                    End If
                End With
                bscom4all.ResetBindings(False)

            Else
                Dim f2 As New Перевозчики
                Dim nom As String = НомерДог(i)
                With f2
                    .Названиеорганизации = Trim(RichTextBox1.Text)
                    .Форма_собственности = ComboBox1.Text
                    .Адрес_организации = Trim(RichTextBox2.Text)
                    .Почтовый_адрес = Trim(RichTextBox3.Text)
                    .РасчСчетРубли = Trim(RichTextBox6.Text)
                    .РасчСчетРоссРубли = Trim(RichTextBox9.Text)
                    .РасчСчетДоллар = Trim(RichTextBox5.Text)
                    .РасчСчетЕвро = Trim(RichTextBox4.Text)
                    .Адрес_банка = Trim(RichTextBox8.Text)
                    .Контактное_лицо = Trim(RichTextBox7.Text)
                    .Должность = Trim(RichTextBox13.Text)
                    .НаОснЧегоДейств = Trim(RichTextBox10.Text)
                    .ФИОРуководителя = Trim(RichTextBox12.Text)
                    .ФИОРодпадеж = Trim(RichTextBox11.Text)
                    .ДолжРодПадеж = Trim(RichTextBox14.Text)
                    .ПерЭкспедитор = f
                    .Договор = nom
                    .Дата = MaskedTextBox1.Text

                End With
                db.Перевозчики.InsertOnSubmit(f2)
                db.SubmitChanges()

                Dim f4 As New List(Of Перевозчики)
                f4.AddRange(listbx1)
                f4.Add(f2)
                If listbx1 IsNot Nothing Then
                    listbx1.Clear()
                End If
                If com4all IsNot Nothing Then
                    com4all.Clear()
                End If
                listbx1.AddRange(f4.OrderBy(Function(x) x.Названиеорганизации).Select(Function(x) x).ToList())
                com4all.AddRange(listbx1)
                bscom4all.ResetBindings(False)
                bslistbx1.ResetBindings(False)
            End If
        End Using

        mo.ПеревозчикиAllAsync()




        'ПеревозчикиRunMoving()

        If CheckBox4.Checked = True And CheckBox2.Checked = False Then
            ДокиПер()
            MessageBox.Show("Договор сформирован!", Рик)
        ElseIf CheckBox4.Checked = True And CheckBox2.Checked = True Then
            ДокиПерЭксп()
            MessageBox.Show("Договор сформирован!", Рик)
        Else
            MessageBox.Show("Данные внесены!", Рик)
        End If


        Me.Cursor = Cursors.Default

    End Sub
    Private Function НомерДог(ByVal i As Integer)
        If Not IO.Directory.Exists("Z:\RICKMANS\" & Now.Year & "\ДОГОВОР П\") Then
            IO.Directory.CreateDirectory("Z:\RICKMANS\" & Now.Year & "\ДОГОВОР П\")
        End If


        Dim Files() As String = IO.Directory.GetFiles("Z:\RICKMANS\" & Now.Year & "\ДОГОВОР П\", "*", IO.SearchOption.TopDirectoryOnly)
        кол = Nothing
        кол = Files.Length + 1
        If i = 0 Then
            Return кол & "/" & Now.Year
        Else
            Return кол & "-Э/" & Now.Year
        End If


    End Function
    Public Sub ДокиПер()


        Dim расч As New List(Of String)() From {RichTextBox6.Text, RichTextBox5.Text, RichTextBox4.Text, RichTextBox9.Text}
        Dim готрасч As String = String.Empty
        For i As Integer = 0 To расч.Count - 1
            If расч(i) <> "" Then
                готрасч = готрасч & ", IBAN - " & расч(i)
            End If
        Next




        Dim nam As String = RichTextBox1.Text
        Dim ds As Перевозчики = listbx1.Where(Function(x) x.Названиеорганизации = nam).Select(Function(x) x).FirstOrDefault()


        If ds Is Nothing Then Return

        Dim oWord As Microsoft.Office.Interop.Word.Application
        Dim oWordDoc As Microsoft.Office.Interop.Word.Document


        oWord = CreateObject("Word.Application")
        oWord.Visible = False


        Dim d As String = ""
        Select Case ComboBox2.Text
            Case "Евро"
                d = "евро"
            Case "Доллар"
                d = "доллар"
            Case "Российский рубль"
                d = "россрубл"
        End Select

        oWordDoc = oWord.Documents.Add("Z:\RICKMANS\ДОГОВОРА\ДОГОВОР П.doc")
        'Dim Files() As String = IO.Directory.GetFiles("Z:\RICKMANS\ДОГОВОРА", "*.txt", IO.SearchOption.TopDirectoryOnly)



        With oWordDoc.Bookmarks
            .Item("П1").Range.Text = ds.Договор
            .Item("П2").Range.Text = ds.Дата & "г."
            If ComboBox1.Text = "Индивидуальный предприниматель" Then
                .Item("П3").Range.Text = ComboBox1.Text & " " & RichTextBox1.Text
                .Item("П6").Range.Text = ComboBox1.Text
                .Item("П15").Range.Text = RichTextBox1.Text
            Else
                .Item("П3").Range.Text = ComboBox1.Text & " «" & RichTextBox1.Text & "»"
                .Item("П6").Range.Text = ComboBox1.Text
                .Item("П15").Range.Text = " «" & RichTextBox1.Text & "»"
            End If

            .Item("П4").Range.Text = Strings.LCase(RichTextBox14.Text)
            .Item("П5").Range.Text = RichTextBox11.Text
            .Item("П14").Range.Text = RichTextBox10.Text
            .Item("П7").Range.Text = RichTextBox2.Text & ", " & Trim(RichTextBox3.Text) & ", " & Trim(RichTextBox7.Text) & готрасч & ", " & Trim(RichTextBox8.Text)
            .Item("П8").Range.Text = RichTextBox13.Text
            .Item("П9").Range.Text = ФИОКорРук(RichTextBox12.Text, False)
            .Item("П11").Range.Text = РасчетнСчет(d)

            If Now > CDate("10.10.2021") Then
                .Item("П16").Range.Text = FaceNew
                .Item("П17").Range.Text = OsnNew
                .Item("П18").Range.Text = DoljNews
                .Item("П19").Range.Text = NameNews
                '.Item("П20").Range.Text = TypeNews

                'Else
                '    .Item("П16").Range.Text = FaceOld
                '    .Item("П17").Range.Text = OsnOld
                '    .Item("П18").Range.Text = DoljOld
                '    .Item("П19").Range.Text = NameOld
            End If

            'If CheckBox3.Checked = False And CheckBox2.Checked = False Then
            '    .Item("Кл12").Range.Text = ""
            '    .Item("Кл13").Range.Text = ""
            'ElseIf CheckBox3.Checked = True Then
            '    .Item("Кл12").Range.Text = "5.8.  Общая стоимость планируемых услуг по данному договору составит"
            '    .Item("Кл13").Range.Text = "500 000 евро."
            'ElseIf CheckBox2.Checked = True Then
            '    .Item("Кл12").Range.Text = "5.8.  Общая стоимость планируемых услуг по данному договору составит"
            '    .Item("Кл13").Range.Text = "500 000 российских рублей."
            'End If

        End With

        Dim NumdeReysa As String = ds.Договор

        If NumdeReysa IsNot Nothing Then
            NumdeReysa = Strings.Left(ds.Договор, CType(CType(NumdeReysa.Length, Integer) - 5, String))
        End If


        If Not IO.Directory.Exists("Z:\RICKMANS\" & Now.Year & "\ДОГОВОР П\") Then
            IO.Directory.CreateDirectory("Z:\RICKMANS\" & Now.Year & "\ДОГОВОР П\")
        End If

        Dim dog As String = "Z:\RICKMANS\" & Now.Year & "\ДОГОВОР П\ДОГОВОР П - " & кол & " " & RichTextBox1.Text & ".doc"
        Dim dog1 As String = "Z:\RICKMANS\" & Now.Year & "\ДОГОВОР П\ДОГОВОР П - " & NumdeReysa & " " & RichTextBox1.Text & ".doc"

        Try
            If кол = 0 Then
                oWordDoc.SaveAs2(dog1,,,,,, False)
            Else
                oWordDoc.SaveAs2(dog,,,,,, False)
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
        Dim СохрЗак As String
        If кол = 0 Then
            СохрЗак = dog1
        Else
            СохрЗак = dog
        End If

        'Try
        '    IO.File.Copy("C:\Users\Public\Documents\Рик\" & ДПодНом & " " & TextBox1.Text & " от " & MaskedTextBox6.Text & "(Договор подряда)" & ".doc", "U:\Офис\Финансовый\6. Бух.услуги\Кадры\" & Клиент & "\Договор подряда\" & Год & "\" & ДПодНом & " " & TextBox1.Text & " от " & MaskedTextBox6.Text & "(Договор подряда)" & ".doc")
        'Catch ex As Exception
        '    If MessageBox.Show("Договор подряда " & ДПодНом & " " & TextBox1.Text & " от " & MaskedTextBox6.Text & "(Договор подряда)" & " существует. Заменить старый документ новым?", Рик, MessageBoxButtons.OKCancel, MessageBoxIcon.Information) = DialogResult.OK Then
        '        IO.File.Delete("U:\Офис\Финансовый\6. Бух.услуги\Кадры\" & Клиент & "\Договор подряда\" & Год & "\" & ДПодНом & " " & TextBox1.Text & " от " & MaskedTextBox6.Text & "(Договор подряда)" & ".doc")
        '        IO.File.Copy("C:\Users\Public\Documents\Рик\" & ДПодНом & " " & TextBox1.Text & " от " & MaskedTextBox6.Text & "(Договор подряда)" & ".doc", "U:\Офис\Финансовый\6. Бух.услуги\Кадры\" & Клиент & "\Договор подряда\" & Год & "\" & ДПодНом & " " & TextBox1.Text & " от " & MaskedTextBox6.Text & "(Договор подряда)" & ".doc")
        '    End If
        'End Try

        oWordDoc.Close(True)
        oWord.Quit(True)
        If MessageBox.Show("Печатать договор?", Рик, MessageBoxButtons.OKCancel, MessageBoxIcon.Question) = DialogResult.OK Then
            ПечатьДоков(СохрЗак, CType(ComboBox3.Text, Integer))
        End If


    End Sub

    Private Sub ComboBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox1.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            RichTextBox1.Focus()
        End If
    End Sub

    Private Sub RichTextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles RichTextBox1.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            RichTextBox2.Focus()
        End If
    End Sub

    Private Sub RichTextBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles RichTextBox2.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            RichTextBox3.Focus()
        End If
    End Sub

    Private Sub RichTextBox3_KeyDown(sender As Object, e As KeyEventArgs) Handles RichTextBox3.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            RichTextBox6.Focus()
        End If
    End Sub

    Private Sub RichTextBox6_KeyDown(sender As Object, e As KeyEventArgs) Handles RichTextBox6.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            RichTextBox5.Focus()
        End If
    End Sub

    Private Sub RichTextBox5_KeyDown(sender As Object, e As KeyEventArgs) Handles RichTextBox5.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            RichTextBox4.Focus()
        End If
    End Sub

    Private Sub RichTextBox4_KeyDown(sender As Object, e As KeyEventArgs) Handles RichTextBox4.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            RichTextBox9.Focus()
        End If
    End Sub

    Private Sub RichTextBox9_KeyDown(sender As Object, e As KeyEventArgs) Handles RichTextBox9.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            RichTextBox8.Focus()
        End If
    End Sub

    Private Sub RichTextBox8_KeyDown(sender As Object, e As KeyEventArgs) Handles RichTextBox8.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            RichTextBox7.Focus()
        End If
    End Sub

    Private Sub RichTextBox7_KeyDown(sender As Object, e As KeyEventArgs) Handles RichTextBox7.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            RichTextBox13.Focus()
        End If
    End Sub

    Private Sub RichTextBox13_KeyDown(sender As Object, e As KeyEventArgs) Handles RichTextBox13.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            RichTextBox14.Focus()
        End If
    End Sub

    Private Sub RichTextBox14_KeyDown(sender As Object, e As KeyEventArgs) Handles RichTextBox14.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            RichTextBox12.Focus()
        End If
    End Sub

    Private Sub RichTextBox12_KeyDown(sender As Object, e As KeyEventArgs) Handles RichTextBox12.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            RichTextBox11.Focus()
        End If
    End Sub

    Private Sub RichTextBox11_KeyDown(sender As Object, e As KeyEventArgs) Handles RichTextBox11.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            RichTextBox10.Focus()
        End If
    End Sub

    Private Sub RichTextBox10_KeyDown(sender As Object, e As KeyEventArgs) Handles RichTextBox10.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Button1.Focus()
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If MessageBox.Show("Удалить перевозчика " & КодДляУдал.Названиеорганизации & " ?", Рик, MessageBoxButtons.OKCancel, MessageBoxIcon.Question) = DialogResult.OK Then
            Dim mo As New AllUpd
            Using db As New dbAllDataContext(_cn3)
                Dim f = db.Перевозчики.Where(Function(x) x.Названиеорганизации = КодДляУдал.Названиеорганизации).Select(Function(x) x).FirstOrDefault()
                If f IsNot Nothing Then
                    db.Перевозчики.DeleteOnSubmit(f)
                    db.SubmitChanges()

                End If
                com4all.Remove(КодДляУдал)
                listbx1.Remove(КодДляУдал)
                bscom4all.ResetBindings(False)
                bslistbx1.ResetBindings(False)
            End Using
            mo.ПеревозчикиAllAsync()
            очистка()
        End If

    End Sub

    Private Sub RichTextBox13_TextChanged(sender As Object, e As EventArgs) Handles RichTextBox13.TextChanged
        'sender.text = StrConv(sender.text, VbStrConv.ProperCase)
        'sender.SelectionStart = sender.text.Length
        Me.RichTextBox14.Text = Me.RichTextBox13.Text
    End Sub

    Private Sub RichTextBox12_TextChanged(sender As Object, e As EventArgs) Handles RichTextBox12.TextChanged
        sender.text = StrConv(sender.text, VbStrConv.ProperCase)
        sender.SelectionStart = sender.text.Length
        Me.RichTextBox11.Text = Me.RichTextBox12.Text
    End Sub



    Private Sub RichTextBox1_TextChanged(sender As Object, e As EventArgs) Handles RichTextBox1.TextChanged
        RichTextBox1.Font = New Font(RichTextBox1.Text, 10)
    End Sub

    Private Sub RichTextBox2_TextChanged(sender As Object, e As EventArgs) Handles RichTextBox2.TextChanged
        RichTextBox2.Font = New Font(RichTextBox2.Text, 10)
    End Sub

    Private Sub RichTextBox3_TextChanged(sender As Object, e As EventArgs) Handles RichTextBox3.TextChanged
        RichTextBox3.Font = New Font(RichTextBox3.Text, 10)
    End Sub

    Private Sub RichTextBox8_TextChanged(sender As Object, e As EventArgs) Handles RichTextBox8.TextChanged
        RichTextBox8.Font = New Font(RichTextBox8.Text, 10)
    End Sub

    Private Sub RichTextBox10_TextChanged(sender As Object, e As EventArgs) Handles RichTextBox10.TextChanged
        RichTextBox10.Font = New Font(RichTextBox10.Text, 10)
    End Sub

    Private Sub ComboBox4_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles ComboBox4.SelectionChangeCommitted
        'For I As Integer = 0 To ListBox1.Items.Count - 1
        '    If ListBox1.Items(I) = ComboBox4.Text Then
        '        ListBox1.SelectedIndex = I
        '        ListBox1R()
        '    End If
        'Next

        ListBox1.SelectedIndex = ComboBox4.SelectedIndex
        ListBox1R()
    End Sub
    Private Sub ComboBox4_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox4.KeyDown
        If e.KeyCode = Keys.Enter Then
            If ComboBox4.Text.Length = 0 Then Return
            Dim f = ComboBox4.SelectedIndex
            ListBox1.SelectedIndex = f
            ListBox1R()
        End If
    End Sub
End Class