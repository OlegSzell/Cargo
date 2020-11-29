Option Explicit On
Imports System.ComponentModel
Imports System.Data.OleDb
Imports System.Threading
Public Class НовыйКлиент
    Dim strsql As String
    Dim ds As DataTable
    Dim КодДляУдал As Клиент
    Dim кол As Integer


    Private com1all As List(Of IDNaz)
    Private bscom1 As BindingSource

    Private com4all As List(Of Клиент)
    Private bscom4 As BindingSource
    Private listbx1 As List(Of Клиент)
    Private bslistbx1 As BindingSource



    Private Sub НовыйКлиент_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        com1all = New List(Of IDNaz)
        bscom1 = New BindingSource
        bscom1.DataSource = com1all
        ComboBox1.DataSource = bscom1
        ComboBox1.DisplayMember = "Naz"
        ComboBox1.Text = String.Empty

        com4all = New List(Of Клиент)
        bscom4 = New BindingSource
        bscom4.DataSource = com4all
        ComboBox4.DataSource = bscom4
        ComboBox4.DisplayMember = "НазваниеОрганизации"
        ComboBox4.Text = String.Empty


        listbx1 = New List(Of Клиент)
        bslistbx1 = New BindingSource
        bslistbx1.DataSource = listbx1
        ListBox1.DataSource = bslistbx1
        ListBox1.DisplayMember = "НазваниеОрганизации"
        ListBox1.Text = String.Empty


        ComboBox3.Text = "1"

        'ComboBox1.AutoCompleteCustomSource.Clear()
        'ComboBox1.Items.Clear()

        'ComboBox4.Items.Clear()
        'ComboBox4.AutoCompleteCustomSource.Clear()

        Dim mo As New AllUpd
        Do While AllClass.Клиент Is Nothing
            mo.КлиентAll()
        Loop

        Do While AllClass.ФормаСобств Is Nothing
            mo.ФормаСобствAll()
        Loop


        Dim f = (From x In AllClass.Клиент
                 Order By x.НазваниеОрганизации
                 Select x).ToList()


        If f IsNot Nothing Then
            For Each b In f
                com4all.Add(b)
                listbx1.Add(b)
            Next
        End If
        bscom4.ResetBindings(False)
        bslistbx1.ResetBindings(False)



        Dim f1 = (From x In AllClass.ФормаСобств
                  Order By x.ПолноеНазвание
                  Select x).ToList()

        If f1 IsNot Nothing Then
            For Each b In f1
                Dim f2 As New IDNaz With {.Naz = b.ПолноеНазвание}
                com1all.Add(f2)
            Next
        End If
        bscom1.ResetBindings(False)

        'Using db As New dbAllDataContext
        '    Dim var = db.Клиент.OrderBy(Function(x) x.НазваниеОрганизации).Select(Function(x) x.НазваниеОрганизации).ToList()
        '    If var.Count > 0 Then
        '        For Each r In var
        '            ListBox1.Items.Add(r)
        '            ComboBox4.Items.Add(r)
        '            ComboBox4.AutoCompleteCustomSource.Add(r)
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







        MaskedTextBox1.Text = Now.ToShortDateString
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
    Private Sub ВставкаКлиент(ByVal ds As Клиент)


        ComboBox1.Text = ds.Форма_собственности
        RichTextBox1.Text = ds.НазваниеОрганизации
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
        RichTextBox11.Text = ds.ФИОРодПадеж
        RichTextBox13.Text = ds.Должность
        RichTextBox14.Text = ds.ДолжРодПадеж
    End Sub
    Private Sub ListBox1_Click(sender As Object, e As EventArgs) Handles ListBox1.Click
        ListBox1R()
    End Sub
    Private Sub ListBox1R()
        If Not ListBox1.SelectedIndex = -1 Then
            Try
                очистка()
                Dim mo As New AllUpd
                Do While AllClass.Клиент Is Nothing
                    mo.КлиентAll()
                Loop



                'Dim ds = AllClass.Клиент.Where(Function(x) x = ListBox1.SelectedItem).Select(Function(x) x).FirstOrDefault()

                ВставкаКлиент(ListBox1.SelectedItem)
                КодДляУдал = ListBox1.SelectedItem
                ComboBox4.Text = ""

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

        Return 0
    End Function
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Cursor = Cursors.WaitCursor
        If Проверка() = 1 Then
            Me.Cursor = Cursors.Default
            Exit Sub
        End If

        Dim mo As New AllUpd
        Do While AllClass.Клиент Is Nothing
            mo.КлиентAll()
        Loop

        If AllClass.Клиент Is Nothing Then Return
        Dim ds = AllClass.Клиент.Where(Function(x) x.НазваниеОрганизации = RichTextBox1.Text).Select(Function(x) x).FirstOrDefault()

        If ds IsNot Nothing Then
            Using db As New dbAllDataContext
                Dim f = db.Клиент.Where(Function(x) x.НазваниеОрганизации = RichTextBox1.Text).Select(Function(x) x).FirstOrDefault()
                With f
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
                    .ФИОРодПадеж = Trim(RichTextBox11.Text)
                    .ДолжРодПадеж = Trim(RichTextBox14.Text)
                    .Дата = MaskedTextBox1.Text
                End With
                db.SubmitChanges()
            End Using
            mo.КлиентAllAsync()
            'обновляем листбокс
            Dim f1 = listbx1.ElementAt(ListBox1.SelectedIndex)
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
                .ФИОРодПадеж = Trim(RichTextBox11.Text)
                .ДолжРодПадеж = Trim(RichTextBox14.Text)
                .Дата = MaskedTextBox1.Text
            End With
            bslistbx1.ResetBindings(False)

            Dim f2 = com4all.ElementAt(ListBox1.SelectedIndex)
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
                .ФИОРодПадеж = Trim(RichTextBox11.Text)
                .ДолжРодПадеж = Trim(RichTextBox14.Text)
                .Дата = MaskedTextBox1.Text
            End With
            bscom4.ResetBindings(False)


        Else
            Dim dog As String = НомерДог()
            Using db As New dbAllDataContext()
                Dim f As New Клиент

                With f
                    .НазваниеОрганизации = Trim(RichTextBox1.Text)
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
                    .ФИОРодПадеж = Trim(RichTextBox11.Text)
                    .ДолжРодПадеж = Trim(RichTextBox14.Text)
                    .Договор = dog
                    .Дата = MaskedTextBox1.Text
                End With
                db.Клиент.InsertOnSubmit(f)
                db.SubmitChanges()
            End Using


            mo.КлиентAllAsync()


            Dim gh As New List(Of Клиент)
            gh.AddRange(listbx1)
            Dim gh1 As New Клиент
            With gh1
                .НазваниеОрганизации = Trim(RichTextBox1.Text)
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
                .ФИОРодПадеж = Trim(RichTextBox11.Text)
                .ДолжРодПадеж = Trim(RichTextBox14.Text)
                .Договор = dog
                .Дата = MaskedTextBox1.Text
            End With
            gh.Add(gh1)

            If listbx1 IsNot Nothing Then
                listbx1.Clear()
            End If

            If com4all IsNot Nothing Then
                com4all.Clear()
            End If


            If gh IsNot Nothing Then
                Dim lop = gh.OrderBy(Function(x) x.НазваниеОрганизации).Select(Function(x) x).ToList()
                If lop IsNot Nothing Then
                    For Each b1 In lop
                        listbx1.Add(b1)
                        com4all.Add(b1)
                    Next
                    bscom4.ResetBindings(False)
                    bslistbx1.ResetBindings(False)
                End If

            End If

        End If


        If CheckBox4.Checked = True Then
            Доки()
        Else
            MessageBox.Show("Данные внесены!", Рик)
        End If




        Me.Cursor = Cursors.Default








    End Sub
    Private Function НомерДог()
        If Not IO.Directory.Exists("Z:\RICKMANS\" & Now.Year & "\ДОГОВОР З\") Then
            IO.Directory.CreateDirectory("Z:\RICKMANS\" & Now.Year & "\ДОГОВОР З\")
        End If

        Dim Files = IO.Directory.GetFiles("Z:\RICKMANS\" & Now.Year & "\ДОГОВОР З\", "*", IO.SearchOption.TopDirectoryOnly).ToList()
        кол = Nothing
        кол = Files.Count + 1
        Return кол & "/" & Now.Year

    End Function
    Public Sub Доки()


        Dim расч As New List(Of String)() From {RichTextBox6.Text, RichTextBox5.Text, RichTextBox4.Text, RichTextBox9.Text}
        Dim готрасч As String
        For i As Integer = 0 To расч.Count - 1
            If расч(i) <> "" Then
                готрасч = готрасч & ", IBAN - " & расч(i)
            End If
        Next



        Dim ds = listbx1.Where(Function(x) x.НазваниеОрганизации = RichTextBox1.Text).Select(Function(x) x).FirstOrDefault()

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

        oWordDoc = oWord.Documents.Add("Z:\RICKMANS\ДОГОВОРА\ДОГОВОР З.doc")
        'Dim Files() As String = IO.Directory.GetFiles("Z:\RICKMANS\ДОГОВОРА", "*.txt", IO.SearchOption.TopDirectoryOnly)



        With oWordDoc.Bookmarks
            .Item("Кл1").Range.Text = ds.Договор
            .Item("Кл2").Range.Text = ds.Дата & "г."
            If ComboBox1.Text = "Индивидуальный предприниматель" Then
                .Item("Кл3").Range.Text = ComboBox1.Text & " " & RichTextBox1.Text
                .Item("Кл6").Range.Text = ComboBox1.Text
                .Item("Кл15").Range.Text = RichTextBox1.Text
            Else
                .Item("Кл3").Range.Text = ComboBox1.Text & " «" & RichTextBox1.Text & "»"
                .Item("Кл6").Range.Text = ComboBox1.Text
                .Item("Кл15").Range.Text = " «" & RichTextBox1.Text & "»"
            End If

            .Item("Кл4").Range.Text = Strings.LCase(RichTextBox14.Text)
            .Item("Кл5").Range.Text = RichTextBox11.Text
            .Item("Кл14").Range.Text = RichTextBox10.Text
            .Item("Кл7").Range.Text = Trim(RichTextBox2.Text) & ", " & Trim(RichTextBox3.Text) & ", " & Trim(RichTextBox7.Text)
            .Item("Кл8").Range.Text = готрасч & ", " & Trim(RichTextBox8.Text)
            .Item("Кл9").Range.Text = RichTextBox13.Text
            .Item("Кл10").Range.Text = ФИОКорРук(RichTextBox12.Text, False)
            .Item("Кл11").Range.Text = РасчетнСчет(d)

            If CheckBox3.Checked = False And CheckBox2.Checked = False Then
                .Item("Кл12").Range.Text = ""
                .Item("Кл13").Range.Text = ""
            ElseIf CheckBox3.Checked = True Then
                .Item("Кл12").Range.Text = "5.8.  Общая стоимость планируемых услуг по данному договору составит"
                .Item("Кл13").Range.Text = "500 000 евро."
            ElseIf CheckBox2.Checked = True Then
                .Item("Кл12").Range.Text = "5.8.  Общая стоимость планируемых услуг по данному договору составит"
                .Item("Кл13").Range.Text = "500 000 российских рублей."
            End If

        End With

        Dim NumdeReysa As String = ds.Договор

        NumdeReysa = Strings.Left(ds.Договор, CType(CType(NumdeReysa.Length, Integer) - 5, String))

        If Not IO.Directory.Exists("Z:\RICKMANS\" & Now.Year & "\ДОГОВОР З\") Then
            IO.Directory.CreateDirectory("Z:\RICKMANS\" & Now.Year & "\ДОГОВОР З\")
        End If

        Dim dog As String = "Z:\RICKMANS\" & Now.Year & "\ДОГОВОР З\ДОГОВОР З - " & кол & " " & RichTextBox1.Text & ".doc"
        Dim dog1 As String = "Z:\RICKMANS\" & Now.Year & "\ДОГОВОР З\ДОГОВОР З - " & NumdeReysa & " " & RichTextBox1.Text & ".doc"
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




        oWordDoc.Close(True)
        oWord.Quit(True)
        If MessageBox.Show("Договор оформлен.Распечатать?", Рик, MessageBoxButtons.OKCancel, MessageBoxIcon.Question) = DialogResult.OK Then
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
        If MessageBox.Show("Удалить клиента " & КодДляУдал.НазваниеОрганизации & " ?", Рик, MessageBoxButtons.OKCancel, MessageBoxIcon.Question) = DialogResult.OK Then
            If КодДляУдал Is Nothing Then
                MessageBox.Show("Выберите клиента для удаления!", Рик)
                Return
            End If

            Using db As New dbAllDataContext()
                'Dim f = КодДляУдал
                Dim f = db.Клиент.Where(Function(x) x.НазваниеОрганизации = КодДляУдал.НазваниеОрганизации).Select(Function(x) x).FirstOrDefault()
                If f IsNot Nothing Then
                    db.Клиент.DeleteOnSubmit(f)
                    db.SubmitChanges()
                End If
            End Using
            Dim mo As New AllUpd
            mo.КлиентAllAsync()


            очистка()


            listbx1.Remove(КодДляУдал)
            com4all.Remove(КодДляУдал)
            bscom4.ResetBindings(False)
            bslistbx1.ResetBindings(False)
        End If

    End Sub

    Private Sub CheckBox3_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox3.CheckedChanged
        If CheckBox3.Checked = True Then
            CheckBox2.Checked = False

        End If
    End Sub

    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox2.CheckedChanged
        If CheckBox2.Checked = True Then
            CheckBox3.Checked = False

        End If
    End Sub

    Private Sub RichTextBox13_TextChanged(sender As Object, e As EventArgs) Handles RichTextBox13.TextChanged
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

        Dim f = ComboBox4.SelectedIndex
        ListBox1.SelectedIndex = f
        ListBox1R()
        'For I As Integer = 0 To ListBox1.Items.Count - 1
        '    If ListBox1.Items(I) = ComboBox4.Text Then
        '        ListBox1.SelectedIndex = I
        '        ListBox1R()
        '    End If
        'Next
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