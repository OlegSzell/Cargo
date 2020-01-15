Option Explicit On
Imports System.Data.OleDb
Public Class НовыйПеревоз
    Dim strsql As String
    Dim ds As DataTable
    Dim КодДляУдал As String
    Dim кол As Integer
    Private Sub НовыйПеревоз_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        conn = New OleDbConnection
        conn.ConnectionString = ConString
        Try
            conn.Open()
        Catch ex As Exception
            MessageBox.Show("Не подключен диск U")
        End Try

        strsql = "SELECT ПолноеНазвание FROM ФормаСобств ORDER BY ПолноеНазвание"
        ds = Selects(strsql)

        Me.ComboBox1.AutoCompleteCustomSource.Clear()
        Me.ComboBox1.Items.Clear()
        For Each r As DataRow In ds.Rows
            Me.ComboBox1.AutoCompleteCustomSource.Add(r.Item(0).ToString())
            Me.ComboBox1.Items.Add(r(0).ToString)
        Next

        'ComboBox2.Items.Clear()
        'Dim d() As String = {Now.Year, Now.Year - 1, Now.Year - 2, Now.Year - 3, Now.Year - 4}
        'ComboBox2.Items.AddRange(d)


        Dim strsql1 As String = "SELECT Названиеорганизации FROM Перевозчики ORDER BY НазваниеОрганизации"
        Dim ds1 As DataTable = Selects(strsql1)
        ComboBox3.Text = "1"

        Me.ListBox1.Items.Clear()
        For Each r As DataRow In ds1.Rows
            Me.ListBox1.Items.Add(r(0).ToString)
        Next
        MaskedTextBox1.Text = Now.ToShortDateString
    End Sub
    Private Sub очистка()
        ComboBox1.Text = ""
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
    Private Sub ВставкаКлиент(ByVal ds As DataTable)


        ComboBox1.Text = ds.Rows(0).Item(1).ToString
        RichTextBox1.Text = ds.Rows(0).Item(0).ToString
        RichTextBox2.Text = ds.Rows(0).Item(2).ToString
        RichTextBox3.Text = ds.Rows(0).Item(3).ToString
        RichTextBox4.Text = ds.Rows(0).Item(7).ToString
        RichTextBox5.Text = ds.Rows(0).Item(6).ToString
        RichTextBox6.Text = ds.Rows(0).Item(4).ToString
        RichTextBox7.Text = ds.Rows(0).Item(9).ToString & " " & ds.Rows(0).Item(10).ToString
        RichTextBox8.Text = ds.Rows(0).Item(8).ToString
        RichTextBox9.Text = ds.Rows(0).Item(5).ToString
        RichTextBox10.Text = ds.Rows(0).Item(14).ToString
        RichTextBox12.Text = ds.Rows(0).Item(15).ToString
        RichTextBox11.Text = ds.Rows(0).Item(19).ToString
        RichTextBox13.Text = ds.Rows(0).Item(13).ToString
        RichTextBox14.Text = ds.Rows(0).Item(21).ToString
        If ds.Rows(0).Item(20).ToString = "Да" Then
            CheckBox2.Checked = True
        Else
            CheckBox2.Checked = False
        End If
    End Sub
    Private Sub ListBox1_Click(sender As Object, e As EventArgs) Handles ListBox1.Click
        If Not ListBox1.SelectedIndex = -1 Then
            Try
                очистка()
                strsql = ""
                strsql = "SELECT * FROM Перевозчики WHERE Названиеорганизации='" & ListBox1.SelectedItem & "'"
                Dim ds2 As DataTable = Selects(strsql)
                ВставкаКлиент(ds2)
                КодДляУдал = ds2.Rows(0).Item(0).ToString
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
    Private Sub ДокиПерЭксп()
        Dim расч As New List(Of String)() From {RichTextBox6.Text, RichTextBox5.Text, RichTextBox4.Text, RichTextBox9.Text}
        Dim готрасч As String
        For i As Integer = 0 To расч.Count - 1
            If расч(i) <> "" Then
                готрасч = готрасч & ", IBAN - " & расч(i)
            End If
        Next

        'Dim Files() As String = IO.Directory.GetFiles("Z:\RICKMANS\" & Now.Year & "\ДОГОВОР П\", "*", IO.SearchOption.TopDirectoryOnly)
        'Dim кол As Integer = Files.Length + 1


        Dim strsql2 As String = "SELECT Договор, Дата FROM Перевозчики WHERE  Названиеорганизации='" & RichTextBox1.Text & "'"
        Dim ds As DataTable = Selects(strsql2)


        Dim oWord As Microsoft.Office.Interop.Word.Application
        Dim oWordDoc As Microsoft.Office.Interop.Word.Document
        KillProc()
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
            .Item("ПЭ1").Range.Text = ds.Rows(0).Item(0).ToString
            .Item("ПЭ2").Range.Text = ds.Rows(0).Item(1).ToString & "г."
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


        End With
        Try
            oWordDoc.SaveAs2("Z:\RICKMANS\" & Now.Year & "\ДОГОВОР П\ДОГОВОР ПЭ - " & кол & " " & RichTextBox1.Text & ".doc",,,,,, False)
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

        Dim СохрЗак As String = "Z:\RICKMANS\" & Now.Year & "\ДОГОВОР П\ДОГОВОР ПЭ - " & кол & " " & RichTextBox1.Text & ".doc"

        oWordDoc.Close(True)
        oWord.Quit(True)
        ПечатьДоков(СохрЗак, CType(ComboBox3.Text, Integer))

    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Cursor = Cursors.WaitCursor
        If Проверка() = 1 Then Exit Sub

        Dim strsql1 As String = "SELECT РасчСчетРубли FROM Перевозчики WHERE Названиеорганизации='" & RichTextBox1.Text & "'"
        Dim ds As DataTable = Selects(strsql1)

        Dim f As String
        Dim i As Integer
        If CheckBox2.Checked = True Then
            f = "Да"
            i = 1
        Else
            f = ""
            i = 0
        End If


        If errds = 1 Then
            Dim strsql As String = "INSERT INTO Перевозчики(Названиеорганизации,[Форма собственности],[Адрес организации],[Почтовый адрес],РасчСчетРубли,РасчСчетРоссРубли,
        РасчСчетДоллар,РасчСчетЕвро,[Адрес банка],[Контактное лицо],Должность,НаОснЧегоДейств,ФИОРуководителя,ФИОРодПадеж,ДолжРодПадеж,ПерЭкспедитор,Договор,Дата)  VALUES('" & Trim(RichTextBox1.Text) & "','" & ComboBox1.Text & "','" & Trim(RichTextBox2.Text) & "',
        '" & Trim(RichTextBox3.Text) & "','" & Trim(RichTextBox6.Text) & "','" & Trim(RichTextBox9.Text) & "','" & Trim(RichTextBox5.Text) & "','" & Trim(RichTextBox4.Text) & "','" & Trim(RichTextBox8.Text) & "',
        '" & Trim(RichTextBox7.Text) & "','" & Trim(RichTextBox13.Text) & "','" & Trim(RichTextBox10.Text) & "','" & Trim(RichTextBox12.Text) & "','" & Trim(RichTextBox11.Text) & "','" & Trim(RichTextBox14.Text) & "', '" & f & "', '" & НомерДог(i) & "','" & MaskedTextBox1.Text & "')"
            Updates(strsql)
        Else
            Dim strsql As String = "UPDATE Перевозчики SET [Форма собственности]='" & ComboBox1.Text & "',[Адрес организации]='" & Trim(RichTextBox2.Text) & "',
        [Почтовый адрес]='" & Trim(RichTextBox3.Text) & "',[РасчСчетРубли]='" & Trim(RichTextBox6.Text) & "',[РасчСчетРоссРубли]='" & Trim(RichTextBox9.Text) & "',[РасчСчетДоллар]='" & Trim(RichTextBox5.Text) & "',
        РасчСчетЕвро='" & Trim(RichTextBox4.Text) & "',[Адрес банка]='" & Trim(RichTextBox8.Text) & "',[Контактное лицо]='" & Trim(RichTextBox7.Text) & "',Должность ='" & Trim(RichTextBox13.Text) & "', 
        НаОснЧегоДейств='" & Trim(RichTextBox10.Text) & "',ФИОРуководителя ='" & Trim(RichTextBox12.Text) & "',ФИОРодПадеж ='" & Trim(RichTextBox11.Text) & "',ДолжРодПадеж='" & Trim(RichTextBox14.Text) & "', ПерЭкспедитор='" & f & "', Дата='" & MaskedTextBox1.Text & "'
        WHERE Названиеорганизации='" & RichTextBox1.Text & "'"
            Updates(strsql)
        End If
        If CheckBox4.Checked = True And CheckBox2.Checked = False Then
            ДокиПер()
            MessageBox.Show("Договор сформирован. Идет печать!", Рик)
        ElseIf CheckBox4.Checked = True And CheckBox2.Checked = True Then
            ДокиПерЭксп()
            MessageBox.Show("Договор сформирован.Идет печать!", Рик)
        Else

            MessageBox.Show("Данные изменены!", Рик)
        End If
        Me.Cursor = Cursors.Default
    End Sub
    Private Function НомерДог(ByVal i As Integer)
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
        Dim готрасч As String
        For i As Integer = 0 To расч.Count - 1
            If расч(i) <> "" Then
                готрасч = готрасч & ", IBAN - " & расч(i)
            End If
        Next




        Dim strsql2 As String = "SELECT Договор,Дата FROM Перевозчики WHERE  Названиеорганизации='" & RichTextBox1.Text & "'"
        Dim ds As DataTable = Selects(strsql2)

        Dim oWord As Microsoft.Office.Interop.Word.Application
        Dim oWordDoc As Microsoft.Office.Interop.Word.Document
        'Dim oWordPara As Microsoft.Office.Interop.Word.Paragraph

        KillProc()

        oWord = CreateObject("Word.Application")
        oWord.Visible = False

        'Try
        '    IO.File.Copy("U:\Офис\Финансовый\6. Бух.услуги\ОБЩДОКИ\General\ДПодряда.doc", "C:\Users\Public\Documents\Рик\ДПодряда.doc")
        'Catch ex As Exception
        '    'If "Заявление.doc" <> "" Then IO.File.Delete("C:\Users\Public\Documents\Рик\Заявление.doc")
        '    If Not IO.Directory.Exists("c:\Users\Public\Documents\Рик") Then
        '        IO.Directory.CreateDirectory("c:\Users\Public\Documents\Рик")
        '    End If
        '    IO.File.Copy("U:\Офис\Финансовый\6. Бух.услуги\ОБЩДОКИ\General\ДПодряда.doc", "C:\Users\Public\Documents\Рик\ДПодряда.doc")
        'End Try
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
            .Item("П1").Range.Text = ds.Rows(0).Item(0).ToString
            .Item("П2").Range.Text = ds.Rows(0).Item(1).ToString & "г."
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
        Try
            oWordDoc.SaveAs2("Z:\RICKMANS\" & Now.Year & "\ДОГОВОР П\ДОГОВОР П - " & кол & " " & RichTextBox1.Text & ".doc",,,,,, False)
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

        Dim СохрЗак As String = "Z:\RICKMANS\" & Now.Year & "\ДОГОВОР П\ДОГОВОР П - " & кол & " " & RichTextBox1.Text & ".doc"

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
        ПечатьДоков(СохрЗак, CType(ComboBox3.Text, Integer))

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
        If MessageBox.Show("Удалить перевозчика " & КодДляУдал & " ?", Рик, MessageBoxButtons.OKCancel, MessageBoxIcon.Question) = DialogResult.OK Then
            Dim strsql As String = "DELETE * FROM Перевозчики WHERE Названиеорганизации='" & КодДляУдал & "'"
            Updates(strsql)
            очистка()

            Dim strsql1 As String = "SELECT Названиеорганизации FROM Перевозчики ORDER BY Названиеорганизации"
            Dim ds1 As DataTable = Selects(strsql1)
            Me.ListBox1.Items.Clear()
            For Each r As DataRow In ds1.Rows
                Me.ListBox1.Items.Add(r(0).ToString)
            Next
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
End Class