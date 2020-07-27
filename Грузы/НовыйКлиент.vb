Option Explicit On
Imports System.Data.OleDb
Imports System.Threading
Public Class НовыйКлиент
    Dim strsql As String
    Dim ds As DataTable
    Dim КодДляУдал As String
    Dim кол As Integer



    Private Sub НовыйКлиент_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ''conn = New OleDbConnection
        ''conn.ConnectionString = ConString
        ''Try
        ''    conn.Open()
        ''Catch ex As Exception
        ''    MessageBox.Show("Не подключен диск U")
        ''End Try

        'strsql = "SELECT ПолноеНазвание FROM ФормаСобств ORDER BY ПолноеНазвание"
        'ds = Selects3(strsql)

        'Me.ComboBox1.AutoCompleteCustomSource.Clear()
        'Me.ComboBox1.Items.Clear()
        'For Each r As DataRow In ds.Rows
        '    Me.ComboBox1.AutoCompleteCustomSource.Add(r.Item(0).ToString())
        '    Me.ComboBox1.Items.Add(r(0).ToString)
        'Next


        ComboBox3.Text = "1"

        'Dim strsql1 As String = "SELECT НазваниеОрганизации FROM Клиент ORDER BY НазваниеОрганизации"
        'Dim ds1 As DataTable = Selects3(strsql1)


        'Me.ListBox1.Items.Clear()
        'For Each r As DataRow In ds1.Rows
        '    Me.ListBox1.Items.Add(r(0).ToString)
        'Next

        ComboBox1.AutoCompleteCustomSource.Clear()
        ComboBox1.Items.Clear()

        ComboBox4.Items.Clear()
        ComboBox4.AutoCompleteCustomSource.Clear()

        Using db As New dbAllDataContext
            Dim var = db.Клиент.OrderBy(Function(x) x.НазваниеОрганизации).Select(Function(x) x.НазваниеОрганизации).ToList()
            If var.Count > 0 Then
                For Each r In var
                    ListBox1.Items.Add(r)
                    ComboBox4.Items.Add(r)
                    ComboBox4.AutoCompleteCustomSource.Add(r)
                Next

            End If
            Dim var1 = db.ФормаСобств.OrderBy(Function(x) x.ПолноеНазвание).Select(Function(x) x.ПолноеНазвание).ToList()
            If var1.Count > 0 Then
                For Each f In var1
                    ComboBox1.AutoCompleteCustomSource.Add(f)
                    ComboBox1.Items.Add(f)
                Next
            End If
        End Using





        'ComboBox4.Items.Clear()
        'ComboBox4.AutoCompleteCustomSource.Clear()
        'ListBox1.Items.Clear()
        'For Each r As DataRow In dtКлиенты.Rows
        '    ListBox1.Items.Add(r(0).ToString)
        '    ComboBox4.Items.Add(r(0).ToString)
        '    ComboBox4.AutoCompleteCustomSource.Add(r(0).ToString)
        'Next





        'Dim lv() As ComboBox = {ComboBox4}

        'Dim f As New Thread(Sub() Listxt(Me, "SELECT НазваниеОрганизации FROM Клиент ORDER BY НазваниеОрганизации", ListBox1, ComboBox4))
        'f.IsBackground = True
        'f.Start()
        'Dim f1 As New Thread(Sub() COMxt(Me, "SELECT ПолноеНазвание FROM ФормаСобств ORDER BY ПолноеНазвание", ComboBox1))
        'f1.IsBackground = True
        'f1.Start()


        'Dim c = From x In dtФормаСобствAll Order By x.Item("ПолноеНазвание").ToString Select x.Item("ПолноеНазвание")

        'ComboBox1.AutoCompleteCustomSource.Clear()
        'ComboBox1.Items.Clear()
        'For Each r In c
        '    ComboBox1.AutoCompleteCustomSource.Add(r.ToString)
        '    ComboBox1.Items.Add(r.ToString)
        'Next


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
    Private Sub ВставкаКлиент(ByVal ds() As DataRow)


        ComboBox1.Text = ds(0).Item(1).ToString
        RichTextBox1.Text = ds(0).Item(0).ToString
        RichTextBox2.Text = ds(0).Item(2).ToString
        RichTextBox3.Text = ds(0).Item(3).ToString
        RichTextBox4.Text = ds(0).Item(7).ToString
        RichTextBox5.Text = ds(0).Item(6).ToString
        RichTextBox6.Text = ds(0).Item(4).ToString
        RichTextBox7.Text = ds(0).Item(9).ToString & " " & ds(0).Item(10).ToString
        RichTextBox8.Text = ds(0).Item(8).ToString
        RichTextBox9.Text = ds(0).Item(5).ToString
        RichTextBox10.Text = ds(0).Item(14).ToString
        RichTextBox12.Text = ds(0).Item(15).ToString
        RichTextBox11.Text = ds(0).Item(19).ToString
        RichTextBox13.Text = ds(0).Item(13).ToString
        RichTextBox14.Text = ds(0).Item(20).ToString
    End Sub
    Private Sub ListBox1_Click(sender As Object, e As EventArgs) Handles ListBox1.Click
        ListBox1R()
    End Sub
    Private Sub ListBox1R()
        If Not ListBox1.SelectedIndex = -1 Then
            Try
                очистка()
                'strsql = ""
                'strsql = "SELECT * FROM Клиент WHERE НазваниеОрганизации='" & ListBox1.SelectedItem & "'"
                'Dim ds As DataTable = Selects3(StrSql:="SELECT * FROM Клиент WHERE НазваниеОрганизации='" & ListBox1.SelectedItem & "'")
                Dim ds = dtКлиенты.Select("НазваниеОрганизации='" & ListBox1.SelectedItem & "'")


                ВставкаКлиент(ds)
                КодДляУдал = ds(0).Item(0).ToString
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

        'Dim strsql1 As String = "SELECT [РасчСчетРоссРубли] FROM Клиент WHERE НазваниеОрганизации='" & RichTextBox1.Text & "'"
        'Dim ds As DataTable = Selects3(strsql1)

        Dim ds = dtКлиенты.Select("НазваниеОрганизации='" & RichTextBox1.Text & "'")

        If ds.Length = 0 Then
            Dim strsql As String = "INSERT INTO Клиент(НазваниеОрганизации,[Форма собственности],[Адрес организации],[Почтовый адрес],РасчСчетРубли,РасчСчетРоссРубли,
            РасчСчетДоллар,РасчСчетЕвро,[Адрес банка],[Контактное лицо],Должность,НаОснЧегоДейств,ФИОРуководителя,ФИОРодПадеж,ДолжРодПадеж,Договор,Дата) VALUES('" & Trim(RichTextBox1.Text) & "','" & ComboBox1.Text & "','" & Trim(RichTextBox2.Text) & "',
            '" & Trim(RichTextBox3.Text) & "','" & Trim(RichTextBox6.Text) & "','" & Trim(RichTextBox9.Text) & "','" & Trim(RichTextBox5.Text) & "','" & Trim(RichTextBox4.Text) & "','" & Trim(RichTextBox8.Text) & "',
            '" & Trim(RichTextBox7.Text) & "','" & Trim(RichTextBox13.Text) & "','" & Trim(RichTextBox10.Text) & "','" & Trim(RichTextBox12.Text) & "','" & Trim(RichTextBox11.Text) & "','" & Trim(RichTextBox14.Text) & "','" & НомерДог() & "', '" & MaskedTextBox1.Text & "')"
            Updates3(strsql)

        Else
            Dim strsql As String = "UPDATE Клиент SET [Форма собственности]='" & ComboBox1.Text & "',[Адрес организации]='" & Trim(RichTextBox2.Text) & "',
        [Почтовый адрес]='" & Trim(RichTextBox3.Text) & "',[РасчСчетРубли]='" & Trim(RichTextBox6.Text) & "',[РасчСчетРоссРубли]='" & Trim(RichTextBox9.Text) & "',[РасчСчетДоллар]='" & Trim(RichTextBox5.Text) & "',
        РасчСчетЕвро='" & Trim(RichTextBox4.Text) & "',[Адрес банка]='" & Trim(RichTextBox8.Text) & "',[Контактное лицо]='" & Trim(RichTextBox7.Text) & "',Должность ='" & Trim(RichTextBox13.Text) & "', 
        НаОснЧегоДейств='" & Trim(RichTextBox10.Text) & "',ФИОРуководителя ='" & Trim(RichTextBox12.Text) & "',ФИОРодПадеж ='" & Trim(RichTextBox11.Text) & "',ДолжРодПадеж='" & Trim(RichTextBox14.Text) & "',Дата='" & MaskedTextBox1.Text & "'
        WHERE НазваниеОрганизации='" & RichTextBox1.Text & "'"
            Updates3(strsql)
        End If
        КлиентыRunMoving()
        'Dim d2 As New Thread(AddressOf Клиенты)
        'd2.IsBackground = True
        'd2.Start()


        If CheckBox4.Checked = True Then
            Доки()
        Else
            MessageBox.Show("Данные внесены!", Рик)
        End If

        ComboBox4.Items.Clear()
        ComboBox4.AutoCompleteCustomSource.Clear()
        ListBox1.Items.Clear()
        For Each r As DataRow In dtКлиенты.Rows
            ListBox1.Items.Add(r(0).ToString)
            ComboBox4.Items.Add(r(0).ToString)
            ComboBox4.AutoCompleteCustomSource.Add(r(0).ToString)
        Next



        'Dim lv() As ComboBox = {ComboBox4} 'обновляем данные

        'Dim f5 As New Thread(Sub() Listxt(Me, "SELECT Названиеорганизации FROM Перевозчики ORDER BY НазваниеОрганизации", ListBox1, ComboBox4))
        'f5.IsBackground = True
        'f5.Start()
        Me.Cursor = Cursors.Default

        'Dim f1 As New Thread(Sub() COMxt(Me, "SELECT ПолноеНазвание FROM ФормаСобств ORDER BY ПолноеНазвание", ComboBox1))
        'f1.IsBackground = True
        'f1.Start()



        Dim c = From x In dtФормаСобствAll Order By x.Item("ПолноеНазвание").ToString Select x.Item("ПолноеНазвание")

        ComboBox1.AutoCompleteCustomSource.Clear()
        ComboBox1.Items.Clear()
        For Each r In c
            ComboBox1.AutoCompleteCustomSource.Add(r.ToString)
            ComboBox1.Items.Add(r.ToString)
        Next





    End Sub
    Private Function НомерДог()
        If Not IO.Directory.Exists("Z:\RICKMANS\" & Now.Year & "\ДОГОВОР З\") Then
            IO.Directory.CreateDirectory("Z:\RICKMANS\" & Now.Year & "\ДОГОВОР З\")
        End If

        Dim Files() As String = IO.Directory.GetFiles("Z:\RICKMANS\" & Now.Year & "\ДОГОВОР З\", "*", IO.SearchOption.TopDirectoryOnly)
        кол = Nothing
        кол = Files.Length + 1
        Return кол & "/" & Now.Year
        'Dim strsql2 As String = "UPDATE Клиент SET Договор='" & кол & "/" & Now.Year & "', Дата='" & MaskedTextBox1.Text & "' WHERE  НазваниеОрганизации ='" & RichTextBox1.Text & "'"
        'Updates(strsql2)
    End Function
    Public Sub Доки()


        Dim расч As New List(Of String)() From {RichTextBox6.Text, RichTextBox5.Text, RichTextBox4.Text, RichTextBox9.Text}
        Dim готрасч As String
        For i As Integer = 0 To расч.Count - 1
            If расч(i) <> "" Then
                готрасч = готрасч & ", IBAN - " & расч(i)
            End If
        Next

        'Dim strsql2 As String = "SELECT Договор, Дата FROM Клиент WHERE НазваниеОрганизации='" & RichTextBox1.Text & "'"
        'Dim ds As DataTable = Selects3(strsql2)

        Dim ds = dtКлиенты.Select("НазваниеОрганизации='" & RichTextBox1.Text & "'")

        Dim oWord As Microsoft.Office.Interop.Word.Application
        Dim oWordDoc As Microsoft.Office.Interop.Word.Document
        'Dim oWordPara As Microsoft.Office.Interop.Word.Paragraph

        'KillProc()

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

        oWordDoc = oWord.Documents.Add("Z:\RICKMANS\ДОГОВОРА\ДОГОВОР З.doc")
        'Dim Files() As String = IO.Directory.GetFiles("Z:\RICKMANS\ДОГОВОРА", "*.txt", IO.SearchOption.TopDirectoryOnly)



        With oWordDoc.Bookmarks
            .Item("Кл1").Range.Text = ds(0).Item("Договор").ToString
            .Item("Кл2").Range.Text = ds(0).Item("Дата").ToString & "г."
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

        Dim NumdeReysa As String = ds(0).Item("Договор").ToString

        NumdeReysa = Strings.Left(ds(0).Item("Договор").ToString, CType(CType(NumdeReysa.Length, Integer) - 5, String))

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
        If MessageBox.Show("Удалить клиента " & КодДляУдал & " ?", Рик, MessageBoxButtons.OKCancel, MessageBoxIcon.Question) = DialogResult.OK Then
            Dim strsql As String = "DELETE FROM Клиент WHERE НазваниеОрганизации='" & КодДляУдал & "'"
            Updates3(strsql)
            очистка()
            КлиентыRunMoving()


            'Dim strsql1 As String = "SELECT НазваниеОрганизации FROM Клиент ORDER BY НазваниеОрганизации"
            'Dim ds1 As DataTable = Selects3(strsql1)

            Me.ListBox1.Items.Clear()
            For Each r As DataRow In dtКлиенты.Rows
                Me.ListBox1.Items.Add(r(0).ToString)

            Next
            Me.ListBox1.SelectedIndex = 1
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

    Private Sub ComboBox4_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox4.SelectedIndexChanged
        For I As Integer = 0 To ListBox1.Items.Count - 1
            If ListBox1.Items(I) = ComboBox4.Text Then
                ListBox1.SelectedIndex = I
                ListBox1R()
            End If
        Next
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
End Class