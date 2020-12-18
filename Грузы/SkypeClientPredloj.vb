Public Class SkypeClientPredloj
    Private Com1all As List(Of IDNaz)
    Private bscom1 As BindingSource
    Private Sub SkypeClientPredloj_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Com1all = New List(Of IDNaz)
        bscom1 = New BindingSource
        bscom1.DataSource = Com1all
        ComboBox1.DataSource = bscom1
        ComboBox1.DisplayMember = "Naz"

        Dim mo As New AllUpd
        Do While AllClass.SkypeКлиентПредложение Is Nothing
            mo.SkypeКлиентПредложениеAll()
        Loop
        Dim f = AllClass.SkypeКлиентПредложение.OrderBy(Function(x) x.Клиент).Where(Function(x) x.Экспедитор = Экспедитор).Select(Function(x) x.Клиент).Distinct().ToList()
        If f IsNot Nothing Then
            For Each b In f
                Dim f1 As New IDNaz With {.Naz = b}
                Com1all.Add(f1)
            Next
            bscom1.ResetBindings(False)
        End If
        ComboBox1.Text = String.Empty

    End Sub

    Private Sub ComboBox1_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles ComboBox1.SelectionChangeCommitted
        Dim f2 As IDNaz = ComboBox1.SelectedItem
        RichTextBox1.Text = String.Empty
        RichTextBox2.Text = String.Empty

        Dim mo As New AllUpd
        Do While AllClass.SkypeКлиентПредложение Is Nothing
            mo.SkypeКлиентПредложениеAll()
        Loop
        Dim f = AllClass.SkypeКлиентПредложение.OrderBy(Function(x) x.Дата).Where(Function(x) x.Клиент = f2.Naz And x.Экспедитор = Экспедитор).ToList()
        If f IsNot Nothing Then
            For Each b In f
                If b.Дата Is Nothing Then Continue For
                If b.Время IsNot Nothing Then
                    If RichTextBox1.Text.Length = 0 Then
                        RichTextBox1.Text = " ____________________________________ " & vbCrLf & CDate(b.Дата).ToShortDateString & " - " & CDate(b.Время).ToShortTimeString & vbCrLf & vbCrLf & b.Сообщение
                    Else
                        RichTextBox1.Text = RichTextBox1.Text & vbCrLf & " ____________________________________ " & vbCrLf & CDate(b.Дата).ToShortDateString & " - " & CDate(b.Время).ToShortTimeString & vbCrLf & vbCrLf & b.Сообщение
                    End If

                Else
                    If RichTextBox1.Text.Length = 0 Then
                        RichTextBox1.Text = " ____________________________________ " & vbCrLf & CDate(b.Дата).ToShortDateString & vbCrLf & vbCrLf & b.Сообщение
                    Else
                        RichTextBox1.Text = RichTextBox1.Text & vbCrLf & " ____________________________________ " & vbCrLf & CDate(b.Дата).ToShortDateString & vbCrLf & vbCrLf & b.Сообщение
                    End If

                End If


            Next
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If MessageBox.Show("Сохранить нового клиента?", Рик, MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.No Then
            Return
        End If

        If TextBox1.Text.Length = 0 Then
            MessageBox.Show("Заполните поле!", Рик)
            Return
        End If
        Dim mo As New AllUpd
        Do While AllClass.SkypeКлиентПредложение Is Nothing
            mo.SkypeКлиентПредложениеAll()
        Loop
        Dim f = AllClass.SkypeКлиентПредложение.Where(Function(x) x.Клиент = Trim(TextBox1.Text) And x.Экспедитор = Экспедитор).FirstOrDefault()
        If f IsNot Nothing Then
            MessageBox.Show("Такой клиент уже существует!", Рик)
            Return
        End If
        Dim f4 As New IDNaz With {.Naz = Trim(TextBox1.Text)}
        Com1all.Add(f4)
        bscom1.ResetBindings(False)
        Dim d As String = Trim(TextBox1.Text)
        addAsync(d)
        TextBox1.Text = String.Empty

    End Sub
    Private Async Sub addAsync(ByVal d As String)
        Await Task.Run(Sub() add(d))
    End Sub
    Private Sub add(ByVal d As String)
        Dim mo As New AllUpd
        Using db As New dbAllDataContext()
            Dim f2 As New SkypeКлиентПредложение
            f2.Клиент = d
            f2.Экспедитор = Экспедитор
            f2.Дата = Now.ToShortDateString
            db.SkypeКлиентПредложение.InsertOnSubmit(f2)
            db.SubmitChanges()
        End Using
        mo.SkypeКлиентПредложениеAll()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If RichTextBox2.Text.Length = 0 Then
            MessageBox.Show("Заполните поле сообщения!", Рик)
            Return
        End If

        If ComboBox1.Text.Length = 0 Then
            MessageBox.Show("Выберите клиента!", Рик)
            Return
        End If
        Dim fn As String = RichTextBox2.Text
        AddNewAsync(ComboBox1.Text, fn)
        RichTextBox1.Text = RichTextBox1.Text & vbCrLf & " ____________________________________ " & vbCrLf & Now.ToShortDateString & " - " & Now.ToShortTimeString & vbCrLf & vbCrLf & fn
        RichTextBox2.Text = String.Empty
    End Sub
    Private Async Sub AddNewAsync(ByVal com1 As String, ByVal rch2 As String)
        Await Task.Run(Sub() AddNew(com1, rch2))
    End Sub
    Private Sub AddNew(ByVal com1 As String, ByVal rch2 As String)
        Dim mo As New AllUpd

        Using db As New dbAllDataContext()
            Dim f = db.SkypeКлиентПредложение.Where(Function(x) x.Клиент = com1 And x.Дата = Now.ToShortDateString And x.Экспедитор = Экспедитор).FirstOrDefault()
            If f IsNot Nothing Then
                If f.Время IsNot Nothing Then
                    f.Сообщение = f.Сообщение & vbCrLf & Now.ToShortDateString & " - " & f?.Время.Value.ToShortTimeString & vbCrLf & vbCrLf & rch2
                    f.Время = Now
                Else
                    f.Сообщение = f.Сообщение & vbCrLf & Now.ToShortDateString & vbCrLf & vbCrLf & rch2
                    f.Время = Now
                End If
                db.SubmitChanges()

            Else
                    Dim f1 As New SkypeКлиентПредложение
                f1.Дата = Now
                f1.Время = Now
                f1.Клиент = com1
                f1.Экспедитор = Экспедитор
                f1.Сообщение = Now.ToShortDateString & " - " & Now.ToShortTimeString & vbCrLf & vbCrLf & rch2
                db.SkypeКлиентПредложение.InsertOnSubmit(f1)
                db.SubmitChanges()

            End If
        End Using
        mo.SkypeКлиентПредложениеAll()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If TextBox2.Text.Length = 0 Then Return
        Dim mo As New AllUpd
        Do While AllClass.SkypeКлиентПредложение Is Nothing
            mo.SkypeКлиентПредложениеAll()
        Loop
        Dim txt2 As String = TextBox2.Text
        RichTextBox1.Text = String.Empty
        RichTextBox2.Text = String.Empty
        Dim f3 As String = ComboBox1.Text
        Dim f As New List(Of SkypeКлиентПредложение)
        If ComboBox1.Text.Length > 0 Then

            f = (From x In AllClass.SkypeКлиентПредложение
                 Order By x.Дата Descending
                 Where x?.Сообщение?.ToUpper.Contains(txt2.ToUpper) And x.Клиент = f3 And x.Экспедитор = Экспедитор
                 Select x).ToList()
        Else
            f = (From x In AllClass.SkypeКлиентПредложение
                 Order By x.Дата Descending
                 Where x?.Сообщение?.ToUpper.Contains(txt2.ToUpper) And x.Экспедитор = Экспедитор
                 Select x).ToList()
        End If



        If f IsNot Nothing Then
            For Each b In f
                If b.Дата Is Nothing Then Continue For
                If RichTextBox1.Text.Length = 0 Then
                    If b.Время Is Nothing Then
                        RichTextBox1.Text = " \-------------- Начало --------------" & vbCrLf & CDate(b.Дата).ToShortDateString & ", Клиент - " & b.Клиент _
                               & vbCrLf & "___________________________" & vbCrLf & b.Сообщение & vbCrLf & " -------------- Конец --------------/"
                    Else
                        RichTextBox1.Text = " \-------------- Начало --------------" & vbCrLf & CDate(b.Дата).ToShortDateString & " - " & CDate(b.Время).ToShortTimeString & ", Клиент - " & b.Клиент _
                                 & vbCrLf & "___________________________" & vbCrLf & b.Сообщение & vbCrLf & " -------------- Конец --------------/"
                    End If

                Else
                    If b.Время Is Nothing Then
                        RichTextBox1.Text = RichTextBox1.Text & vbCrLf & vbCrLf & " \-------------- Начало --------------" & vbCrLf & vbCrLf & CDate(b.Дата).ToShortDateString & ", Клиент - " & b.Клиент _
                               & vbCrLf & "___________________________" & vbCrLf & b.Сообщение & vbCrLf & " -------------- Конец --------------/"
                    Else
                        RichTextBox1.Text = RichTextBox1.Text & vbCrLf & vbCrLf & " \-------------- Начало --------------" & vbCrLf & vbCrLf & CDate(b.Дата).ToShortDateString & " - " & CDate(b.Время).ToShortTimeString & ", Клиент - " & b.Клиент _
                                 & vbCrLf & "___________________________" & vbCrLf & b.Сообщение & vbCrLf & " -------------- Конец --------------/"
                    End If
                End If
            Next
        End If

    End Sub
End Class