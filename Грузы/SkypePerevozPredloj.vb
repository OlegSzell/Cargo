Public Class SkypePerevozPredloj
    Private Com1all As List(Of IDNaz)
    Private bscom1 As BindingSource
    Private Sub SkypePerevozPredloj_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Com1all = New List(Of IDNaz)
        bscom1 = New BindingSource
        bscom1.DataSource = Com1all
        ComboBox1.DataSource = bscom1
        ComboBox1.DisplayMember = "Naz"

        Dim mo As New AllUpd
        Do While AllClass.SkypeПеревозчикПредложение Is Nothing
            mo.SkypeПеревозчикПредложениеAll()
        Loop
        Dim f = AllClass.SkypeПеревозчикПредложение.OrderBy(Function(x) x.Перевозчик).Where(Function(x) x.Экспедитор = Экспедитор).Select(Function(x) x.Перевозчик).Distinct().ToList()
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
        Do While AllClass.SkypeПеревозчикПредложение Is Nothing
            mo.SkypeПеревозчикПредложениеAll()
        Loop
        Dim f = AllClass.SkypeПеревозчикПредложение.OrderBy(Function(x) x.Дата).Where(Function(x) x.Перевозчик = f2.Naz And x.Экспедитор = Экспедитор).ToList()
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
        If MessageBox.Show("Сохранить нового перевозчика?", Рик, MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.No Then
            Return
        End If

        If TextBox1.Text.Length = 0 Then
            MessageBox.Show("Заполните поле!", Рик)
            Return
        End If
        Dim mo As New AllUpd
        Do While AllClass.SkypeПеревозчикПредложение Is Nothing
            mo.SkypeПеревозчикПредложениеAll()
        Loop
        Dim f = AllClass.SkypeПеревозчикПредложение.Where(Function(x) x.Перевозчик = Trim(TextBox1.Text) And x.Экспедитор = Экспедитор).FirstOrDefault()
        If f IsNot Nothing Then
            MessageBox.Show("Такой перевозчик уже существует!", Рик)
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
            Dim f2 As New SkypeПеревозчикПредложение
            f2.Перевозчик = d
            f2.Экспедитор = Экспедитор
            f2.Дата = Now.ToShortDateString
            db.SkypeПеревозчикПредложение.InsertOnSubmit(f2)
            db.SubmitChanges()
        End Using
        mo.SkypeПеревозчикПредложениеAll()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If RichTextBox2.Text.Length = 0 Then
            MessageBox.Show("Заполните поле сообщения!", Рик)
            Return
        End If

        If ComboBox1.Text.Length = 0 Then
            MessageBox.Show("Выберите перевозчика!", Рик)
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
            Dim f = db.SkypeПеревозчикПредложение.Where(Function(x) x.Перевозчик = com1 And x.Дата = Now.ToShortDateString And x.Экспедитор = Экспедитор).FirstOrDefault()
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
                Dim f1 As New SkypeПеревозчикПредложение
                f1.Дата = Now
                f1.Время = Now
                f1.Перевозчик = com1
                f1.Экспедитор = Экспедитор
                f1.Сообщение = Now.ToShortDateString & " - " & Now.ToShortTimeString & vbCrLf & vbCrLf & rch2
                db.SkypeПеревозчикПредложение.InsertOnSubmit(f1)
                db.SubmitChanges()

            End If
        End Using
        mo.SkypeПеревозчикПредложениеAll()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If TextBox2.Text.Length = 0 Then Return
        Dim mo As New AllUpd
        Do While AllClass.SkypeПеревозчикПредложение Is Nothing
            mo.SkypeПеревозчикПредложениеAll()
        Loop
        Dim txt2 As String = TextBox2.Text
        RichTextBox1.Text = String.Empty
        RichTextBox2.Text = String.Empty
        Dim f3 As String = ComboBox1.Text
        Dim f As New List(Of SkypeПеревозчикПредложение)
        If ComboBox1.Text.Length > 0 Then

            f = (From x In AllClass.SkypeПеревозчикПредложение
                 Order By x.Дата Descending
                 Where x?.Сообщение?.ToUpper.Contains(txt2.ToUpper) And x.Перевозчик = f3 And x.Экспедитор = Экспедитор
                 Select x).ToList()
        Else
            f = (From x In AllClass.SkypeПеревозчикПредложение
                 Order By x.Дата Descending
                 Where x?.Сообщение?.ToUpper.Contains(txt2.ToUpper) And x.Экспедитор = Экспедитор
                 Select x).ToList()
        End If



        If f IsNot Nothing Then
            For Each b In f
                If b.Дата Is Nothing Then Continue For
                If RichTextBox1.Text.Length = 0 Then
                    If b.Время Is Nothing Then
                        RichTextBox1.Text = " \-------------- Начало --------------" & vbCrLf & CDate(b.Дата).ToShortDateString & ", Перевозчик - " & b.Перевозчик _
                               & vbCrLf & "___________________________" & vbCrLf & b.Сообщение & vbCrLf & " -------------- Конец --------------/"
                    Else
                        RichTextBox1.Text = " \-------------- Начало --------------" & vbCrLf & CDate(b.Дата).ToShortDateString & " - " & CDate(b.Время).ToShortTimeString & ", Перевозчик - " & b.Перевозчик _
                                 & vbCrLf & "___________________________" & vbCrLf & b.Сообщение & vbCrLf & " -------------- Конец --------------/"
                    End If

                Else
                    If b.Время Is Nothing Then
                        RichTextBox1.Text = RichTextBox1.Text & vbCrLf & vbCrLf & " \-------------- Начало --------------" & vbCrLf & vbCrLf & CDate(b.Дата).ToShortDateString & ", Перевозчик - " & b.Перевозчик _
                               & vbCrLf & "___________________________" & vbCrLf & b.Сообщение & vbCrLf & " -------------- Конец --------------/"
                    Else
                        RichTextBox1.Text = RichTextBox1.Text & vbCrLf & vbCrLf & " \-------------- Начало --------------" & vbCrLf & vbCrLf & CDate(b.Дата).ToShortDateString & " - " & CDate(b.Время).ToShortTimeString & ", Перевозчик - " & b.Перевозчик _
                                 & vbCrLf & "___________________________" & vbCrLf & b.Сообщение & vbCrLf & " -------------- Конец --------------/"
                    End If
                End If
            Next
        End If
    End Sub
End Class