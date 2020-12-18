Imports System.ComponentModel

Public Class ИзменПеревозчика
    Private cod As Grid1ЖурналClass
    Private lst2 As BindingList(Of Страна)
    Private bslst2 As BindingSource
    Private lst1 As BindingList(Of String)
    Private bslst1 As BindingSource


    Sub New(ByVal _b As Grid1ЖурналClass)
        cod = _b
        ' Этот вызов является обязательным для конструктора.
        InitializeComponent()

        ' Добавить код инициализации после вызова InitializeComponent().

    End Sub

    Private Sub ИзменПеревозчика_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        lst2 = New BindingList(Of Страна)
        bslst2 = New BindingSource With {.DataSource = lst2}
        ListBox2.DataSource = bslst2
        ListBox2.DisplayMember = "Страна"

        lst1 = New BindingList(Of String)
        bslst1 = New BindingSource
        bslst1.DataSource = lst1
        ListBox1.DataSource = bslst1

        Dim mo As New AllUpd
        Do While AllClass.Страна Is Nothing
            mo.СтранаAll()
        Loop
        Do While AllClass.ПеревозчикиБаза Is Nothing
            mo.ПеревозчикиБазаAll()
        Loop

        Dim f = AllClass.Страна.OrderBy(Function(x) x.Страна).ToList()
        If f IsNot Nothing Then
            For Each b In f
                lst2.Add(b)
            Next
        End If

        Dim f2 = AllClass.ПеревозчикиБаза.Where(Function(x) x.ID = cod.КодПеревозчика).FirstOrDefault()

        Dim f3 = f2.Страны_перевозок.Split(",").ToList()
        If f3 IsNot Nothing Then
            For Each b In f3
                lst1.Add(Trim(b))
            Next
        End If

        TextBox12.Text = f2.Форма_собственности
        TextBox11.Text = f2.Наименование_фирмы
        TextBox2.Text = f2.Контактное_лицо
        TextBox3.Text = f2.Телефоны
        TextBox10.Text = f2.Вид_авто
        TextBox5.Text = f2.Тоннаж
        TextBox4.Text = f2.Кол_во_авто
        TextBox8.Text = f2.Примечание
        TextBox1.Text = f2.ADR
        TextBox7.Text = f2.Объем
        TextBox6.Text = f2.Ставка
        If f2.ДатаИзменения Is Nothing Then
            TextBox9.Text = String.Empty
        Else
            TextBox9.Text = f2.ДатаИзменения
        End If



    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If ListBox1.SelectedIndex = -1 Then
            Return
        End If

        ListBox1.BeginUpdate()
        lst1.Remove(ListBox1.SelectedItem)
        ListBox1.EndUpdate()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If ListBox2.SelectedIndex = -1 Then
            Return
        End If


        Dim k As Страна = ListBox2.SelectedItem
        If lst1.Contains(k.Страна) = False Then
            ListBox2.BeginUpdate()
            lst1.Add(k.Страна)
            ListBox2.EndUpdate()
        End If

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Close()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If MessageBox.Show("Сохранить изменения?", Рик, MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.No Then
            Return
        End If

        Dim strLst1 As String = Nothing
        For Each b In lst1

            If strLst1 Is Nothing Then
                strLst1 = b
            Else
                strLst1 = strLst1 & "," & b
            End If

        Next

        Using db As New dbAllDataContext
            Dim f = db.ПеревозчикиБаза.Where(Function(x) x.ID = cod.КодПеревозчика).FirstOrDefault()
            If f IsNot Nothing Then
                With f
                    .ADR = TextBox1.Text
                    .Форма_собственности = TextBox12.Text
                    .Наименование_фирмы = TextBox11.Text
                    .Контактное_лицо = TextBox2.Text
                    .Телефоны = TextBox3.Text
                    .Вид_авто = TextBox10.Text
                    .Тоннаж = TextBox5.Text
                    .Кол_во_авто = TextBox4.Text
                    .Примечание = TextBox8.Text
                    .Объем = TextBox7.Text
                    .Ставка = TextBox6.Text
                    .ДатаИзменения = Now
                    .Страны_перевозок = strLst1
                    db.SubmitChanges()
                End With
            End If
        End Using
        Dim mo As New AllUpd
        mo.ПеревозчикиБазаAllAsync()
        MessageBox.Show("Сохранение внесено!?", Рик)
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Dim f As New ImageForm(False, cod.КодПеревозчика)
        f.ShowDialog()
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Dim mo As New AllUpd
        Do While AllClass.ПеревозчикиБаза Is Nothing
            mo.ПеревозчикиБазаAll()
        Loop

        Dim f = AllClass.ПеревозчикиБаза.Where(Function(x) x.ID = cod.КодПеревозчика).Select(Function(x) x.ФотоДанные).FirstOrDefault()
        If f IsNot Nothing Then
            Dim f1 As New ImageForm(False, cod.КодПеревозчика, True)
            f1.ShowDialog()
        Else
            MessageBox.Show("Нет фото!", Рик)
        End If

    End Sub
End Class