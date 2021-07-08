Option Explicit On
Imports System.Data.OleDb
Public Class ИзменПеревоз
    Public Da As New OleDbDataAdapter 'Адаптер
    Public Ds As New DataSet 'Пустой набор записей
    Dim tbl As New DataTable
    Dim cb As OleDb.OleDbCommandBuilder
    Dim Рик As String = "ООО Рикманс"
    Dim загр, strsql As String
    Dim ind As Boolean = False


    Private Sub ИзменПеревоз_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        TextBox5.Enabled = False
        TextBox4.Enabled = False

        дан()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()
    End Sub

    Public Sub дан()
        Dim mo As New AllUpd
        Do While AllClass.ПеревозчикиБаза Is Nothing
            mo.ПеревозчикиБазаAll()
        Loop

        Dim f = AllClass.ПеревозчикиБаза.Where(Function(x) x.ID = idtabl).FirstOrDefault()
        If f IsNot Nothing Then
            With f
                TextBox1.Text = .Форма_собственности
                TextBox2.Text = .Наименование_фирмы
                TextBox3.Text = .Контактное_лицо
                TextBox4.Text = .Регионы
                TextBox5.Text = .Страны_перевозок
                TextBox6.Text = .Телефоны
                TextBox7.Text = .Объем
                TextBox8.Text = .Тоннаж
                TextBox9.Text = .Вид_авто
                TextBox10.Text = .Кол_во_авто
                TextBox11.Text = .ADR
                TextBox12.Text = .Города
                TextBox13.Text = .Примечание
                TextBox14.Text = .Ставка
            End With

        End If

        'Dim ds As DataTable = Selects3(StrSql:="SELECT * FROM ПЕревозчикиБаза WHERE ID=" & idtabl & "")


        'TextBox1.Text = ds.Rows(0).Item(1).ToString
        'TextBox2.Text = ds.Rows(0).Item(2).ToString
        'TextBox3.Text = ds.Rows(0).Item(3).ToString
        'TextBox4.Text = ds.Rows(0).Item(6).ToString
        'TextBox5.Text = ds.Rows(0).Item(5).ToString
        'TextBox6.Text = ds.Rows(0).Item(4).ToString
        'TextBox7.Text = ds.Rows(0).Item(12).ToString
        'TextBox8.Text = ds.Rows(0).Item(11).ToString
        'TextBox9.Text = ds.Rows(0).Item(10).ToString
        'TextBox10.Text = ds.Rows(0).Item(9).ToString
        'TextBox11.Text = ds.Rows(0).Item(8).ToString
        'TextBox12.Text = ds.Rows(0).Item(7).ToString
        'TextBox13.Text = ds.Rows(0).Item(14).ToString
        'TextBox14.Text = ds.Rows(0).Item(13).ToString




    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked = True Then
            CheckBox2.Checked = False
            Dim f As New СтраныПерЛист
            f.Show()
        Else
            СтраныПерЛист.Close()
        End If
    End Sub

    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox2.CheckedChanged
        If CheckBox2.Checked = True Then
            CheckBox1.Checked = False

            Dim f As New РегионыПерЛист
            f.Show()

        Else
            РегионыПерЛист.Close()
        End If
    End Sub
    Private Sub Ok()
        If MessageBox.Show("Сохранить изменения?", Рик, MessageBoxButtons.OKCancel, MessageBoxIcon.Question) = DialogResult.Cancel Then
            Exit Sub
            Me.Close()
        End If
        TextBox1.BeginInvoke(New MethodInvoker(Sub() Upd()))

        '        Updates3(stroka:="UPDATE ПеревозчикиБаза SET [Форма собственности]='" & TextBox1.Text & "', [Наименование фирмы]='" & TextBox2.Text & "',
        '[Контактное лицо]='" & TextBox3.Text & "',[Телефоны]='" & TextBox6.Text & "',[Города]='" & TextBox12.Text & "',
        '[ADR]='" & TextBox11.Text & "',[Кол-во авто]='" & TextBox10.Text & "',[Вид_авто]='" & TextBox9.Text & "',
        '[Тоннаж]='" & TextBox8.Text & "',[Объем]='" & TextBox7.Text & "',[Ставка]='" & TextBox14.Text & "',[Примечание]='" & TextBox13.Text & "'
        'WHERE ID=" & idtabl & "")


    End Sub
    Private Sub Upd()
        Dim mo As New AllUpd
        Using db As New dbAllDataContext(_cn3)
            Dim f = db.ПеревозчикиБаза.Where(Function(x) x.ID = idtabl).FirstOrDefault()
            If f IsNot Nothing Then
                With f
                    .Форма_собственности = TextBox1.Text
                    .Наименование_фирмы = TextBox2.Text
                    .Контактное_лицо = TextBox3.Text
                    .Регионы = TextBox4.Text
                    .Страны_перевозок = TextBox5.Text
                    .Телефоны = TextBox6.Text
                    .Объем = TextBox7.Text
                    .Тоннаж = TextBox8.Text
                    .Вид_авто = TextBox9.Text
                    .Кол_во_авто = TextBox10.Text
                    .ADR = TextBox11.Text
                    .Города = TextBox12.Text
                    .Примечание = TextBox13.Text
                    .Ставка = TextBox14.Text
                End With
                db.SubmitChanges()
                mo.ПеревозчикиБазаAll()
            End If
        End Using

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Ok()
        MessageBox.Show("Данные изменены!", Рик)
        If pr2 = 0 Then
            ПоискПеревозчиков.ОбнКом1()
        End If

        Me.Close()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim f As New ImageForm(False, idtabl, True)
        f.ShowDialog()
    End Sub

    Private Sub TextBox5_DoubleClick(sender As Object, e As EventArgs) Handles TextBox5.DoubleClick
        TextBox5.Enabled = True

    End Sub
End Class