Public Class ИзменПорНомерКлиента2
    Private НомРес As Integer
    Private com3 As String
    Private com4 As String
    Public Rez As String = Nothing
    Sub New(ByVal _НомРес As Integer, ByVal _com3 As String, ByVal _com4 As String)
        НомРес = _НомРес
        com3 = _com3
        com4 = _com4

        InitializeComponent()

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If TextBox2.Text = "" Then

            MessageBox.Show("Введите новый порядковый номер рейса!", Рик)
            Exit Sub
        End If

        If IsNumeric(TextBox2.Text) = False Then
            MessageBox.Show("Введите корректно номер рейса!", Рик)
            Return
        End If

        Dim Num As Integer = CType(TextBox2.Text, Integer)
        Using db As New dbAllDataContext()


            If ОтКогоИзмен = 1 Then
                Dim f = db.РейсыКлиента.Where(Function(x) x.НомерРейса = НомРес).Select(Function(x) x).FirstOrDefault()
                f.КоличРейсов = Num
                db.SubmitChanges()

                If MessageBox.Show("Данные удачно изменены в базе!" & vbCrLf & "Внести изменения в файле эксель?", Рик, MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.No Then
                    Me.Close()
                    Return
                End If
                Me.Close()

                Rez = "Клиент"
                TextBox1.Text = ""
                TextBox2.Text = ""
            Else
                Dim f2 = db.РейсыПеревозчика.Where(Function(x) x.НомерРейса = НомРес).Select(Function(x) x).FirstOrDefault()
                f2.КоличРейсов = Num
                db.SubmitChanges()


                If MessageBox.Show("Данные удачно изменены в базе!" & vbCrLf & "Внести изменения в файле эксель?", Рик, MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.No Then
                    Me.Close()
                    Exit Sub
                End If
                Me.Close()

                Rez = "Перевоз"
                TextBox1.Text = ""
                TextBox2.Text = ""

            End If

        End Using
    End Sub
    Private Sub ИзменПорНомерКлиента2_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Using db As New dbAllDataContext()
            If ОтКогоИзмен = 1 Then
                Dim f = db.РейсыКлиента.Where(Function(x) x.НомерРейса = НомРес).Select(Function(x) x.КоличРейсов).FirstOrDefault()
                If f IsNot Nothing Then
                    TextBox1.Text = f
                    Label3.Text = com3
                End If

            Else
                Dim f = db.РейсыПеревозчика.Where(Function(x) x.НомерРейса = НомРес).Select(Function(x) x.КоличРейсов).FirstOrDefault()
                If f IsNot Nothing Then
                    TextBox1.Text = CType(f, String)
                    Label3.Text = com4
                End If

            End If
        End Using
    End Sub
End Class