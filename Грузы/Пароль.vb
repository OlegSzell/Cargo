Option Explicit On
Imports System.ComponentModel
Imports System.Data.OleDb
Imports System.Data.SqlClient
Imports System.Net.NetworkInformation

Public Class Пароль
    Dim strsql As String
    Dim fd As String
    Dim Flag As Boolean = False
    Public flg As Boolean = False
    Dim clicks As Integer = 0
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        ClickOK()

    End Sub
    Private Sub ClickOK()

        Dim macAddr = (From nic In NetworkInterface.GetAllNetworkInterfaces()
                       Where nic.OperationalStatus = OperationalStatus.Up
                       Select nic.GetPhysicalAddress()).FirstOrDefault().ToString

        Dim nmComp = Environment.MachineName

        'Dim mk As String
        'For Each b As NetworkInterface In NetworkInterface.GetAllNetworkInterfaces
        '    If b.OperationalStatus = OperationalStatus.Up Then
        '        mk = b.GetPhysicalAddress.ToString
        '    End If
        'Next




        Dim mo As New AllUpd
        Do While AllClass.Пароли Is Nothing
            mo.ПаролиAll()
        Loop
        Dim f = AllClass.Пароли.Where(Function(x) x.Логин = ComboBox1.Text).Select(Function(x) x).FirstOrDefault()

        'Dim ds1 As DataTable = Selects3(StrSql:="SELECT Парол FROM Пароли WHERE Логин='" & ComboBox1.Text & "'")
        Dim par As String = TextBox1.Text
        If ComboBox1.Text.Length = 0 Then
            MessageBox.Show("Выберите пользователя!", Рик)
            Return
        End If

        If f IsNot Nothing Then
            If f.Парол = par Then
                MessageBox.Show("Пароль принят!")
                Flag = True
                flg = True
                Экспедитор = ComboBox1.Text
                Me.Close()
                TextBox1.Text = ""
                CheckBox1.Checked = False
            Else

                MessageBox.Show("Пароль НЕ принят! Повторите")
                clicks += 1
                Using db As New dbAllDataContext()
                    Dim f7 = db.Пароли.Where(Function(x) x.Логин = ComboBox1.Text).Select(Function(x) x).FirstOrDefault()
                    If f7 IsNot Nothing Then


                        Dim f8 As New ПаролиВвод
                        With f8
                            .IDPassword = f7.Код
                            .ВводПароля = TextBox1.Text
                            .Логин = f7.Логин
                            .Дата = Now
                            .MacAdress = macAddr
                            .nameComp = nmComp
                            db.ПаролиВвод.InsertOnSubmit(f8)
                            db.SubmitChanges()
                        End With

                    End If
                End Using

                TextBox1.Text = String.Empty

                If clicks = 3 Then
                    Using db As New dbAllDataContext()
                        Dim f5 = db.Пароли.Where(Function(x) x.Логин = ComboBox1.Text).Select(Function(x) x).FirstOrDefault()
                        If f5 IsNot Nothing Then
                            f5.Блокировка = "True"
                            db.SubmitChanges()
                            MessageBox.Show("Пользователь заблоктирован!" & vbCrLf & " Обратаитесь к администратору")
                            Close()
                            MDIParent1.Close()
                        End If
                    End Using
                End If
            End If
        Else
            MessageBox.Show("Пароль НЕ принят! Повторите")

            clicks += 1
            Using db As New dbAllDataContext()
                Dim f7 = db.Пароли.Where(Function(x) x.Логин = ComboBox1.Text).Select(Function(x) x).FirstOrDefault()
                If f7 IsNot Nothing Then


                    Dim f8 As New ПаролиВвод
                    With f8
                        .IDPassword = f7.Код
                        .ВводПароля = TextBox1.Text
                        .Логин = f7.Логин
                        .Дата = Now
                        .MacAdress = macAddr
                        .nameComp = nmComp
                        db.ПаролиВвод.InsertOnSubmit(f8)
                        db.SubmitChanges()
                    End With

                End If
            End Using

            TextBox1.Text = String.Empty
            If clicks = 3 Then

                Using db As New dbAllDataContext()
                    Dim f5 = db.Пароли.Where(Function(x) x.Логин = ComboBox1.Text).Select(Function(x) x).FirstOrDefault()
                    If f5 IsNot Nothing Then
                        f5.Блокировка = "True"
                        db.SubmitChanges()
                        MessageBox.Show("Пользователь заблоктирован!" & vbCrLf & " Обратаитесь к администратору")
                        Close()
                        MDIParent1.Close()
                    End If
                End Using



            End If
        End If

    End Sub
    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked = True Then
            TextBox1.PasswordChar = ""
        Else
            TextBox1.PasswordChar = "*"
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If ComboBox1.Text.Length = 0 Then
            MessageBox.Show("Введите имя!", Рик)
            Return
        End If
        If TextBox1.Text.Length = 0 Then
            MessageBox.Show("Заполните поле пароль!", Рик)
            Return
        End If

        Using db As New dbAllDataContext()
            Dim f As New Пароли
            f.Логин = ComboBox1.Text
            f.Парол = TextBox1.Text
            db.Пароли.InsertOnSubmit(f)
            db.SubmitChanges()
        End Using
        'Updates3(stroka:="INSERT INTO Пароли(Логин,Парол) VALUES('" & ComboBox1.Text & "','" & TextBox1.Text & "')")
        MessageBox.Show("Ваши данные зарегистрированы!")
        ComboBox1.Text = ""
        TextBox1.Text = ""
        Combo1()
    End Sub
    Private Sub Combo1()
        Dim mo As New AllUpd
        Do While AllClass.Пароли Is Nothing
            mo.ПаролиAll()
        Loop
        Do While AllClass.ПаролиВвод Is Nothing
            mo.ПаролиВводAll()
        Loop
        Dim f As New List(Of String)
        Try
            f = (From x In AllClass.Пароли
                 Where x.Блокировка = "False"
                 Order By x.Логин
                 Select x.Логин).ToList()
        Catch ex As Exception
            Close()
            MDIParent1.Close()
        End Try


        If f IsNot Nothing Then
            If f.Count > 0 Then
                Me.ComboBox1.Items.Clear()
                Me.ComboBox1.AutoCompleteCustomSource.Clear()

                For Each r In f
                    Me.ComboBox1.AutoCompleteCustomSource.Add(r)
                    Me.ComboBox1.Items.Add(r)
                Next
            End If
        Else
            Close()
            MDIParent1.Close()
        End If



    End Sub
    Private Sub Пароль_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim f As Boolean = False
        Dim f1 As Boolean = False
        fd = My.Computer.Name.ToString
        Экспедитор = ""




        fall.IsBackground = True
        fall.Start()
        Combo1()

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Me.Close()
        MDIParent1.Close()
    End Sub

    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            ClickOK()
        End If
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        If ComboBox1.Text = "Олег" And fd = "OLEGLAPTOP" Then
            Dim mo As New AllUpd
            Do While AllClass.Пароли Is Nothing
                mo.ПаролиAll()
            Loop
            Dim f = AllClass.Пароли.Where(Function(x) x.Логин = "Олег").Select(Function(x) x).FirstOrDefault()
            If f IsNot Nothing Then
                TextBox1.Text = f.Парол
                Button1.PerformClick()
            Else
                MessageBox.Show("Пароль не принят!", Рик)

            End If

            'Dim ds As DataTable = Selects3(StrSql:="SELECT Парол FROM Пароли WHERE Логин='Олег'")
            'TextBox1.Text = ds.Rows(0).Item(0).ToString
            'Button1.PerformClick()
        End If
    End Sub

    Private Sub Пароль_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        If Flag = False Then
            MDIParent1.Close()
        End If
    End Sub
End Class