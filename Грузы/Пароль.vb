Option Explicit On
Imports System.Data.OleDb
Imports System.Data.SqlClient
Public Class Пароль
    Dim strsql As String
    Dim fd As String
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        ClickOK()

    End Sub
    Private Sub ClickOK()
        Dim ds1 As DataTable = Selects3(StrSql:="SELECT Парол FROM Пароли WHERE Логин='" & ComboBox1.Text & "'")
        'strsql = "SELECT Парол FROM Пароли WHERE Логин='" & ComboBox1.Text & "'"
        'Dim c1 As New OleDbCommand With {
        '    .Connection = conn,
        '    .CommandText = strsql
        '}
        'Dim ds1 As New DataTable
        'Dim da1 As New OleDbDataAdapter(c1)
        'da1.Fill(ds1)

        If ds1.Rows(0).Item(0).ToString = TextBox1.Text Then
            'MessageBox.Show("Пароль принят!")
            Экспедитор = ComboBox1.Text
            Me.Close()
            TextBox1.Text = ""
            CheckBox1.Checked = False
        Else
            MessageBox.Show("Пароль НЕ принят! Повторите")
            TextBox1.Text = ""
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
        Updates3(stroka:="INSERT INTO Пароли(Логин,Парол) VALUES('" & ComboBox1.Text & "','" & TextBox1.Text & "')")
        MessageBox.Show("Ваши данные зарегистрированы!")
        ComboBox1.Text = ""
        TextBox1.Text = ""
        Combo1()
    End Sub
    Private Sub Combo1()

        'Dim c1 As New OleDbCommand With {
        '    .Connection = conn,
        '    .CommandText = StrSql1
        '}
        'Dim ds1 As New DataTable
        'Dim da1 As New OleDbDataAdapter(c1)
        'da1.Fill(ds1)

        'Using (SqlConnection connection = New SqlConnection(ConnString))
        '    {
        '        //SqlCommand cmd = New SqlCommand(d, connection);
        '        connection.Open();

        '        SqlDataAdapter adapter = New SqlDataAdapter(d, connection); //создаем адаптер для связи с данными. 
        '       /* adapter.SelectCommand = New SqlCommand(d, connection); */// указываем ему команду для выборки

        '        //SqlDataReader dr = cmd.ExecuteReader();
        '        DataSet ds = New DataSet();

        '        Try
        '        {
        '            adapter.Fill(ds);//заполняем датасет с помощьб адаптера (будет исполнена команда на выборку)
        '            dt = ds.Tables[0];


        '            Return dt;

        Dim ds1 As DataTable = Selects3(StrSql:="SELECT Логин FROM Пароли ORDER BY Логин")

        Me.ComboBox1.Items.Clear()
        Me.ComboBox1.AutoCompleteCustomSource.Clear()

        For Each r As DataRow In ds1.Rows
            Me.ComboBox1.AutoCompleteCustomSource.Add(r.Item(0).ToString())
            Me.ComboBox1.Items.Add(r(0).ToString)
        Next

    End Sub
    Private Sub Пароль_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim f As Boolean = False
        Dim f1 As Boolean = False
        fd = My.Computer.Name.ToString
        Экспедитор = ""



        'Catch ex As Exception
        '            MessageBox.Show("Подключите диск U и нажмите OK", Рик)
        '            f1 = True
        '        Finally
        '            If f1 = True Then
        '        Try
        '            conn3.Open()
        '        Catch ex As Exception
        '            MessageBox.Show("Вы не подключили диск U.Приложение будет закрыто!", Рик)
        '            Me.Close()
        '            MDIParent1.Close()
        '            f = True
        '        End Try
        '    End If

        'End Try


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
            'Dim ds As DataTable = Selects3(StrSql:="SELECT Парол FROM Пароли WHERE Логин='Олег'")
            Dim ds As DataTable = Selects3(StrSql:="SELECT Парол FROM Пароли WHERE Логин='Олег'")
            TextBox1.Text = ds.Rows(0).Item(0).ToString
            Button1.PerformClick()
        End If
    End Sub
End Class