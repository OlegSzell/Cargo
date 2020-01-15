Option Explicit On
Imports System.Data.OleDb
Public Class Отчет
    Public datatbl As DataTable
    Public Clickоплата As String = ""

    Public назворг As String = ""
    Dim ds3, ds4 As DataTable
    Public ID As Integer = 0

    Private Sub Отчет_FormClosed(sender As Object, e As FormClosedEventArgs) Handles MyBase.FormClosed
        MaskedTextBox1.Text = ""
        MaskedTextBox2.Text = ""

        НомРейса = Nothing
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        MaskedTextBox1.Text = ""
        MaskedTextBox2.Text = ""

        НомРейса = Nothing
        Me.Close()
    End Sub

    Private Sub Отчет_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If bl = True Then 'переменная для определия откуда запущена форма
            Label3.Text = НомРейса
        Else
            НомРейса = Рейс.НомРес
            Label3.Text = НомРейса
        End If

        If Not ds3 Is Nothing Then
            ds3.Clear()
        End If
        If Not ds4 Is Nothing Then
            ds4.Clear()
        End If

        ОсткиОплат()

        UpdateGrid()

    End Sub
    Public Sub ОсткиОплат()
        'рейс клиента
        Dim strsql As String = "SELECT ОстатокОплаты,СтоимостьФрахта,Валюта FROM РейсыКлиента WHERE НомерРейса=" & НомРейса & ""
        Dim ds As DataTable = Selects3(strsql)

        If ds.Rows(0).Item(0).ToString = "" Then
            Label9.Text = "Рейс не оплачен!"
            Label9.ForeColor = Color.Red
        ElseIf ds.Rows(0).Item(0).ToString = "0" Then
            Label9.Text = "Рейс оплачен!"
            Label9.ForeColor = Color.Green
        Else
            'Dim d, d1 As Double
            'If ds.Rows(0).Item(0).ToString.Contains(".") Then
            '    ds.Rows(0).Item(0) = Replace(ds.Rows(0).Item(0).ToString, ".", ",")
            'End If
            'If ds.Rows(0).Item(1).ToString.Contains(".") Then
            '    ds.Rows(0).Item(1) = Replace(ds.Rows(0).Item(1).ToString, ".", ",")
            'End If

            'd = CDbl(ds.Rows(0).Item(0).ToString)
            'd1 = CDbl(ds.Rows(0).Item(1).ToString)
            Label9.Text = ds.Rows(0).Item(0).ToString & " " & ds.Rows(0).Item(2).ToString
        End If
        'рейс перевозчика
        Dim strsql1 As String = "SELECT ОстатокОплаты,СтоимостьФрахта,Валюта FROM РейсыПеревозчика WHERE НомерРейса=" & НомРейса & ""
        Dim ds1 As DataTable = Selects3(strsql1)

        If ds1.Rows(0).Item(0).ToString = "" Then
            Label8.Text = "Рейс не оплачен!"
            Label8.ForeColor = Color.Red
        ElseIf ds1.Rows(0).Item(0).ToString = "0" Then
            Label8.Text = "Рейс оплачен!"
            Label8.ForeColor = Color.Green
        Else
            'Dim d, d1 As Double
            'If ds1.Rows(0).Item(0).ToString.Contains(".") Then
            '    ds1.Rows(0).Item(0) = Replace(ds1.Rows(0).Item(0).ToString, ".", ",")
            'End If
            'If ds1.Rows(0).Item(1).ToString.Contains(".") Then
            '    ds1.Rows(0).Item(1) = Replace(ds1.Rows(0).Item(1).ToString, ".", ",")
            'End If

            'd = CDbl(ds1.Rows(0).Item(0).ToString)
            'd1 = CDbl(ds1.Rows(0).Item(1).ToString)
            Label8.Text = ds1.Rows(0).Item(0).ToString & " " & ds1.Rows(0).Item(2).ToString
        End If


    End Sub
    Private Sub UpdateGrid()
        Dim strsql As String = "SELECT ДатаОтправкиДоков FROM РейсыКлиента WHERE НомерРейса=" & НомРейса & ""
        Dim ds As DataTable = Selects3(strsql)

        If ds.Rows(0).Item(0).ToString <> "" Then

            'If Not Grid1 Is Nothing Then
            '    Grid1.Clear()
            'End If
            MaskedTextBox2.Text = ds.Rows(0).Item(0).ToString

            Dim strsql3 As String = "SELECT НазвОрганизации as [Организац], НомерРейса as [Рейс],
Маршрут, НаименованиеГруза as [Груз], СтоимостьФрахта as [Фрахт], Валюта, СрокОплаты as [Срок оплаты],
УсловияОплаты as [Условия оплаты], ДатаОтправкиДоков as [Дата отпр доков], ДатаОплаты as [Дата оплаты]
FROM РейсыКлиента 
WHERE НомерРейса=" & НомРейса & ""

            ds3 = Selects3(strsql3)
            Grid1.DataSource = ds3
            'Grid1.Columns(1).Width = 60
            'Grid1.Columns(4).Width = 60
            'Grid1.Columns(5).Width = 60
            'Grid1.Columns(6).Width = 100
            'Grid1.Columns(8).Width = 100
            'Grid1.Columns(9).Width = 100
            'Grid1.Columns(7).Width = 100
            GridView(Grid1)

            Grid1.Rows(0).Cells(9).Style.Font = New Font(Grid1.DefaultCellStyle.Font, FontStyle.Bold)
        End If

        Dim strsql1 As String = "SELECT ДатаПолученияДоков FROM РейсыПеревозчика WHERE НомерРейса=" & НомРейса & ""
        Dim ds1 As DataTable = Selects3(strsql1)

        If ds1.Rows(0).Item(0).ToString <> "" Then


            MaskedTextBox1.Text = ds1.Rows(0).Item(0).ToString

            Dim strsql4 As String = "SELECT НазвОрганизации as [Организац], НомерРейса as [Рейс], Маршрут,
НаименованиеГруза as [Груз], СтоимостьФрахта as [Фрахт], Валюта, СрокОплаты as [Срок оплаты],
УсловияОплаты as [Условия оплаты], ДатаПолученияДоков as [Дата получ доков], ДатаОплаты as [Дата оплаты]
FROM РейсыПеревозчика 
WHERE НомерРейса=" & НомРейса & ""
            ds4 = Selects3(strsql4)
            Grid2.DataSource = ds4
            GridView(Grid2)
            'Grid2.Columns(1).Width = 60
            'Grid2.Columns(4).Width = 60
            'Grid2.Columns(5).Width = 60
            'Grid2.Columns(6).Width = 100
            'Grid2.Columns(8).Width = 100
            'Grid2.Columns(9).Width = 100
            'Grid2.Columns(7).Width = 100
            Grid1.ClearSelection()
            Grid2.Rows(0).Cells(9).Style.Font = New Font(Grid2.DefaultCellStyle.Font, FontStyle.Bold)
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        СортПеревоз()
        СортКлиент()
        UpdateGrid()
    End Sub
    Private Sub СортКлиент()
        Dim strsql, strsql1, str As String
        Dim ds As DataTable
        Dim d As Integer
        Dim df As Double

        If MaskedTextBox2.MaskCompleted = True Then
            strsql = "SELECT СрокОплаты, УсловияОплаты FROM РейсыКлиента WHERE НомерРейса=" & НомРейса & ""
            ds = Selects3(strsql)
            str = ds.Rows(0).Item(0).ToString
            Dim dat As Date = MaskedTextBox2.Text

            If str.Length > 2 Then 'отбираем количество дней оплаты
                If str.Length > 5 Then 'если цифр больше 5
                    MessageBox.Show("Число цифр больше 5, сообщите администратору!")
                    Exit Sub
                End If

                If str.Length = 3 Then
                    d = CType(Strings.Right(str, 1), Integer)
                Else
                    d = CType(Strings.Right(str, 2), Integer)
                End If
            Else
                d = CType(str, Integer)
            End If

            If ds.Rows(0).Item(1).ToString = "БанкПоДок" Or ds.Rows(0).Item(1).ToString = "БанкПоВыг" Then

                If d <= 5 Then 'отсортировываем календарные и рабочие дни
                    d += 2
                Else
                    df = Math.Round(d / 5)
                    d += (df * 2)
                End If
            End If

            dat = dat.AddDays(d)
            strsql1 = "UPDATE РейсыКлиента SET ДатаОплаты='" & dat & "', ДатаОтправкиДоков='" & MaskedTextBox2.Text & "' WHERE НомерРейса=" & НомРейса & ""
            Updates3(strsql1)
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Перевоз()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Кли()
    End Sub
    Public Sub Перевоз()
        Clickоплата = "Перевозчик"
        НомерРейса3 = Label3.Text
        Dim strsql As String = "SELECT Код,НазвОрганизации,СтоимостьФрахта,Валюта FROM РейсыПеревозчика WHERE НомерРейса=" & НомРейса & ""
        Dim ds As DataTable = Selects3(strsql)
        ID = ds.Rows(0).Item(0)
        назворг = ds.Rows(0).Item(1).ToString
        datatbl = ds
        Оплата.ShowDialog()
    End Sub
    Public Sub Кли()
        НомерРейса3 = Label3.Text
        Clickоплата = "Клиент"
        Dim strsql As String = "SELECT Код,НазвОрганизации,СтоимостьФрахта,Валюта FROM РейсыКлиента WHERE НомерРейса=" & НомРейса & ""
        Dim ds As DataTable = Selects3(strsql)
        ID = ds.Rows(0).Item(0)
        назворг = ds.Rows(0).Item(1).ToString
        datatbl = ds
        Оплата.ShowDialog()
    End Sub
    Private Sub СортПеревоз()
        Dim strsql, strsql1, str As String
        Dim ds As DataTable
        Dim d As Integer
        Dim df As Double

        If MaskedTextBox1.MaskCompleted = True Then
            strsql = "SELECT СрокОплаты, УсловияОплаты FROM РейсыПеревозчика WHERE НомерРейса=" & НомРейса & ""
            ds = Selects3(strsql)
            str = ds.Rows(0).Item(0).ToString
            Dim dat As Date = MaskedTextBox1.Text

            If str.Length > 2 Then 'отбираем количество дней оплаты
                If str.Length > 5 Then 'если цифр больше 5
                    MessageBox.Show("Число цифр больше 5, сообщите администратору!")
                    Exit Sub
                End If

                If str.Length = 3 Then
                    d = CType(Strings.Right(str, 1), Integer)
                Else
                    d = CType(Strings.Right(str, 2), Integer)
                End If
            Else
                d = CType(str, Integer)
            End If

            If ds.Rows(0).Item(1).ToString = "БанкПоДок" Or ds.Rows(0).Item(1).ToString = "БанкПоВыг" Then

                If d <= 5 Then 'отсортировываем календарные и рабочие дни
                    d += 2
                Else
                    df = Math.Round(d / 5)
                    d += (df * 2)
                End If
            End If

            dat = dat.AddDays(d)
            strsql1 = "UPDATE РейсыПеревозчика SET ДатаОплаты='" & dat & "', ДатаПолученияДоков='" & MaskedTextBox1.Text & "' WHERE НомерРейса=" & НомРейса & ""
            Updates3(strsql1)
        End If
    End Sub
End Class