Option Explicit On
Imports System.Data.OleDb
Public Class Оплата
    Dim кто As String
    Dim NomerRejsa As Integer
    Dim f As Boolean = False
    Dim insertinbaza, Оплаты As String
    Dim ID, СтрокиВходящие As Integer
    Dim ostatok As Double = 0

    Private Sub Оплата_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Label2.Text = НомРейса
        кто = Отчет.Clickоплата
        Label20.Text = кто
        ID = Отчет.ID
        Label19.Text = Отчет.назворг
        NomerRejsa = НомРейса
        Label23.Text = Отчет.datatbl.Rows(0).Item(2).ToString & " " & Отчет.datatbl.Rows(0).Item(3).ToString

        If кто = "Перевозчик" Then
            Историяперевоз()
            insertinbaza = "IDПер"
            Оплаты = "ОплатыПер"
        Else
            Историяклиент()
            insertinbaza = "IDКлиента"
            Оплаты = "ОплатыКлиент"
        End If
        Label24.Text = Math.Round(CDbl(Отчет.datatbl.Rows(0).Item(2).ToString) - ostatok, 2) & " " & Отчет.datatbl.Rows(0).Item(3).ToString


    End Sub
    Private Sub Историяперевоз()
        Dim strsql As String = "SELECT РейсыПеревозчика.НомерРейса, ОплатыПер.Рейс, ОплатыПер.ДатаОплаты, ОплатыПер.Сумма
FROM РейсыПеревозчика INNER JOIN ОплатыПер ON РейсыПеревозчика.Код = ОплатыПер.IDПер
WHERE РейсыПеревозчика.НомерРейса=" & NomerRejsa & ""
        Dim ds As DataTable = Selects3(strsql)

        If errds = 1 Then
            f = True 'использовать insert
            Exit Sub
        Else
            СтрокиВходящие = ds.Rows.Count
            ЗаполнОплаты(ds)
        End If

    End Sub
    Private Sub ЗаполнОплаты(ByVal ds As DataTable)

        For i As Integer = 1 To ds.Rows.Count
            Select Case i
                Case 1
                    CheckBox1.Checked = True
                Case 2
                    CheckBox2.Checked = True
                Case 3
                    CheckBox3.Checked = True
                Case 4
                    CheckBox4.Checked = True
                Case 5
                    CheckBox5.Checked = True
                Case 6
                    CheckBox6.Checked = True
                Case 7
                    CheckBox7.Checked = True
                Case 8
                    CheckBox8.Checked = True
            End Select

            Controls("TextBox" & i).Text = ds.Rows(i - 1).Item(3).ToString
            Controls("MaskedTextBox" & i).Text = ds.Rows(i - 1).Item(2).ToString
            If ds.Rows(i - 1).Item(3).ToString.Contains(".") Then
                ds.Rows(i - 1).Item(3) = Replace(ds.Rows(i - 1).Item(3).ToString, ".", ",")
            End If
            ostatok += CDbl(ds.Rows(i - 1).Item(3).ToString)
        Next


    End Sub
    Private Sub Историяклиент()
        Dim strsql As String = "SELECT РейсыКлиента.НомерРейса, ОплатыКлиент.Рейс, ОплатыКлиент.ДатаОплаты, ОплатыКлиент.Сумма
FROM РейсыКлиента INNER JOIN ОплатыКлиент ON РейсыКлиента.Код = ОплатыКлиент.IDКлиента
WHERE РейсыКлиента.НомерРейса =" & NomerRejsa & ""
        Dim ds As DataTable = Selects3(strsql)

        If errds = 1 Then
            Exit Sub
            f = True 'использовать insert
        Else
            СтрокиВходящие = ds.Rows.Count
            ЗаполнОплаты(ds)
        End If
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged

        If CheckBox1.Checked = True Then
            TextBox1.Enabled = True
            MaskedTextBox1.Enabled = True
        Else
            TextBox1.Enabled = False
            MaskedTextBox1.Enabled = False
            TextBox1.Text = ""
            MaskedTextBox1.Text = ""
        End If

    End Sub

    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox2.CheckedChanged
        If CheckBox2.Checked = True Then
            TextBox2.Enabled = True
            MaskedTextBox2.Enabled = True
        Else
            TextBox2.Enabled = False
            MaskedTextBox2.Enabled = False
            TextBox2.Text = ""
            MaskedTextBox2.Text = ""
        End If
    End Sub

    Private Sub CheckBox3_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox3.CheckedChanged
        If CheckBox3.Checked = True Then
            TextBox3.Enabled = True
            MaskedTextBox3.Enabled = True
        Else
            TextBox3.Enabled = False
            MaskedTextBox3.Enabled = False
            TextBox3.Text = ""
            MaskedTextBox3.Text = ""
        End If
    End Sub

    Private Sub CheckBox4_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox4.CheckedChanged
        If CheckBox4.Checked = True Then
            TextBox4.Enabled = True
            MaskedTextBox4.Enabled = True
        Else
            TextBox4.Enabled = False
            MaskedTextBox4.Enabled = False
            TextBox4.Text = ""
            MaskedTextBox4.Text = ""
        End If
    End Sub

    Private Sub CheckBox5_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox5.CheckedChanged
        If CheckBox5.Checked = True Then
            TextBox5.Enabled = True
            MaskedTextBox5.Enabled = True
        Else
            TextBox5.Enabled = False
            MaskedTextBox5.Enabled = False
            TextBox5.Text = ""
            MaskedTextBox5.Text = ""
        End If
    End Sub

    Private Sub CheckBox6_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox6.CheckedChanged
        If CheckBox6.Checked = True Then
            TextBox6.Enabled = True
            MaskedTextBox6.Enabled = True
        Else
            TextBox6.Enabled = False
            MaskedTextBox6.Enabled = False
            TextBox6.Text = ""
            MaskedTextBox6.Text = ""
        End If
    End Sub

    Private Sub CheckBox7_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox7.CheckedChanged
        If CheckBox7.Checked = True Then
            TextBox7.Enabled = True
            MaskedTextBox7.Enabled = True
        Else
            TextBox7.Enabled = False
            MaskedTextBox7.Enabled = False
            TextBox7.Text = ""
            MaskedTextBox7.Text = ""
        End If
    End Sub

    Private Sub CheckBox8_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox8.CheckedChanged
        If CheckBox8.Checked = True Then
            TextBox8.Enabled = True
            MaskedTextBox8.Enabled = True
        Else
            TextBox8.Enabled = False
            MaskedTextBox8.Enabled = False
            TextBox8.Text = ""
            MaskedTextBox8.Text = ""
        End If
    End Sub
    Private Function proverka() As Integer

        If CheckBox1.Checked = True And CheckBox2.Checked = True And CheckBox3.Checked = True And CheckBox4.Checked = True And CheckBox5.Checked = True And CheckBox6.Checked = True And CheckBox7.Checked = True And CheckBox8.Checked = True Then
            Return 8
        ElseIf CheckBox1.Checked = True And CheckBox2.Checked = True And CheckBox3.Checked = True And CheckBox4.Checked = True And CheckBox5.Checked = True And CheckBox6.Checked = True And CheckBox7.Checked = True And CheckBox8.Checked = False Then
            Return 7
        ElseIf CheckBox1.Checked = True And CheckBox2.Checked = True And CheckBox3.Checked = True And CheckBox4.Checked = True And CheckBox5.Checked = True And CheckBox6.Checked = True And CheckBox7.Checked = False And CheckBox8.Checked = False Then
            Return 6
        ElseIf CheckBox1.Checked = True And CheckBox2.Checked = True And CheckBox3.Checked = True And CheckBox4.Checked = True And CheckBox5.Checked = True And CheckBox6.Checked = False And CheckBox7.Checked = False And CheckBox8.Checked = False Then
            Return 5
        ElseIf CheckBox1.Checked = True And CheckBox2.Checked = True And CheckBox3.Checked = True And CheckBox4.Checked = True And CheckBox5.Checked = False And CheckBox6.Checked = False And CheckBox7.Checked = False And CheckBox8.Checked = False Then
            Return 4
        ElseIf CheckBox1.Checked = True And CheckBox2.Checked = True And CheckBox3.Checked = True And CheckBox4.Checked = False And CheckBox5.Checked = False And CheckBox6.Checked = False And CheckBox7.Checked = False And CheckBox8.Checked = False Then
            Return 3
        ElseIf CheckBox1.Checked = True And CheckBox2.Checked = True And CheckBox3.Checked = False And CheckBox4.Checked = False And CheckBox5.Checked = False And CheckBox6.Checked = False And CheckBox7.Checked = False And CheckBox8.Checked = False Then
            Return 2
        ElseIf CheckBox1.Checked = True And CheckBox2.Checked = False And CheckBox3.Checked = False And CheckBox4.Checked = False And CheckBox5.Checked = False And CheckBox6.Checked = False And CheckBox7.Checked = False And CheckBox8.Checked = False Then
            Return 1
        Else
            Return 0
        End If
    End Function
    Private Function ПроверкаПоЗаполненности()
        If CheckBox1.Checked = True Then
            If TextBox1.Text = "" Then
                Return 1
            End If
            If MaskedTextBox1.MaskCompleted = False Then
                Return 1
            End If
        End If

        If CheckBox2.Checked = True Then
            If TextBox2.Text = "" Then
                Return 1
            End If
            If MaskedTextBox2.MaskCompleted = False Then
                Return 1
            End If
        End If


        If CheckBox3.Checked = True Then
            If TextBox3.Text = "" Then
                Return 1
            End If
            If MaskedTextBox3.MaskCompleted = False Then
                Return 1
            End If
        End If

        If CheckBox4.Checked = True Then
            If TextBox4.Text = "" Then
                Return 1
            End If
            If MaskedTextBox4.MaskCompleted = False Then
                Return 1
            End If
        End If


        If CheckBox5.Checked = True Then
            If TextBox5.Text = "" Then
                Return 1
            End If
            If MaskedTextBox5.MaskCompleted = False Then
                Return 1
            End If
        End If

        If CheckBox6.Checked = True Then
            If TextBox6.Text = "" Then
                Return 1
            End If
            If MaskedTextBox6.MaskCompleted = False Then
                Return 1
            End If
        End If

        If CheckBox7.Checked = True Then
            If TextBox7.Text = "" Then
                Return 1
            End If
            If MaskedTextBox7.MaskCompleted = False Then
                Return 1
            End If
        End If

        If CheckBox8.Checked = True Then
            If TextBox8.Text = "" Then
                Return 1
            End If
            If MaskedTextBox8.MaskCompleted = False Then
                Return 1
            End If
        End If
        Return 0
    End Function
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Dim k As Integer = ПроверкаПоЗаполненности()
        If k = 1 Then
            MessageBox.Show("Заполните все поля!", Рик)
            Exit Sub
        End If

        Dim g As Integer = proverka()
        Dim strsql As String

        'If f = True Then
        'For i As Integer = 1 To g
        '        strsql = "INSERT INTO " & Оплаты & "(" & insertinbaza & ",Рейс,ДатаОплаты,Сумма) VALUES(" & ID & "," & NomerRejsa & ",'" & Controls("MaskedTextBox" & i).Text & "','" & Controls("TextBox" & i).Text & "')"
        '        Updates(strsql)
        '    Next
        'Else
        Try
            strsql = "DELETE FROM " & Оплаты & " WHERE " & insertinbaza & "=" & ID & " And Рейс=" & NomerRejsa & ""
            Updates3(strsql)
        Catch ex As Exception

        End Try

        If g > 0 Then
            For i As Integer = 1 To g
                strsql = "INSERT INTO " & Оплаты & "(" & insertinbaza & ",Рейс,ДатаОплаты,Сумма)
VALUES(" & ID & "," & NomerRejsa & ",'" & Controls("MaskedTextBox" & i).Text & "','" & Controls("TextBox" & i).Text & "')"
                Updates3(strsql)
            Next
        End If
        'Next
        'End If



        MessageBox.Show("Данные приняты!", Рик)
        Очистка()
        Обновление()

    End Sub
    Private Sub Обновление()
        Dim dbl As Double = Nothing
        Dim stsql As String = "SELECT * FROM " & Оплаты & " WHERE " & insertinbaza & "=" & ID & " And Рейс=" & NomerRejsa & ""
        Dim ds As DataTable = Selects3(stsql)
        For i As Integer = 0 To ds.Rows.Count - 1
            If ds.Rows(i).Item(4).ToString.Contains(".") Then
                ds.Rows(i).Item(4) = Replace(ds.Rows(i).Item(4).ToString, ".", ",")
            End If
            dbl += CDbl(ds.Rows(i).Item(4).ToString)
        Next
        Label24.Text = Math.Round(CDbl(Отчет.datatbl.Rows(0).Item(2).ToString) - dbl, 2) & " " & Отчет.datatbl.Rows(0).Item(3).ToString
        Dim strsql1 As String

        If кто = "Перевозчик" Then
            strsql1 = "UPDATE РейсыПеревозчика SET ОстатокОплаты='" & CType(Math.Round(CDbl(Отчет.datatbl.Rows(0).Item(2).ToString) - dbl, 2), String) & "' WHERE Код=" & Отчет.datatbl.Rows(0).Item(0) & ""

        Else
            strsql1 = "UPDATE РейсыКлиента SET ОстатокОплаты='" & CType(Math.Round(CDbl(Отчет.datatbl.Rows(0).Item(2).ToString) - dbl, 2), String) & "' WHERE Код=" & Отчет.datatbl.Rows(0).Item(0) & ""
        End If
        Updates3(strsql1)
        Отчет.ОсткиОплат()

    End Sub
    Private Sub Очистка()
        For i As Integer = 1 To 8
            Select Case i
                Case 1
                    CheckBox1.Checked = False
                Case 2
                    CheckBox2.Checked = False
                Case 3
                    CheckBox3.Checked = False
                Case 4
                    CheckBox4.Checked = False
                Case 5
                    CheckBox5.Checked = False
                Case 6
                    CheckBox6.Checked = False
                Case 7
                    CheckBox7.Checked = False
                Case 8
                    CheckBox8.Checked = False
            End Select
        Next

        MaskedTextBox1.Text = ""
        MaskedTextBox2.Text = ""
        MaskedTextBox3.Text = ""
        MaskedTextBox4.Text = ""
        MaskedTextBox5.Text = ""
        MaskedTextBox6.Text = ""
        MaskedTextBox7.Text = ""
        MaskedTextBox8.Text = ""
        MaskedTextBox1.Enabled = False
        MaskedTextBox2.Enabled = False
        MaskedTextBox3.Enabled = False
        MaskedTextBox4.Enabled = False
        MaskedTextBox5.Enabled = False
        MaskedTextBox6.Enabled = False
        MaskedTextBox7.Enabled = False
        MaskedTextBox8.Enabled = False
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        TextBox6.Text = ""
        TextBox7.Text = ""
        TextBox8.Text = ""
        TextBox1.Enabled = False
        TextBox2.Enabled = False
        TextBox3.Enabled = False
        TextBox4.Enabled = False
        TextBox5.Enabled = False
        TextBox6.Enabled = False
        TextBox7.Enabled = False
        TextBox8.Enabled = False

        ostatok = Nothing


    End Sub

    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.MaskedTextBox1.Focus()
        End If
    End Sub

    Private Sub TextBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox2.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.MaskedTextBox2.Focus()
        End If
    End Sub

    Private Sub TextBox3_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox3.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.MaskedTextBox3.Focus()
        End If
    End Sub

    Private Sub TextBox4_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox4.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.MaskedTextBox4.Focus()
        End If
    End Sub

    Private Sub TextBox5_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox5.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.MaskedTextBox5.Focus()
        End If
    End Sub

    Private Sub TextBox6_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox6.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.MaskedTextBox6.Focus()
        End If
    End Sub

    Private Sub Оплата_FormClosed(sender As Object, e As FormClosedEventArgs) Handles MyBase.FormClosed
        Очистка()
    End Sub

    Private Sub TextBox7_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox7.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.MaskedTextBox7.Focus()
        End If
    End Sub

    Private Sub TextBox8_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox8.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.MaskedTextBox8.Focus()
        End If
    End Sub
End Class