Option Explicit On
Imports System.Data.OleDb
Public Class РабочаяФормаСостояние
    Private Sub РабочаяФормаСостояние_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Label4.Visible = False
        Label5.Visible = False
        TextBox3.Visible = False
        TextBox4.Visible = False
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        If ComboBox1.Text = "Перенесен" Then
            Label4.Visible = True
            Label5.Visible = True
            TextBox3.Visible = True
        Else
            Label4.Visible = False
            Label5.Visible = False
            TextBox3.Visible = False
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim all As String

        If TextBox3.Visible = True Then
            all = ComboBox1.Text & " на " & TextBox3.Text & " дней"
        ElseIf TextBox3.Visible = True And TextBox4.Visible = False Then
            all = ComboBox1.Text
        ElseIf TextBox4.Visible = True Then
            all = TextBox4.Text
        Else
            all = ComboBox1.Text
        End If

        Updates3(stroka:="UPDATE ГрузыКлиентов SET Состояние='" & all & "',Груз='" & TextBox2.Text & "' WHERE Код=" & IDГруз & "")
        Me.Close()
        Выборка.Refreshgrid()


    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked = True Then
            TextBox4.Visible = True
        Else
            TextBox4.Visible = False
        End If
    End Sub

    Private Sub РабочаяФормаСостояние_Closed(sender As Object, e As EventArgs) Handles Me.Closed
        CheckBox1.Checked = False
        TextBox4.Text = ""
        TextBox1.Text = ""
        TextBox2.Text = ""
    End Sub
End Class