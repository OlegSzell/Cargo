Option Explicit On
Imports System.Data.OleDb
Public Class ДобЧернСписок
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim rch1, rch2 As String
        rch1 = Trim(RichTextBox1.Text)
        rch2 = Trim(TextBox1.Text)

        If ЧерныйСписок.g = 1 Then
            Dim strsql2 As String = "UPDATE ЧерныйСписок SET Примечание = '" & Trim(TextBox1.Text) & "' WHERE Организация='" & RichTextBox1.Text & "'"
            Updates3(strsql2)
            MessageBox.Show("Данные обновлены!", Рик)
        Else
            Dim strsql As String = "INSERT INTO ЧерныйСписок(Организация,Примечание) VALUES('" & rch1 & "','" & rch2 & "')"
            Updates3(strsql)
            MessageBox.Show("Данные добавлены в базу!", Рик)
        End If

        RichTextBox1.Text = ""
        TextBox1.Text = ""
        ЧерныйСписок.refgrid1()
        ЧерныйСписок.g = 0
        Me.Close()

    End Sub

    Private Sub ДобЧернСписок_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
End Class