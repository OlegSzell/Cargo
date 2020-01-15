Imports System.ComponentModel

Public Class ПереговорыРедакция
    Private Sub ПереговорыРедакция_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ЗагрСис()
    End Sub
    Private Sub ЗагрСис()
        Dim id As Integer = ПереговорыВсе.id
        Dim strsql As String = "SELECT * FROM ПереговорыКлиент  WHERE Код=" & id & ""
        Dim ds As DataTable = Selects3(strsql)
        RichTextBox1.Text = ds.Rows(0).Item(1).ToString
        RichTextBox2.Text = ds.Rows(0).Item(3).ToString
        RichTextBox3.Text = ds.Rows(0).Item(5).ToString
        RichTextBox4.Text = ds.Rows(0).Item(6).ToString
        RichTextBox5.Text = ds.Rows(0).Item(8).ToString
        Label10.Text = ds.Rows(0).Item(7).ToString
        Label12.Text = ds.Rows(0).Item(0).ToString
        MaskedTextBox1.Text = ds.Rows(0).Item(2).ToString
        MaskedTextBox2.Text = ds.Rows(0).Item(4).ToString
        MaskedTextBox3.Text = ds.Rows(0).Item(9).ToString

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

        If MaskedTextBox1.MaskCompleted = False Then
            MessageBox.Show("Заполните дату переговоров!", Рик)
            Exit Sub
        End If

        Dim id As Integer = CType(Label12.Text, Integer)
        Dim strsql As String

        If MaskedTextBox2.MaskCompleted = False And MaskedTextBox3.MaskCompleted = False Then
            strsql = "UPDATE ПереговорыКлиент SET ТекстПереговора='" & RichTextBox2.Text & "', ТекстНапоминания='" & RichTextBox3.Text & "',
КонтДанные='" & RichTextBox4.Text & "', ОЧемДоговорВсплывФорма='" & RichTextBox5.Text & "'
WHERE Код=" & id & ""

        ElseIf MaskedTextBox2.MaskCompleted = True And MaskedTextBox3.MaskCompleted = False Then

            strsql = "UPDATE ПереговорыКлиент SET ТекстПереговора='" & RichTextBox2.Text & "', ТекстНапоминания='" & RichTextBox3.Text & "',
КонтДанные='" & RichTextBox4.Text & "', ОЧемДоговорВсплывФорма='" & RichTextBox5.Text & "', ДатаНапоминания='" & MaskedTextBox2.Text & "'
WHERE Код=" & id & ""

        ElseIf MaskedTextBox2.MaskCompleted = False And MaskedTextBox3.MaskCompleted = True Then

            strsql = "UPDATE ПереговорыКлиент SET ТекстПереговора='" & RichTextBox2.Text & "', ТекстНапоминания='" & RichTextBox3.Text & "',
КонтДанные='" & RichTextBox4.Text & "', ОЧемДоговорВсплывФорма='" & RichTextBox5.Text & "', ДатаОчемДоговорилис='" & MaskedTextBox3.Text & "'
WHERE Код=" & id & ""

        Else
            strsql = "UPDATE ПереговорыКлиент SET ТекстПереговора='" & RichTextBox2.Text & "', ТекстНапоминания='" & RichTextBox3.Text & "',
КонтДанные='" & RichTextBox4.Text & "', ОЧемДоговорВсплывФорма='" & RichTextBox5.Text & "', ДатаНапоминания='" & MaskedTextBox2.Text & "',
ДатаОчемДоговорилис='" & MaskedTextBox3.Text & "'
WHERE Код=" & id & ""
        End If
        Updates3(strsql)
        MessageBox.Show("Данные обновлены!", Рик)
    End Sub

    Private Sub ПереговорыРедакция_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        RichTextBox1.Text = ""
        RichTextBox2.Text = ""
        RichTextBox3.Text = ""
        RichTextBox4.Text = ""
        RichTextBox5.Text = ""
        Label10.Text = ""
        Label12.Text = ""
        MaskedTextBox1.Text = ""
        MaskedTextBox2.Text = ""
        MaskedTextBox3.Text = ""
        ПереговорыВсе.комб1()

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim strsql As String

        If MaskedTextBox1.MaskCompleted = False Then
            MessageBox.Show("Заполните дату переговоров!", Рик)
            Exit Sub
        End If

        If MaskedTextBox2.MaskCompleted = False And MaskedTextBox3.MaskCompleted = False Then
            strsql = "INSERT INTO ПереговорыКлиент(Клиент,ДатаПереговоров,ТекстПереговора,ТекстНапоминания,КонтДанные,Экспедитор,ОЧемДоговорВсплывФорма)
VALUES('" & RichTextBox1.Text & "','" & MaskedTextBox1.Text & "','" & RichTextBox2.Text & "','" & RichTextBox3.Text & "',
'" & RichTextBox4.Text & "','" & Label10.Text & "','" & RichTextBox5.Text & "')"

        ElseIf MaskedTextBox2.MaskCompleted = True And MaskedTextBox3.MaskCompleted = False Then

            strsql = "INSERT INTO ПереговорыКлиент(Клиент,ДатаПереговоров,ТекстПереговора,ТекстНапоминания,КонтДанные,Экспедитор,ОЧемДоговорВсплывФорма,ДатаНапоминания)
VALUES('" & RichTextBox1.Text & "','" & MaskedTextBox1.Text & "','" & RichTextBox2.Text & "','" & RichTextBox3.Text & "',
'" & RichTextBox4.Text & "','" & Label10.Text & "','" & RichTextBox5.Text & "','" & MaskedTextBox2.Text & "')"

        ElseIf MaskedTextBox2.MaskCompleted = False And MaskedTextBox3.MaskCompleted = True Then

            strsql = "INSERT INTO ПереговорыКлиент(Клиент,ДатаПереговоров,ТекстПереговора,ТекстНапоминания,КонтДанные,Экспедитор,ОЧемДоговорВсплывФорма,ДатаОчемДоговорилис)
VALUES('" & RichTextBox1.Text & "','" & MaskedTextBox1.Text & "','" & RichTextBox2.Text & "','" & RichTextBox3.Text & "',
'" & RichTextBox4.Text & "','" & Label10.Text & "','" & RichTextBox5.Text & "','" & MaskedTextBox3.Text & "')"

        Else
            strsql = "INSERT INTO ПереговорыКлиент(Клиент,ДатаПереговоров,ТекстПереговора,ТекстНапоминания,КонтДанные,Экспедитор,ОЧемДоговорВсплывФорма,ДатаНапоминания,ДатаОчемДоговорилис)
VALUES('" & RichTextBox1.Text & "','" & MaskedTextBox1.Text & "','" & RichTextBox2.Text & "','" & RichTextBox3.Text & "',
'" & RichTextBox4.Text & "','" & Label10.Text & "','" & RichTextBox5.Text & "','" & MaskedTextBox2.Text & "','" & MaskedTextBox3.Text & "')"
        End If
        Updates3(strsql)
        MessageBox.Show("Переговоры добавлены!", Рик)


    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked = True Then
            MaskedTextBox1.Enabled = True
        Else
            MaskedTextBox1.Enabled = False
        End If
    End Sub
End Class