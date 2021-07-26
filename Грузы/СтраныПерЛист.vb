Option Explicit On
Imports System.Data.OleDb

Public Class СтраныПерЛист
    Dim strsql As String
    Dim Рик As String = "ООО Рикманс"

    Private Sub СтраныПерЛист_Load(sender As Object, e As EventArgs) Handles MyBase.Load


        Dim ds As DataTable = Selects3(StrSql:="SELECT Страна FROM Страна ORDER BY Страна")
        For Each r As DataRow In ds.Rows
            ListBox1.Items.Add(r(0).ToString)
        Next

    End Sub
    Private Sub Interbaza()
        If MessageBox.Show("Изменить данные?", Рик, MessageBoxButtons.OKCancel) = DialogResult.Cancel Then

            Exit Sub
        End If
        If ListBox1.SelectedIndex = -1 Then

            Exit Sub
        End If
        Dim country As String = String.Empty
        Dim i As Integer
        For i = 0 To ListBox1.SelectedItems.Count - 1
            country = ListBox1.SelectedItems(i).ToString & ", " & country
        Next


        Updates3(stroka:="UPDATE ПеревозчикиБаза SET [Страны перевозок]='" & country & "' WHERE ПеревозчикиБаза.ID = " & idtabl & "")



    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Interbaza()
        ИзменПеревоз.CheckBox1.Checked = False
        If pr2 = 0 Then
            ПоискПеревозчиков.ОбнКом1()
        End If
        ИзменПеревоз.дан()
        Me.Close()
    End Sub

    Private Sub СтраныПерЛист_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        ИзменПеревоз.CheckBox1.Checked = False
    End Sub
End Class