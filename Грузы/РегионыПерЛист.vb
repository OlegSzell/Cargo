Option Explicit On
Imports System.Data.OleDb
Public Class РегионыПерЛист
    Dim strsql As String
    Dim Рик As String = "ООО Рикманс"
    Private Sub РегионыПерЛист_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Dim ds As DataTable = Selects3(StrSql:="SELECT Регионы FROM РегионыРоссии ORDER BY Регионы")
        ListBox1.Items.Clear()

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

        Dim country As String
        Dim i As Integer
        For i = 0 To ListBox1.SelectedItems.Count - 1
            country = ListBox1.SelectedItems(i).ToString & ", " & country
        Next


        Updates3(stroka:="UPDATE ПеревозчикиБаза SET [Регионы]='" & country & "' WHERE ПеревозчикиБаза.ID = " & idtabl & "")

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Interbaza()
        ИзменПеревоз.CheckBox2.Checked = False
        ПоискПеревозчиков.ОбнКом1()
        ИзменПеревоз.дан()
        Me.Close()
    End Sub

    Private Sub РегионыПерЛист_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        ИзменПеревоз.CheckBox2.Checked = False
    End Sub
End Class