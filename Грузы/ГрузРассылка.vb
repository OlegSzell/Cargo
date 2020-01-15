Option Explicit On
Imports System.Data.OleDb
Public Class ГрузРассылка
    Private Sub ГрузРассылка_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim strsql As String = "SELECT ДляСкайпа FROM ГрузыКлиентов WHERE Код=" & IDГруз & ""
        'Dim ds As DataTable = Selects3(strsql)
        Dim ds As DataTable = Selects3(strsql)
        RichTextBox1.Text = ds.Rows(0).Item(0).ToString

    End Sub
End Class