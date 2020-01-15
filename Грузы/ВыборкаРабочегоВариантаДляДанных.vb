Option Explicit On
Imports System.Data.OleDb
Imports System.Threading
Public Class ВыборкаРабочегоВариантаДляДанных
    Dim fg As String
    Dim ds As DataTable
    Private Sub ВыборкаРабочегоВариантаДляДанных_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim df As String = Данные.idClient2
        'ds = Selects3(StrSql:="SELECT Код,Организация,ДляСкайпа FROM ГрузыКлиентов WHERE Организация='" & df & "'")
        ds = Selects3(StrSql:="SELECT Код,Организация,ДляСкайпа FROM ГрузыКлиентов WHERE Организация='" & df & "'")
        For i As Integer = 0 To ds.Rows.Count - 1
            If ds.Rows(i).Item(2).ToString = "" Then
                ds.Rows(i).Delete()
            End If
        Next

        Grid1.DataSource = ds


    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If fg.Length = 0 Then
            MessageBox.Show("Выберите текст для вставки", Рик)
        Else
            Данные.RichTextBox2.Text = fg
            MessageBox.Show("Данные вставлены", Рик)
            ds.Clear()
            Me.Close()
        End If
    End Sub

    Private Sub Grid1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles Grid1.CellClick
        fg = Grid1.CurrentRow.Cells(2).Value
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        ds.Clear()
        Me.Close()
    End Sub
End Class