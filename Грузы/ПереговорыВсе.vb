Option Explicit On
Imports System.Data.OleDb
Imports HtmlAgilityPack

Public Class ПереговорыВсе
    Public id As Integer
    Private Sub ПереговорыВсе_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        комб1()
    End Sub
    Public Sub комб1()
        Me.MdiParent = MDIParent1


        Dim strsql As String = "SELECT Код, Клиент, ДатаПереговоров as [Дата перегов],
ТекстПереговора as [Текст], ДатаНапоминания as [Дата напом],
ПереговорыКлиент.ТекстНапоминания as [Текст Напоминания], ПереговорыКлиент.КонтДанные as [Контактные данные],
ПереговорыКлиент.Экспедитор as [Экспед], ПереговорыКлиент.ОЧемДоговорВсплывФорма as [О чем договорились],
ПереговорыКлиент.ДатаОчемДоговорилис as [Дата о чем договорились]
FROM ПереговорыКлиент
WHERE ТекстПереговора Is Not Null ORDER BY Клиент"
        'WHERE ТекстПереговора Is Not Null ORDER BY ДатаПереговоров DESC"
        Dim ds As DataTable = Selects3(strsql)
        Grid1.DataSource = ds
        Grid1.Columns(3).Width = 250
        Grid1.Columns(2).Width = 70
        Grid1.Columns(4).Width = 70
        Grid1.Columns(9).Width = 70
        Grid1.Columns(7).Width = 70
        Grid1.Columns(5).Width = 250
        Grid1.Columns(8).Width = 200
        Grid1.Columns(0).Visible = False

        'For Each r As DataGridViewRow In Grid1.Rows
        '    If r.Cells(2).Value = Now.ToShortDateString Then
        '        r.DefaultCellStyle.BackColor = Color.YellowGreen
        '    End If
        'Next


        GridView(Grid1)


    End Sub

    Private Sub Grid1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles Grid1.CellContentClick

    End Sub


    Private Sub Grid1_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles Grid1.CellDoubleClick
        id = Nothing
        id = Grid1.CurrentRow.Cells(0).Value
        ПереговорыРедакция.ShowDialog()
    End Sub


End Class