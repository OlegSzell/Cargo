Option Explicit On
Imports System.Data.OleDb

Public Class Напоминание
    Dim ds As DataTable
    Private Sub Напоминание_Load(sender As Object, e As EventArgs) Handles MyBase.Load


        Dim strsql As String = "SELECT DISTINCT Клиент FROM ПереговорыКлиент ORDER BY Клиент"
        Dim ds As DataTable = Selects3(strsql)

        Me.ComboBox1.AutoCompleteCustomSource.Clear()
        Me.ComboBox1.Items.Clear()
        For Each r As DataRow In ds.Rows
            Me.ComboBox1.AutoCompleteCustomSource.Add(r.Item(0).ToString())
            Me.ComboBox1.Items.Add(r(0).ToString)
        Next
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        Try
            ds.Clear()
        Catch ex As Exception

        End Try
        Dim strsql As String = "SELECT ДатаПереговоров as [Дата переговоров], ТекстПереговора as [Текст переговора], ДатаНапоминания as [Дата напоминания],
ТекстНапоминания as [Текст напоминания], ОЧемДоговорВсплывФорма as [Последние переговоры], ДатаОчемДоговорилис as [Дата последних переговоров] 
FROM ПереговорыКлиент WHERE Клиент='" & ComboBox1.Text & "'"
        ds = Selects3(strsql)
        Grid1.DataSource = ds
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Try
            ds.Clear()
        Catch ex As Exception

        End Try

        Dim strsql As String = "SELECT Клиент as [Клиент], ДатаПереговоров as [Дата переговоров], ТекстПереговора as [Текст переговора], ДатаНапоминания as [Дата напоминания],
ТекстНапоминания as [Текст напоминания], ОЧемДоговорВсплывФорма as [Последние переговоры], ДатаОчемДоговорилис as [Дата последних переговоров] 
FROM ПереговорыКлиент WHERE Экспедитор='" & Экспедитор & "' ORDER BY Клиент"
        ds = Selects3(strsql)
        Grid1.DataSource = ds
    End Sub
End Class