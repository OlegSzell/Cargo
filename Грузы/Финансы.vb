Public Class Финансы
    Public grid1all As New List(Of AppEFC.Finans)
    Public bsgrid1 As New BindingSource
    Private Sub Label9_Click(sender As Object, e As EventArgs) Handles Label9.Click

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Using db As New AppEFC
            Dim f As New AppEFC.Finans
            f.Dates = Now
            f.Клиент = "Виталюр"
            f.Перевозчик = "РТЛТранс"
            f.ОплаченоКлиентом = "1500"
            f.ОплаченоПеревозчику = "1300"
            f.Остаток = CType(f.ОплаченоКлиентом, Integer) - CType(f.ОплаченоПеревозчику, Integer)
            f.СтавкаПеревозчика = "1300"
            f.СуммаПеревозки = "1500"

            db.Finanss.Add(f)
            db.SaveChanges()

            Dim f1 = db.Finanss.ToList()
            grid1all.AddRange(f1)
            bsgrid1.ResetBindings(False)
        End Using
    End Sub

    Private Sub Финансы_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        bsgrid1.DataSource = grid1all
        Grid1.DataSource = bsgrid1
        GridView(Grid1)

    End Sub
End Class