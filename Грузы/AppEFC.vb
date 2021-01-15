Imports Microsoft.EntityFrameworkCore
Public Class AppEFC
    Inherits DbContext
    Public Property Finanss As DbSet(Of Finans)
    Public Sub New()
        'Database.EnsureCreated()
    End Sub

    Protected Overrides Sub OnConfiguring(ByVal optionsBuilder As DbContextOptionsBuilder)
        Try
            optionsBuilder.UseSqlServer("Data Source=178.124.211.175,52891;Initial Catalog=Finanses;Persist Security Info=True;User ID=Fin;Password=1w1FuXgNHyyM")
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub

    Public Class Finans
        Public Property Id As Integer
        Public Property Dates As Date
        Public Property Клиент As String
        Public Property СуммаПеревозки As String
        Public Property ОплаченоКлиентом As String
        Public Property Перевозчик As String
        Public Property СтавкаПеревозчика As String
        Public Property ОплаченоПеревозчику As String
        Public Property Остаток As String
        Public Property Примечание As String


    End Class

End Class
