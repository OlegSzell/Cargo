Public Class ЖурналClass

End Class
Public Class ГрузыClass
    Public Property Номер As Integer
    Public Property Наименование As String
    Public Property Вес As Integer
    Public Property Обьем As String
    Public Property Длина As String
    Public Property Ширина As String
    Public Property Высота As String
    Public Property ТипПогрузки As String
    Public Property ПаллетыШтук As Integer
    Public Property РазмерПаллет As String
    Public Property ADR As String
    Public Property ТипАвто As String
    Public Property ДополнитИнформация As String
End Class
Public Class МаршрутыClass
    Public Property Номер As Integer
    Public Property СтранаПогрузки As String
    Public Property СтранаВыгрузки As String
    Public Property ГородПогрузки As String
    Public Property ГородВыгрузки As String
    Public Property КвадратПогрузки As String
    Public Property КвадратВыгрузки As String
    Public Property ТаможняОтправления As String
    Public Property ТаможняНазначения As String
    Public Property Ставка As String
    Public Property EX As String
    Public Property ДополнитИнформация As String
End Class
Public Class AllClientClass
    Inherits МаршрутыClass
    Public Property Наименование As String
    Public Property Вес As Integer
    Public Property Обьем As String
    Public Property Длина As String
    Public Property Ширина As String
    Public Property Высота As String
    Public Property ТипПогрузки As String
    Public Property ПаллетыШтук As Integer
    Public Property РазмерПаллет As String
    Public Property ADR As String
    Public Property ТипАвто As String
    Public Property ДополнитИнформация1 As String

End Class
Public Class Grid2ЖурналClass
    Public Property Номер As Integer
    Public Property Клиент As String
    Public Property Дата As Date
    Public Property ДатаЗагрузки As Date
    Public Property Mаршрут As String
    Public Property МаршрутList As List(Of String)
End Class
