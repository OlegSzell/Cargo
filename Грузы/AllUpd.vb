Public Class AllUpd
    Inherits AllClass
    Public Async Sub ОплатыПерAllAsync()
        Await Task.Run(Sub() ОплатыПерAll())
    End Sub

    Public Sub ОплатыПерAll()
        Using db As New dbAllDataContext()
            Dim m = db.ОплатыПер.Select(Function(x) x).ToList()
            If m.Count > 0 Then
                ОплатыПер = New List(Of ОплатыПер)
                ОплатыПер = m
            Else
                Dim k As New List(Of ОплатыПер)
                ОплатыПер = k
            End If
        End Using
    End Sub
    Public Async Sub ОтчетРаботыСотрудникаAllAsync()
        Await Task.Run(Sub() ОтчетРаботыСотрудникаAll())
    End Sub

    Public Sub ОтчетРаботыСотрудникаAll()
        Using db As New dbAllDataContext()
            Dim m = db.ОтчетРаботыСотрудника.Select(Function(x) x).ToList()
            If m.Count > 0 Then
                ОтчетРаботыСотрудника = New List(Of ОтчетРаботыСотрудника)
                ОтчетРаботыСотрудника = m
            Else
                Dim k As New List(Of ОтчетРаботыСотрудника)
                ОтчетРаботыСотрудника = k
            End If
        End Using
    End Sub

    Public Async Sub ОтчетРаботыСотрудникаСводнаяAllAsync()
        Await Task.Run(Sub() ОтчетРаботыСотрудникаСводнаяAll())
    End Sub

    Public Sub ОтчетРаботыСотрудникаСводнаяAll()
        Using db As New dbAllDataContext()
            Dim m = db.ОтчетРаботыСотрудникаСводная.Select(Function(x) x).ToList()
            If m.Count > 0 Then
                ОтчетРаботыСотрудникаСводная = New List(Of ОтчетРаботыСотрудникаСводная)
                ОтчетРаботыСотрудникаСводная = m
            Else
                Dim k As New List(Of ОтчетРаботыСотрудникаСводная)
                ОтчетРаботыСотрудникаСводная = k
            End If
        End Using
    End Sub

    Public Async Sub ОплатыКлиентAllAsync()
        Await Task.Run(Sub() ОплатыКлиентAll())
    End Sub

    Public Sub ОплатыКлиентAll()
        Using db As New dbAllDataContext()
            Dim m = db.ОплатыКлиент.Select(Function(x) x).ToList()
            If m.Count > 0 Then
                ОплатыКлиент = New List(Of ОплатыКлиент)
                ОплатыКлиент = m
            Else
                Dim k As New List(Of ОплатыКлиент)
                ОплатыКлиент = k
            End If
        End Using
    End Sub

    Public Async Sub РейсыПеревозчикаAllAsync()
        Await Task.Run(Sub() РейсыПеревозчикаAll())
    End Sub

    Public Sub РейсыПеревозчикаAll()
        Using db As New dbAllDataContext()
            Dim m = db.РейсыПеревозчика.Select(Function(x) x).ToList()
            If m.Count > 0 Then
                РейсыПеревозчика = New List(Of РейсыПеревозчика)
                РейсыПеревозчика = m
            Else
                Dim k As New List(Of РейсыПеревозчика)
                РейсыПеревозчика = k
            End If
        End Using
    End Sub

    Public Async Sub ПеревозчикиAllAsync()
        Await Task.Run(Sub() ПеревозчикиAll())
    End Sub

    Public Sub ПеревозчикиAll()
        Using db As New dbAllDataContext()
            Dim m = db.Перевозчики.Select(Function(x) x).ToList()
            If m.Count > 0 Then
                Перевозчики = New List(Of Перевозчики)
                Перевозчики = m
            Else
                Dim k As New List(Of Перевозчики)
                Перевозчики = k
            End If
        End Using
    End Sub

    Public Async Sub ФормаСобствAllAsync()
        Await Task.Run(Sub() ФормаСобствAll())
    End Sub

    Public Sub ФормаСобствAll()
        Using db As New dbAllDataContext()
            Dim m = db.ФормаСобств.Select(Function(x) x).ToList()
            If m.Count > 0 Then
                ФормаСобств = New List(Of ФормаСобств)
                ФормаСобств = m
            Else
                Dim k As New List(Of ФормаСобств)
                ФормаСобств = k
            End If
        End Using
    End Sub

    Public Async Sub РейсыКлиентаAllAsync()
        Await Task.Run(Sub() РейсыКлиентаAll())
    End Sub

    Public Sub РейсыКлиентаAll()
        Using db As New dbAllDataContext()
            Dim m = db.РейсыКлиента.Select(Function(x) x).ToList()
            If m.Count > 0 Then
                РейсыКлиента = New List(Of РейсыКлиента)
                РейсыКлиента = m
            Else
                Dim k As New List(Of РейсыКлиента)
                РейсыКлиента = k
            End If
        End Using
    End Sub

    Public Async Sub КлиентAllAsync()
        Await Task.Run(Sub() КлиентAll())
    End Sub

    Public Sub КлиентAll()
        Using db As New dbAllDataContext()
            Dim m = db.Клиент.Select(Function(x) x).ToList()
            If m.Count > 0 Then
                Клиент = New List(Of Клиент)
                Клиент = m
            Else
                Dim k As New List(Of Клиент)
                Клиент = k
            End If
        End Using
    End Sub

    Public Async Sub ТипАвтоAllAsync()
        Await Task.Run(Sub() ТипАвтоAll())
    End Sub

    Public Sub ТипАвтоAll()
        Using db As New dbAllDataContext()
            Dim m = db.ТипАвто.Select(Function(x) x).ToList()
            If m.Count > 0 Then
                ТипАвто = New List(Of ТипАвто)
                ТипАвто = m
            Else
                Dim k As New List(Of ТипАвто)
                ТипАвто = k
            End If
        End Using
    End Sub

    Public Async Sub ГрузыКлиентовAllAsync()
        Await Task.Run(Sub() ГрузыКлиентовAll())
    End Sub

    Public Sub ГрузыКлиентовAll()
        Using db As New dbAllDataContext()
            Dim m = db.ГрузыКлиентов.Select(Function(x) x).ToList()
            If m.Count > 0 Then
                ГрузыКлиентов = New List(Of ГрузыКлиентов)
                ГрузыКлиентов = m
            Else
                Dim k As New List(Of ГрузыКлиентов)
                ГрузыКлиентов = k
            End If
        End Using
    End Sub

    Public Async Sub Календарь_ДатыAllAsync()
        Await Task.Run(Sub() Календарь_ДатыAll())
    End Sub

    Public Sub Календарь_ДатыAll()
        Using db As New dbAllDataContext()
            Dim m = db.Календарь_Даты.Select(Function(x) x).ToList()
            If m.Count > 0 Then
                Календарь_Даты = New List(Of Календарь_Даты)
                Календарь_Даты = m
            Else
                Dim k As New List(Of Календарь_Даты)
                Календарь_Даты = k
            End If
        End Using
    End Sub

    Public Async Sub ПереговорыКлиентAllAsync()
        Await Task.Run(Sub() ПереговорыКлиентAll())
    End Sub

    Public Sub ПереговорыКлиентAll()
        Using db As New dbAllDataContext()
            Dim m = db.ПереговорыКлиент.Select(Function(x) x).ToList()
            If m.Count > 0 Then
                ПереговорыКлиент = New List(Of ПереговорыКлиент)
                ПереговорыКлиент = m
            Else
                Dim k As New List(Of ПереговорыКлиент)
                ПереговорыКлиент = k
            End If
        End Using
    End Sub

    Public Async Sub КалендарьНапоминаниеAllAsync()
        Await Task.Run(Sub() КалендарьНапоминаниеAll())
    End Sub

    Public Sub КалендарьНапоминаниеAll()
        Using db As New dbAllDataContext()
            Dim m = db.КалендарьНапоминание.Select(Function(x) x).ToList()
            If m.Count > 0 Then
                КалендарьНапоминание = New List(Of КалендарьНапоминание)
                КалендарьНапоминание = m
            Else
                Dim k As New List(Of КалендарьНапоминание)
                КалендарьНапоминание = k
            End If
        End Using
    End Sub
    Public Async Sub ПаролиAllAsync()
        Await Task.Run(Sub() ПаролиAll())
    End Sub

    Public Sub ПаролиAll()
        Using db As New dbAllDataContext()
            Dim m = db.Пароли.Select(Function(x) x).ToList()
            If m.Count > 0 Then
                Пароли = New List(Of Пароли)
                Пароли = m
            Else
                Dim k As New List(Of Пароли)
                Пароли = k
            End If
        End Using
    End Sub
    Public Async Sub ЖурналДатаAllAsync()
        Await Task.Run(Sub() ЖурналДатаAll())
    End Sub

    Public Sub ЖурналДатаAll()
        Using db As New dbAllDataContext()
            Dim m = db.ЖурналДата.Select(Function(x) x).ToList()
            If m.Count > 0 Then
                ЖурналДата = New List(Of ЖурналДата)
                ЖурналДата = m
            Else
                Dim k As New List(Of ЖурналДата)
                ЖурналДата = k
            End If
        End Using
    End Sub
    Public Async Sub ЖурналКлиентГрузAllAsync()
        Await Task.Run(Sub() ЖурналКлиентГрузAll())
    End Sub

    Public Sub ЖурналКлиентГрузAll()
        Using db As New dbAllDataContext()
            Dim m = db.ЖурналКлиентГруз.Select(Function(x) x).ToList()
            If m.Count > 0 Then
                ЖурналКлиентГруз = New List(Of ЖурналКлиентГруз)
                ЖурналКлиентГруз = m
            Else
                Dim k As New List(Of ЖурналКлиентГруз)
                ЖурналКлиентГруз = k
            End If
        End Using
    End Sub
    Public Async Sub ЖурналКлиентМаршрутAllAsync()
        Await Task.Run(Sub() ЖурналКлиентМаршрутAll())
    End Sub

    Public Sub ЖурналКлиентМаршрутAll()
        Using db As New dbAllDataContext()
            Dim m = db.ЖурналКлиентМаршрут.Select(Function(x) x).ToList()
            If m.Count > 0 Then
                ЖурналКлиентМаршрут = New List(Of ЖурналКлиентМаршрут)
                ЖурналКлиентМаршрут = m
            Else
                Dim k As New List(Of ЖурналКлиентМаршрут)
                ЖурналКлиентМаршрут = k
            End If
        End Using
    End Sub
    Public Async Sub ЖурналКлиентСписокAllAsync()
        Await Task.Run(Sub() ЖурналКлиентСписокAll())
    End Sub

    Public Sub ЖурналКлиентСписокAll()
        Using db As New dbAllDataContext()
            Dim m = db.ЖурналКлиентСписок.Select(Function(x) x).ToList()
            If m.Count > 0 Then
                ЖурналКлиентСписок = New List(Of ЖурналКлиентСписок)
                ЖурналКлиентСписок = m
            Else
                Dim k As New List(Of ЖурналКлиентСписок)
                ЖурналКлиентСписок = k
            End If
        End Using
    End Sub

End Class
