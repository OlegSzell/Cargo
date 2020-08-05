Imports System.IO

Public Class ПутьЗапускаПрограммы
    Public Property _ПутьЗапускаПрограммы As String
    Sub New()

    End Sub
    Sub New(ByVal bool As Boolean)
        If bool = True Then

            GetDictionary(True)
        Else
            GetDictionary(False)
        End If

    End Sub
    Public Function ВремянкаДляРесурса() As String
        'папка для документов перемещаем из ресурса во временную папку
        Dim k2 = IO.Path.Combine(My.Application.Info.DirectoryPath, "Времянка2")
        Dim po2 = IO.Directory.CreateDirectory(k2)
        If po2.Exists = False Then
            po2.Create()
        End If
        Return k2
    End Function

    Private Sub GetDictionary(ByVal bool As Boolean)
        If bool = False Then
            If IO.Directory.Exists(IO.Path.Combine(My.Application.Info.DirectoryPath, "Времянка")) = False Then
                'Dim m = IO.Path.GetTempPath(Environment.CurrentDirectory)
                'Dim k1 = IO.Path.Combine(Assembly.GetExecutingAssembly().Location, "Времянка")
                Dim k = IO.Path.Combine(My.Application.Info.DirectoryPath, "Времянка")
                Dim po = IO.Directory.CreateDirectory(k)
                If po.Exists = False Then
                    po.Create()
                End If
                _ПутьЗапускаПрограммы = k





            Else

                _ПутьЗапускаПрограммы = IO.Path.Combine(My.Application.Info.DirectoryPath, "Времянка")

            End If
        Else
            'очищаем папку при запуске программы
            If IO.Directory.Exists(IO.Path.Combine(My.Application.Info.DirectoryPath, "Времянка")) = False Then
                'Dim m = IO.Path.GetTempPath(Environment.CurrentDirectory)
                'Dim k1 = IO.Path.Combine(Assembly.GetExecutingAssembly().Location, "Времянка")
                Dim k = IO.Path.Combine(My.Application.Info.DirectoryPath, "Времянка")
                Dim po = IO.Directory.CreateDirectory(k)
                If po.Exists = False Then
                    po.Create()
                    _ПутьЗапускаПрограммы = k
                End If

            Else
                Try
                    Dim po As New DirectoryInfo(IO.Path.Combine(My.Application.Info.DirectoryPath, "Времянка"))
                    po.Delete(True)

                    Dim po1 = IO.Directory.CreateDirectory(IO.Path.Combine(My.Application.Info.DirectoryPath, "Времянка"))
                    If po1.Exists = False Then
                        po1.Create()
                    End If
                Catch ex As Exception

                Finally
                    _ПутьЗапускаПрограммы = IO.Path.Combine(My.Application.Info.DirectoryPath, "Времянка")

                End Try
            End If






        End If

    End Sub

End Class

