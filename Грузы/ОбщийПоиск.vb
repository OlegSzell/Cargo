Imports System.ComponentModel

Public Class ОбщийПоиск
    Private Grid1Alsort As List(Of ФайлыExcelВсе)
    Private grdi1all As List(Of Grid1Class)
    Private bsgrdi1all As BindingSource
    Sub New()

        ' Этот вызов является обязательным для конструктора.
        InitializeComponent()

        ' Добавить код инициализации после вызова InitializeComponent().

    End Sub
    Sub New(ByVal _Grid1Alsort As List(Of ФайлыExcelВсе))
        Grid1Alsort = New List(Of ФайлыExcelВсе)
        Grid1Alsort.AddRange(_Grid1Alsort)


        ' Этот вызов является обязательным для конструктора.
        InitializeComponent()

        ' Добавить код инициализации после вызова InitializeComponent().

    End Sub
    Private Async Sub Предзагрузкаasync()
        Await Task.Run(Sub() Предзагрузка())
    End Sub
    Private Sub Предзагрузка()
        Dim mo As New AllUpd
        Do While AllClass.ФайлыExcelВсе Is Nothing
            mo.ФайлыExcelВсеAll()
        Loop
    End Sub
    Private Sub ОбщийПоиск_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Предзагрузкаasync()
        grdi1all = New List(Of Grid1Class)
        bsgrdi1all = New BindingSource
        bsgrdi1all.DataSource = grdi1all
        Grid1.DataSource = bsgrdi1all
        GridView(Grid1)
        Grid1.Columns(0).Width = 60
        Grid1.Columns(1).Width = 150
        Grid1.Columns(2).Width = 300
        Grid1.Columns(5).Width = 150
        Grid1.DefaultCellStyle.Font = New Font("Calibri", 10)

        Dim dlsr As New List(Of Grid1Class)

        If Grid1Alsort IsNot Nothing Then  'если форму запустили из ЖурналДобавитьГруз
            If grdi1all IsNot Nothing Then
                grdi1all.Clear()
            End If


            For Each b In Grid1Alsort
                Dim last As Integer = 50

                Dim _Груз = b.Груз

                Dim _Клиент As String
                If b.Клиент Is Nothing Then
                    _Клиент = String.Empty
                Else
                    If b.Клиент.Length > last Then
                        _Клиент = b.Клиент.Substring(0, last)
                    Else
                        _Клиент = b.Клиент
                    End If
                End If




                Dim _маршрут As String
                If b.Маршрут Is Nothing Then
                    _маршрут = String.Empty
                Else
                    If b.Маршрут.Length > last Then
                        _маршрут = b.Маршрут.Substring(0, last)
                    Else
                        _маршрут = b.Маршрут
                    End If
                End If

                Dim _ДатаЗагрузки = "- " & IIf(b.ДатаЗагрузки Is Nothing, String.Empty, b.ДатаЗагрузки)

                Dim _АдресЗагрузки As String
                If b.АдресЗагрузки Is Nothing Then
                    _АдресЗагрузки = String.Empty
                Else
                    If b.АдресЗагрузки.Length > last Then
                        _АдресЗагрузки = b.АдресЗагрузки.Substring(0, last)
                    Else
                        _АдресЗагрузки = b.АдресЗагрузки
                    End If
                End If

                Dim _АдресЗатаможки As String
                If b.АдресЗатаможки Is Nothing Then
                    _АдресЗатаможки = String.Empty
                Else
                    If b.АдресЗатаможки.Length > last Then
                        _АдресЗатаможки = b.АдресЗатаможки.Substring(0, last)
                    Else
                        _АдресЗатаможки = b.АдресЗатаможки
                    End If
                End If

                Dim _ДатаПодРастаможку = "- " & IIf(b.ДатаПодРастаможку Is Nothing, String.Empty, b.ДатаПодРастаможку)

                Dim _АдресРастаможки As String
                If b.АдресРастаможки Is Nothing Then
                    _АдресРастаможки = String.Empty
                Else
                    If b.АдресРастаможки.Length > last Then
                        _АдресРастаможки = b.АдресРастаможки.Substring(0, last)
                    Else
                        _АдресРастаможки = b.АдресРастаможки
                    End If
                End If

                Dim _АдресВыгрузки As String
                If b.АдресВыгрузки Is Nothing Then
                    _АдресВыгрузки = String.Empty
                Else
                    If b.АдресВыгрузки.Length > last Then
                        _АдресВыгрузки = b.АдресВыгрузки.Substring(0, last)
                    Else
                        _АдресВыгрузки = b.АдресВыгрузки
                    End If
                End If

                Dim _Перевозчик As String
                If b.Перевозчик Is Nothing Then
                    _Перевозчик = String.Empty
                Else
                    If b.Перевозчик.Length > last Then
                        _Перевозчик = b.Перевозчик.Substring(0, last)
                    Else
                        _Перевозчик = b.Перевозчик
                    End If
                End If

                Dim _Рейс = IIf(b.Рейс Is Nothing, String.Empty, b.Рейс)

                Dim f2 As New Grid1Class With {.Груз = _Груз, .Клиент = _Клиент,
                    .Маршрут = "Маршрут:" & vbCrLf & _маршрут & vbCrLf & "Дата загрузки:" & vbCrLf & _ДатаЗагрузки & vbCrLf & "Адрес загрузки:" & vbCrLf & _АдресЗагрузки & vbCrLf & "Адрес затаможки:" & vbCrLf & _АдресЗатаможки & vbCrLf & "___________" & vbCrLf &
                    "Дата подачи под растаможку:" & vbCrLf & _ДатаПодРастаможку & vbCrLf & "Адрес растаможки:" & vbCrLf & _АдресРастаможки & vbCrLf & "Адрес выгрузки:" & vbCrLf & _АдресВыгрузки, .Перевозчик = _Перевозчик, .Рейс = _Рейс}

                dlsr.Add(f2)

            Next

            grdi1all.AddRange(dlsr.OrderBy(Function(x) x.Клиент).ToList())
            Dim i1 As Integer = 1
            For Each b In grdi1all
                b.ID = i1
                i1 += 1
            Next

            bsgrdi1all.ResetBindings(False)

        End If
    End Sub

    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click

    End Sub
    Public Class Grid1Class
        Public Property ID As Integer
        Public Property Клиент As String
        Public Property Маршрут As String
        Public Property Груз As String
        Public Property Перевозчик As String
        Public Property Рейс As String
    End Class
End Class