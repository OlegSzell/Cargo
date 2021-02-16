Public Class NbRbClassNew
    Sub New()

    End Sub

    Public Function Курс(ByVal Da As Date, ByVal Валюта As String) As String

        Dim NumCod As String
        Select Case Валюта
            Case "Евро"
                NumCod = "978"
            Case "Доллар"
                NumCod = "840"
            Case Else
                NumCod = "643"
        End Select



        Dim obser As New List(Of NbRBClass)

        Dim dat As String = Format(Da, "MM/dd/yyyy")
        Dim str As String = "https://www.nbrb.by/services/xmlexrates.aspx?ondate=" & dat
        Dim xdoc As XDocument
        Try
            xdoc = XDocument.Load(str)
        Catch ex As Exception
            MessageBox.Show("Пробелма с данными из банка НБРБ!", Рик)
            Return 0
        End Try

        For Each x As XElement In xdoc.Element("DailyExRates").Elements("Currency")
            Dim md = x.Attribute("Id").Value

            Dim nb As New NbRBClass With {.Cur_ID = x.Attribute("Id").Value, .NumCode = x.Element("NumCode").Value, .CharCode = x.Element("CharCode").Value, .Scale = x.Element("Scale").Value,
                .Name = x.Element("Name").Value, .Rate = x.Element("Rate").Value}

            obser.Add(nb)

            'Dim attr As XAttribute = x.Attribute("Id")
            'Dim NumCode As XElement = x.Element("NumCode")
            'Dim CharCode As XElement = x.Element("CharCode")
            'Dim Scale As XElement = x.Element("Scale")
            'Dim Name As XElement = x.Element("Name")
            'Dim Rate As XElement = x.Element("Rate")

            'Dim list As New List(Of String) From {attr, NumCode, CharCode, Scale, Name, Rate}



        Next

        Return obser.Where(Function(x) x.NumCode = NumCod).Select(Function(x) x.Rate).FirstOrDefault()




    End Function


End Class
Public Class NbRBClass
    Public Property Cur_ID As Integer
    Public Property NumCode As String
    Public Property CharCode As String
    Public Property Scale As String
    Public Property Name As String
    Public Property Rate As String




End Class