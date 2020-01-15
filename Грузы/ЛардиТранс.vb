Imports HtmlAgilityPack
Imports System.Xml
Imports mshtml
Imports System.Text.RegularExpressions
Imports System.Net
Imports System.IO
Imports System.Text
Imports System.Runtime.InteropServices

Public Class ЛардиТранс
    Private Sub ЛардиТранс_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        'Dim ds2 As DataSet 'не работает
        'ds2 = RegExpParser.RegExpParser.ConvertHTMLTablesToDataSet("https://lardi-trans.com/gruz/?foi=&filter_marker=new&countryfrom=PL&countryto=BY&cityFrom=&cityIdFrom=0&cityTo=&cityIdTo=&dateFrom=&dateTo=&bt_chbs_slc=&strictBodyTypes=on&mass=&mass2=&value=&value2=&gabDl=&gabSh=&gabV=&zagruzFilterId=&adr=-1&showType=all&npredlFrom=all&sortByCountriesFirst=on&startSearch=%D0%A1%D0%B4%D0%B5%D0%BB%D0%B0%D1%82%D1%8C+%D0%B2%D1%8B%D0%B1%D0%BE%D1%80%D0%BA%D1%83&notFirstLoad=done")
        'Grid1.DataSource = ds2.Tables(0)
        'Exit Sub


        'Dim Web As New HtmlWeb' работает но плохо, только одно значение находит
        'Dim Doc1 As New HtmlAgilityPack.HtmlDocument
        'Doc1 = Web.Load("https://lardi-trans.com/gruz/?foi=&filter_marker=new&countryfrom=PL&countryto=BY&cityFrom=&cityIdFrom=0&cityTo=&cityIdTo=&dateFrom=&dateTo=&bt_chbs_slc=&strictBodyTypes=on&mass=&mass2=&value=&value2=&gabDl=&gabSh=&gabV=&zagruzFilterId=&adr=-1&showType=all&npredlFrom=all&sortByCountriesFirst=on&startSearch=%D0%A1%D0%B4%D0%B5%D0%BB%D0%B0%D1%82%D1%8C+%D0%B2%D1%8B%D0%B1%D0%BE%D1%80%D0%BA%D1%83&notFirstLoad=done")
        'For Each table As HtmlNode In Doc1.DocumentNode.SelectNodes("//*[@id='pricing']/tr/td")
        '    MsgBox(table.InnerText)
        'Next

        'Exit Sub


        Dim link As String = "https://lardi-trans.com/gruz/?foi=&filter_marker=new&countryfrom=PL&countryto=BY&cityFrom=&cityIdFrom=0&cityTo=&cityIdTo=&dateFrom=&dateTo=&bt_chbs_slc=&strictBodyTypes=on&mass=&mass2=&value=&value2=&gabDl=&gabSh=&gabV=&zagruzFilterId=&adr=-1&showType=all&npredlFrom=all&sortByCountriesFirst=on&startSearch=%D0%A1%D0%B4%D0%B5%D0%BB%D0%B0%D1%82%D1%8C+%D0%B2%D1%8B%D0%B1%D0%BE%D1%80%D0%BA%D1%83&notFirstLoad=done"
        'download page from the link into an HtmlDocument'
        Dim doc As HtmlAgilityPack.HtmlDocument = New HtmlWeb().Load(link)

        'select <div> having class attribute equals fontdef1'
        Dim div As HtmlNode = doc.DocumentNode.SelectSingleNode("//div[@class='predl-search-res']") '("//div[@class='fontdef1']")
        'if the div is found, print the inner text'

        For Each r As HtmlNode In div.ChildNodes
            'MsgBox(r.InnerText)
        Next



        'Dim ds As New DataTable
        'Dim col = New DataColumn("один")
        'ds.Columns.Add(col)
        'ds.Rows.Add(div.InnerText)
        'Grid1.DataSource = ds

        'RichTextBox1.Text = div.InnerText

        If Not div Is Nothing Then
            Console.WriteLine(div.InnerText.Trim())
        End If
    End Sub

    Public Function GetTables(ByVal strHTML) As ArrayList
        Dim htmlDocument As IHTMLDocument2
        'htmlDocument = New HTMLDocumentClass()

        Dim s As New ArrayList

        htmlDocument.write(strHTML)
        htmlDocument.close()

        Dim allTables As IHTMLElementCollection = htmlDocument.all.tags("table")
        Dim iTable As IHTMLElement

        For Each iTable In allTables
            s.Add(iTable.outerHTML)
        Next
        Return s

    End Function




End Class

Namespace RegExpParser ' не работает на ларди транс
    Public Class RegExpParser
        Public Shared Function ConvertHTMLTablesToDataSet(ByVal HTML As String) As DataSet
            ' Declarations  
            Dim ds As New DataSet
            Dim dt As DataTable
            Dim dr As DataRow
            'Dim dc As DataColumn

            Dim TableExpression As String = "<table[^>]*>(.*?)</table>"
            Dim HeaderExpression As String = "<th[^>]*>(.*?)</th>"
            Dim RowExpression As String = "<tr[^>]*>(.*?)</tr>"
            Dim ColumnExpression As String = "<td[^>]*>(.*?)</td>"
            Dim HeadersExist As Boolean = False
            Dim iCurrentColumn As Integer = 0
            Dim iCurrentRow As Integer = 0

            ' Get a match for all the tables in the HTML  
            Dim Tables As MatchCollection = Regex.Matches(HTML, TableExpression, RegexOptions.Multiline Or RegexOptions.Singleline Or RegexOptions.IgnoreCase)

            ' Loop through each table element  
            On Error Resume Next
            For Each Table As Match In Tables

                ' Reset the current row counter and the header flag  
                iCurrentRow = 0
                HeadersExist = False

                ' Add a new table to the DataSet  
                dt = New DataTable

                ' Create the relevant amount of columns for this table (use the headers if they exist, otherwise use default names)  
                If Table.Value.Contains("<th") Then
                    ' Set the HeadersExist flag  
                    HeadersExist = True

                    ' Get a match for all the rows in the table  
                    Dim Headers As MatchCollection = Regex.Matches(Table.Value, HeaderExpression, RegexOptions.Multiline Or RegexOptions.Singleline Or RegexOptions.IgnoreCase)

                    ' Loop through each header element  
                    For Each Header As Match In Headers
                        dt.Columns.Add(Header.Groups(1).ToString)
                    Next
                Else
                    For iColumns As Integer = 1 To Regex.Matches(Regex.Matches(Regex.Matches(Table.Value, TableExpression, RegexOptions.Multiline Or RegexOptions.Singleline Or RegexOptions.IgnoreCase).Item(0).ToString, RowExpression, RegexOptions.Multiline Or RegexOptions.Singleline Or RegexOptions.IgnoreCase).Item(0).ToString, ColumnExpression, RegexOptions.Multiline Or RegexOptions.Singleline Or RegexOptions.IgnoreCase).Count
                        dt.Columns.Add("Column " & iColumns)
                    Next
                End If

                ' Get a match for all the rows in the table  
                Dim Rows As MatchCollection = Regex.Matches(Table.Value, RowExpression, RegexOptions.Multiline Or RegexOptions.Singleline Or RegexOptions.IgnoreCase)
                ' Loop through each row element  
                For Each Row As Match In Rows
                    ' Only loop through the row if it isn't a header row  
                    If Not (iCurrentRow = 0 And HeadersExist = True) Then
                        ' Create a new row and reset the current column counter  
                        dr = dt.NewRow
                        iCurrentColumn = 0
                        ' Get a match for all the columns in the row  
                        Dim Columns As MatchCollection = Regex.Matches(Row.Value, ColumnExpression, RegexOptions.Multiline Or RegexOptions.Singleline Or RegexOptions.IgnoreCase)
                        ' Loop through each column element  
                        For Each Column As Match In Columns
                            ' Add the value to the DataRow  
                            dr(iCurrentColumn) = Column.Groups(1).ToString

                            ' Increase the current column   
                            iCurrentColumn += 1
                        Next
                        ' Add the DataRow to the DataTable  
                        dt.Rows.Add(dr)
                    End If
                    ' Increase the current row counter  
                    iCurrentRow += 1
                Next
                ' Add the DataTable to the DataSet  
                ds.Tables.Add(dt)
            Next
            Return (ds)

        End Function
    End Class
End Namespace