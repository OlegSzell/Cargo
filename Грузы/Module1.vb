Imports System.Data.OleDb
Imports System.Threading
Imports System.Data.SqlClient
Module Module1
    Public conn As OleDbConnection
    Public conn3 As SqlConnection
    Public conn1 As OleDbConnection
    Public Отдел, Nm As String
    Public idtabl, errds, pro, pr2 As Integer
    Public IDГруз, Отмена As Integer
    Public IDПеревоза, ОтКогоИзмен As Integer
    Public Экспедитор, NameПеревоза As String
    'Public ConString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=U:\Офис\Рикманс\ДанныеРикманс.accdb; Persist Security Info=False;"
    Public ConString1 As String = "Data Source=45.14.50.13\723\SQLEXPRESS,1433;Network Library=DBMSSOCN;Initial Catalog=Rickmans;User ID=userOleg;Password=Zf6VpP37Ol"
    ' Public ConString As String = "Data Source=45.14.50.142\2749\SQLEXPRESS,1433;Network Library=DBMSSOCN;Initial Catalog=Rickmans;User ID=userOleg1;Password=Zf6VpP37Ol"
    Public ConString As String = "Data Source=178.124.211.175,52891;Initial Catalog=Rickmans;Persist Security Info=True;User ID=Rickmans;Password=Zf6VpP37Ol"
    Public ДобПер, ОбнПер As Integer
    Public Рик As String = "ООО Рикманс"
    Public Files3 As List(Of ПутиДоков)
    Public FilesПолнПуть() As String
    Public НомерРейса3 As String
    Public НомРейса As Integer
    Public bl As Boolean
    Public fall As New Thread(AddressOf Сводная_по_рейсам.ALL)
    Public Delegate Sub coxt(ByVal form As Form, ByVal strsql As String, ByVal c As ComboBox)
    Public Delegate Sub coxt1(ByVal form As Form, ByVal strsql As String, ByVal c As ListBox, ByVal com As ComboBox)
    Public Delegate Sub coxt2()
    Public Delegate Sub coxt3()
    Public Delegate Sub coxt4()
    Public Delegate Sub coxt5()
    Public ДанныеДляВставкиСкайпа As Tuple(Of String, String, String, String, Integer)
    Public ПодтверждениеПароля As Boolean
    Public dtPer As DataTable
    Public dtZak As DataTable
    Public dtПеревозчики As DataTable
    Public dtКлиенты, dtФормаСобствAll As DataTable
    Public dtТипАвтоAll As DataTable
    Public Procenty20 As String = "False"
    Public КалендарьПовтор As Boolean = False





    Public Sub COMxt(ByVal form As Form, ByVal strsql As String, ByVal c As ComboBox)
        'Dim strsql As String = "SELECT DISTINCT Страна FROM Страна ORDER BY Страна"
        'Dim sty As String = form.Name
        'Dim c1 As ComboBox = form.Controls("ComboBox" & c)


        If c.InvokeRequired Then
            form.Invoke(New coxt(AddressOf COMxt), form, strsql, c)
        Else
            Dim ds As DataTable = Selects3(strsql)
            c.AutoCompleteCustomSource.Clear()
            c.Items.Clear()
            For Each r As DataRow In ds.Rows
                c.AutoCompleteCustomSource.Add(r.Item(0).ToString())
                c.Items.Add(r(0).ToString)
            Next
        End If
    End Sub
    Public Sub Listxt(ByVal form As Form, ByVal strsql As String, ByVal c As ListBox, ByVal com As ComboBox)
        'Dim strsql As String = "SELECT DISTINCT Страна FROM Страна ORDER BY Страна"
        'Dim sty As String = form.Name
        'Dim c1 As ComboBox = form.Controls("ComboBox" & c)


        If c.InvokeRequired Then
            form.Invoke(New coxt1(AddressOf Listxt), form, strsql, c, com)
        Else
            Dim ds As DataTable = Selects3(strsql)
            com.Items.Clear()
            com.AutoCompleteCustomSource.Clear()
            c.Items.Clear()
            For Each r As DataRow In ds.Rows
                c.Items.Add(r(0).ToString)
                com.Items.Add(r(0).ToString)
                com.AutoCompleteCustomSource.Add(r(0).ToString)
            Next
        End If
    End Sub
    'Public Sub Updates(ByVal stroka As String)
    '    Dim c As New OleDbCommand
    '    c.Connection = conn
    '    c.CommandText = stroka
    '    Try
    '        c.ExecuteNonQuery()
    '    Catch ex As Exception
    '        MessageBox.Show(ex.Message)
    '    End Try
    'End Sub
    Public Sub Updates3(ByVal stroka As String)
        Dim conn4 As New SqlConnection(ConString)
        If conn4.State = ConnectionState.Closed Then
            conn4.Open()
        End If
        Dim c As New SqlCommand(stroka, conn4)
        Try
            c.ExecuteNonQuery()
            If conn4.State = ConnectionState.Open Then
                conn4.Close()
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            If conn4.State = ConnectionState.Open Then
                conn4.Close()
            End If
        End Try
    End Sub
    Public Function Selects(ByVal StrSql As String) As DataTable
        errds = 0
        Dim c As New OleDbCommand With {
                .Connection = conn,
                .CommandText = StrSql
            }
        Dim dst As New DataTable
        Dim da As New OleDbDataAdapter(c)
        Try
            da.Fill(dst)
            Dim gf As Object = dst.Rows(0).Item(0)
            Return dst
        Catch ex As Exception
            errds = 1
            Return dst
        End Try

    End Function
    Public Function Selects3(ByVal StrSql As String) As DataTable
        errds = 0
        Dim conn3 As New SqlConnection(ConString)
        If conn3.State = ConnectionState.Closed Then
            conn3.Open()
        End If

        Dim c As New SqlCommand(StrSql, conn3)

        Dim dst As New DataTable

        Dim da As New SqlDataAdapter(c)
        Try
            da.Fill(dst)
            Dim gf As Object = dst.Rows(0).Item(0)

            If conn3.State = ConnectionState.Open Then
                conn3.Close()
            End If
            Return dst
        Catch ex As Exception
            errds = 1
            If conn3.State = ConnectionState.Open Then
                conn3.Close()
            End If
            Return dst
        End Try

    End Function
    Public Function Selects3(ByVal StrSql As String, ByVal Par As List(Of Date)) As DataTable
        errds = 0
        Dim conn3 As New SqlConnection(ConString)
        If conn3.State = ConnectionState.Closed Then
            conn3.Open()
        End If
        Dim c As New SqlCommand(StrSql, conn3)

        If Par.Count > 0 Then
            For x As Integer = 0 To Par.Count - 1
                c.Parameters.AddWithValue("@d" & x + 1, Par(x))
            Next
        End If
        Dim dst As New DataTable

        Dim da As New SqlDataAdapter(c)
        Try
            da.Fill(dst)
            Dim gf As Object = dst.Rows(0).Item(0)

            If conn3.State = ConnectionState.Open Then
                conn3.Close()
            End If
            Return dst
        Catch ex As Exception
            errds = 1
            If conn3.State = ConnectionState.Open Then
                conn3.Close()
            End If
            Return dst
        End Try

    End Function
    Public Function SelectsAsync(ByVal StrSql As String) As DataTable
        errds = 0
        Dim conn3 As New SqlConnection(ConString)
        If conn3.State = ConnectionState.Closed Then
            conn3.Open()
        End If

        Dim c As New SqlCommand(StrSql, conn3)

        Dim dst As New DataTable

        Dim da As New SqlDataAdapter(c)
        Try
            da.Fill(dst)
            Dim gf As Object = dst.Rows(0).Item(0)

            If conn3.State = ConnectionState.Open Then
                conn3.Close()
            End If
            Return dst
        Catch ex As Exception
            errds = 1
            If conn3.State = ConnectionState.Open Then
                conn3.Close()
            End If
            Return dst
        End Try

    End Function
    Public Sub Справки(ByVal год As Integer)

        Try
            FilesПолнПуть = (IO.Directory.GetFiles("Z:\RICKMANS\" & год, "*.xls*", IO.SearchOption.TopDirectoryOnly))
            'Files3 = (IO.Directory.GetFiles("Z:\RICKMANS\" & год, "*.xls*", IO.SearchOption.TopDirectoryOnly))
            Files3 = New List(Of ПутиДоков)

            For Each b In FilesПолнПуть
                Dim f As New ПутиДоков With {.Путь = IO.Path.GetFileName(b), .ПолныйПуть = b}
                Files3.Add(f)
            Next
        Catch ex As Exception
            If Not IO.Directory.Exists("Z:\RICKMANS\" & год) Then
                IO.Directory.CreateDirectory("Z:\RICKMANS\" & год)
            End If
        End Try


    End Sub


    'Public Sub KillProc()
    '    Try
    '        If IO.Directory.Exists("c: \Users\Public\Documents\Рикманс") Then
    '            IO.Directory.Delete("c:\Users\Public\Documents\Рикманс", True)
    '            IO.Directory.CreateDirectory("c:\Users\Public\Documents\Рикманс")
    '        Else
    '            IO.Directory.CreateDirectory("c:\Users\Public\Documents\Рикманс")
    '        End If
    '    Catch ex As Exception

    '        For Each p As Process In Process.GetProcessesByName("winword")
    '            p.Kill()
    '            p.WaitForExit()
    '        Next
    '        If IO.Directory.Exists("c:\Users\Public\Documents\Рикманс") Then
    '            IO.Directory.Delete("c:\Users\Public\Documents\Рикманс", True)
    '            IO.Directory.CreateDirectory("c:\Users\Public\Documents\Рикманс")
    '        Else
    '            IO.Directory.CreateDirectory("c:\Users\Public\Documents\Рикманс")
    '        End If

    '    End Try
    'End Sub

    Public Sub releaseobject(ByVal obj As Object)
        Try
            Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub
    Public Sub ПечатьДоков(ByVal mass As String, ByVal d As Integer)

        Dim wdApp As New Microsoft.Office.Interop.Word.Application
        Dim wdDoc As Microsoft.Office.Interop.Word.Document
        wdApp.Visible = False
        wdDoc = wdApp.Documents.Open(FileName:=mass) 'заявление
        Try
            wdDoc.PrintOut(True,,,,,,, d)
        Catch ex As Exception

        End Try

        'wdApp.Visible = True


        wdDoc.Close()
        wdApp.Quit()
    End Sub
    Public Sub ПечатьДоковЭксель(ByVal mass As String, ByVal d As Integer)

        Dim dr As Boolean = mass.Contains("ДОГОВОР")
        If dr = False Then
            MessageBox.Show("Возможно распечатать только 'Договор' в формает EXCEL!", Рик)
            Exit Sub
        End If

        Dim xlapp As Microsoft.Office.Interop.Excel.Application
        Dim xlworkbook As Microsoft.Office.Interop.Excel.Workbook
        Dim xlworksheet As Microsoft.Office.Interop.Excel.Worksheet
        'Dim misvalue As Object = Reflection.Missing.Value
        xlapp = New Microsoft.Office.Interop.Excel.Application With {
            .Visible = False
        }
        'xlworkbook = xlapp.Workbooks.Add(misvalue)
        xlworkbook = xlapp.Workbooks.Open(mass,, True)
        xlworksheet = xlworkbook.Sheets("пер")
        Try
            xlworksheet.PrintOutEx(True,, d)
        Catch ex As Exception

        End Try



        xlworkbook.Close(False)
        xlapp.Quit()

        releaseobject(xlapp)
        releaseobject(xlworkbook)
        releaseobject(xlworksheet)
    End Sub

    Public Function ФИОКорРук(ByVal ФИОПол As String, ByVal рукИП As Boolean)

        Dim РуковИП As String
        If рукИП = True Then
            РуковИП = "ИП "
        Else
            РуковИП = ""
        End If



        Dim nm As String = ФИОПол
        Dim nm0 As Integer = Len(ФИОПол)
        Dim nm1 As String = Strings.Left(nm, InStr(nm, " "))
        Dim nm2 As Integer = Len(nm1)
        Dim nm3 As String = Strings.Right(nm, (nm0 - nm2))
        Dim nm31 As Integer = Len(nm3)
        Dim nm4 As String = Strings.UCase(Strings.Left(Strings.Left(nm3, InStr(nm3, " ")), 1))
        Dim nm41 As Integer = Len(Strings.Left(nm3, InStr(nm3, " ")))
        Dim nm5 As String = Strings.UCase(Strings.Left(Strings.Right(nm3, nm31 - nm41), 1))




        Dim ФИОКор As String = РуковИП & nm1 & "" & nm4 & "." & nm5 & "."
        Return ФИОКор
    End Function
    Public Sub ОткрытиеФайловБезПути(ByVal d As String)

        Dim currentdirectory = "Z:\RICKMANS"
        Dim filename = d
        Dim path = System.IO.Directory.GetFiles(currentdirectory, "*" & filename, IO.SearchOption.AllDirectories)(0)
        Process.Start(path)

    End Sub
    Public Function РасчетнСчет(ByVal d As String)

        Dim бел As String = "Белорусские.рубли: IBAN - BY90TECN30125856200100000000 в ОАО «Технобанк» РКЦ №3 г Минск, пр.Независимости 117а; БИК Банка - TECNBY22;"
        Dim евро As String = "EURO: Сorr. Bank: Raiffeisen Bank International AG Vienna, Austria; SWIFT: RZBAATWW; Acc. 155070767; Ben. Bank: JSC Technobank, Minsk; SWIFT: TECNBY22, IBAN - BY92TECN30125856220210000000, (BIC: TECNBY22);"
        Dim доллар As String = "USD: Сorr. Bank: Raiffeisen Bank International AG,Vienna, Austria SWIFT: RZBAATWW; Acc. 7055070767; Ben. Bank: JSC Technobank, Minsk; SWIFT: TECNBY22, IBAN - BY40TECN30125856220180000000 (BIC: TECNBY22);"
        Dim россрубл As String = "RUB: ПАО Сбербанк Москва, РФ корсчет 30101810400000000225 в ОПЕРУ Московского ГТУ Банка России, г.Москва, БИК: 044525225, ИНН: 7707083893 SWIFT: SABRRUMM; Банковский счет. 30111810600000000163; Банк бенефициар: ОАО Технобанк, Минск; SWIFT: TECNBY22, IBAN - BY91TECN30125856220340000000, (BIC: TECNBY22);"

        Select Case d
            Case "евро"
                Return евро
            Case "доллар"
                Return доллар
            Case "россрубл"
                Return россрубл
        End Select

        Return бел

    End Function

    Public Sub GridView(ByVal d As DataGridView)
        d.EnableHeadersVisualStyles = False
        d.ColumnHeadersDefaultCellStyle.BackColor = Color.LightGreen
        d.ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.Raised
        d.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        d.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        d.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
    End Sub
    Public Function РейсыКлиента(ByVal num As Integer) As DataRow()
        Dim rowzak = dtZak.Select("НомерРейса=" & num & "")
        Return rowzak
    End Function

    Public Function РейсыПеревозчик(ByVal num As Integer) As DataRow()
        Dim rowper = dtPer.Select("НомерРейса=" & num & "")
        Return rowper
    End Function
    'Public Function dtVyborka() As Task
    '    Return Task.Run(Sub() RunMoving())
    'End Function
    Public Sub RunMoving()
        'dtZak = SelectsAsync(StrSql:="SELECT * FROM РейсыКлиента") 'Рейсы клиента
    End Sub
    'Public Function dtVyborka1() As Task
    '    Return Task.Run(Sub() RunMoving1())
    'End Function
    Public Sub RunMoving1()
        'dtPer = New DataTable()
        'dtPer = SelectsAsync(StrSql:="SELECT * FROM РейсыПеревозчика") 'Рейсы перевозчика
    End Sub

    'Public Function Перевозчики() As Task
    '    Return Task.Run(Sub() ПеревозчикиRunMoving())
    'End Function
    Public Sub ПеревозчикиRunMoving()
        dtПеревозчики = Selects3(StrSql:="SELECT * FROM Перевозчики ORDER BY Названиеорганизации") 'данные по перевозчику
    End Sub

    'Public Function Клиенты() As Task
    '    Return Task.Run(Sub() КлиентыRunMoving())
    'End Function
    Public Sub КлиентыRunMoving()
        dtКлиенты = Selects3(StrSql:="SELECT * FROM Клиент ORDER BY НазваниеОрганизации") 'данные по клиенту
    End Sub

    'Public Function ТипАвтоAll() As Task
    '    Return Task.Run(Sub() ТипАвтоAllRunMoving())
    'End Function
    Public Sub ТипАвтоAllRunMoving()
        dtТипАвтоAll = Selects3(StrSql:="SELECT * FROM ТипАвто ORDER BY ТипАвто")
    End Sub

    Public Sub RunMoving10()
        dtФормаСобствAll = Selects3(StrSql:="SELECT * FROM ФормаСобств") 'ФормаСобств
    End Sub

    Public Async Sub dtSФормаСобств()
        Await Task.Run(Sub() RunMoving10())
    End Sub






    Public Async Sub dtVyborkaS()
        Await Task.Run((Sub() RunMoving()))
    End Sub
    Public Async Sub dtVyborkaS1()
        Await Task.Run((Sub() RunMoving1()))
    End Sub
    Public Async Sub ТипАвтоAll()
        Await Task.Run((Sub() ТипАвтоAllRunMoving()))
    End Sub
    Public Async Sub Клиенты2()
        Await Task.Run((Sub() КлиентыRunMoving()))
    End Sub
    Public Async Sub Перевозчики2()
        Await Task.Run((Sub() ПеревозчикиRunMoving()))
    End Sub

    Public Sub ALL()
        dtVyborkaS1()
        Перевозчики2()
        dtVyborkaS()
        Клиенты2()
        ТипАвтоAll()
        dtSФормаСобств()

    End Sub



End Module