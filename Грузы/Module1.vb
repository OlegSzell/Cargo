Imports System.Data.OleDb

Module Module1
    Public conn As OleDbConnection
    Public conn1 As OleDbConnection
    Public Отдел, Nm As String
    Public idtabl, errds, pro, pr2 As Integer
    Public IDГруз, Отмена As Integer
    Public IDПеревоза, ОтКогоИзмен As Integer
    Public Экспедитор, NameПеревоза As String
    Public ConString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=U:\Офис\Рикманс\ДанныеРикманс.accdb; Persist Security Info=False;"
    Public ДобПер As Integer
    Public Рик As String = "ООО Рикманс"
    Public Files3() As String

    Public Sub Updates(ByVal stroka As String)
        Dim c As New OleDbCommand
        c.Connection = conn
        c.CommandText = stroka
        Try
            c.ExecuteNonQuery()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
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

    Public Sub Справки(ByVal год As Integer)


        Dim gth3 As String

        Try

            Files3 = (IO.Directory.GetFiles("Z:\RICKMANS\" & год, "*.xls*", IO.SearchOption.TopDirectoryOnly))
            For n As Integer = 0 To Files3.Length - 1
                gth3 = ""
                gth3 = IO.Path.GetFileName(Files3(n))
                Files3(n) = gth3
            Next


        Catch ex As Exception

        End Try


    End Sub


    Public Sub KillProc()
        Try
            If IO.Directory.Exists("c:\Users\Public\Documents\Рикманс") Then
                IO.Directory.Delete("c:\Users\Public\Documents\Рикманс", True)
                IO.Directory.CreateDirectory("c:\Users\Public\Documents\Рикманс")
            Else
                IO.Directory.CreateDirectory("c:\Users\Public\Documents\Рикманс")
            End If
        Catch ex As Exception

            For Each p As Process In Process.GetProcessesByName("winword")
                p.Kill()
                p.WaitForExit()
            Next
            If IO.Directory.Exists("c:\Users\Public\Documents\Рикманс") Then
                IO.Directory.Delete("c:\Users\Public\Documents\Рикманс", True)
                IO.Directory.CreateDirectory("c:\Users\Public\Documents\Рикманс")
            Else
                IO.Directory.CreateDirectory("c:\Users\Public\Documents\Рикманс")
            End If

        End Try
    End Sub

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





End Module