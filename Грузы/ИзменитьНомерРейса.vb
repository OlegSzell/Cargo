Imports System.ComponentModel
Imports System.Reflection

Public Class ИзменитьНомерРейса
    Private Numb As Integer
    Private lst1 As BindingList(Of ПутиДоков)
    Private Pt As ПутиДоков
    Public newPt As ПутиДоков = Nothing

    Private xlapp As Microsoft.Office.Interop.Excel.Application
    Private xlworkbook As Microsoft.Office.Interop.Excel.Workbook
    Private xlworksheet As Microsoft.Office.Interop.Excel.Worksheet
    Private misvalue As Object = Missing.Value
    Sub New(ByVal _Numb As Integer, ByVal _lst As BindingList(Of ПутиДоков), ByVal _pt As ПутиДоков)
        Numb = _Numb
        lst1 = _lst
        Pt = _pt
        ' Этот вызов является обязательным для конструктора.
        InitializeComponent()
        TextBox1.Text = Numb
        ' Добавить код инициализации после вызова InitializeComponent().

    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Close()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If MessageBox.Show("Сохранить изменение?", Рик, MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.No Then
            Return
        End If
        If TextBox2.Text.Length = 0 Then
            MessageBox.Show("Задайте номер рейса или нажмите 'Отмена'", Рик)
            Return
        End If

        Dim f1 = Strings.Left(lst1.Select(Function(x) x.Путь).Last(), 3)
        Dim f = TextBox2.Text

        If f.Length = 1 Then
            f = "00" & f
        ElseIf f.Length = 2 Then
            f = "0" & f
        End If
        If CType(f, Integer) > CType(f1, Integer) Then
            MessageBox.Show("Номер рейса не может быть больше полсденего рейса!", Рик)
            Return
        End If
        ProgressBar1.Maximum = 0
        ProgressBar1.Maximum = 100
        ProgressBar1.Value = 30

        SaveDbAsync(f)

        ProgressBar1.BeginInvoke(New MethodInvoker(Sub() Excel(f))) 'поток параллельно с контролами

    End Sub
    Private Sub Excel(ByVal _Numb As String)
        Cursor = Cursors.WaitCursor
        ProgressBar1.Value = 30
        If Pt Is Nothing Then Return
        Dim lts = Pt.Путь.Length
        Dim ya As Integer = CType(Strings.Right(Strings.Left(Pt.ПолныйПуть, Pt.ПолныйПуть.Length - (lts + 1)), 4), Integer)
        Dim Newпуть = _Numb & Strings.Right(Pt.Путь, Pt.Путь.Length - 3)
        Dim nv As String = "Z:\RICKMANS\" & ya & "\" & Newпуть
        Dim k As New ПутиДоков With {.ПолныйПуть = nv, .Путь = Newпуть}
        newPt = k
        'IO.File.Copy(Pt.ПолныйПуть, nv, True)
        ProgressBar1.Value = 50
        xlapp = New Microsoft.Office.Interop.Excel.Application With {
            .Visible = False
        }

        xlworkbook = xlapp.Workbooks.Add(Pt.ПолныйПуть)

        xlworksheet = xlworkbook.Sheets("ЗАК")
        ProgressBar1.Value = 60
        xlworksheet.Range("N3").Value = _Numb
        xlworksheet.Range("N6").Value = _Numb


        If IO.File.Exists(nv) Then
            IO.File.Delete(nv)
        End If

        xlworkbook.SaveAs(Filename:=nv, FileFormat:=52, CreateBackup:=False)

        ProgressBar1.Value = 70

        xlworkbook.Close(True)
        xlapp.Quit()

        ProgressBar1.Value = 80

        releaseobject(xlapp)
        releaseobject(xlworkbook)
        releaseobject(xlworksheet)
        ProgressBar1.Value = 90
        Dim poln As String = Pt.ПолныйПуть
        IO.File.Delete(poln)
        ProgressBar1.Value = 100
        Cursor = Cursors.Default
        MessageBox.Show("Данные изменены!", Рик)
        Close()
    End Sub
    Private Async Sub SaveDbAsync(ByVal f As String)
        Await Task.Run(Sub() SaveDb(f))
    End Sub
    Private Sub SaveDb(ByVal f As String)
        Dim mo As New AllUpd
        Using db As New dbAllDataContext()
            Dim f2 = db.РейсыКлиента.Where(Function(x) x.НомерРейса = Numb).FirstOrDefault()
            If f2 IsNot Nothing Then
                f2.НомерРейса = f
                db.SubmitChanges()
                mo.РейсыКлиентаAll()
            End If

            Dim f3 = db.РейсыПеревозчика.Where(Function(x) x.НомерРейса = Numb).FirstOrDefault()
            If f3 IsNot Nothing Then
                f3.НомерРейса = f
                db.SubmitChanges()
                mo.РейсыПеревозчикаAll()
            End If
        End Using
    End Sub
    Private Sub releaseobject(ByVal obj As Object, Optional d As List(Of Object) = Nothing)
        If d Is Nothing Then
            Try
                Runtime.InteropServices.Marshal.ReleaseComObject(obj)
                obj = Nothing
            Catch ex As Exception
                MessageBox.Show(ex.Message)
                obj = Nothing
            Finally
                GC.Collect()
            End Try
        Else
            For Each b In d
                Try
                    Runtime.InteropServices.Marshal.ReleaseComObject(b)
                    b = Nothing
                Catch ex As Exception
                    MessageBox.Show(ex.Message)
                    b = Nothing
                Finally
                    GC.Collect()
                End Try
            Next

        End If

    End Sub



End Class