Imports System.Windows.Forms
Imports System.Threading
Imports System.Net
Imports System.Reflection
Imports Timer = System.Threading.Timer
Imports WMPLib
Imports System.Configuration
'Imports System.Reflection

Public Class MDIParent1

    Dim dgh As DataTable


    Private Sub ShowNewForm(ByVal sender As Object, ByVal e As EventArgs)
        ' Создать новый экземпляр дочерней формы.
        Dim ChildForm As New System.Windows.Forms.Form
        ' Сделать ее дочерней для данной формы MDI перед отображением.
        ChildForm.MdiParent = Me

        m_ChildFormNumber += 1
        ChildForm.Text = "Окно " & m_ChildFormNumber

        ChildForm.Show()
    End Sub

    Private Sub OpenFile(ByVal sender As Object, ByVal e As EventArgs)
        Dim OpenFileDialog As New OpenFileDialog
        OpenFileDialog.InitialDirectory = My.Computer.FileSystem.SpecialDirectories.MyDocuments
        OpenFileDialog.Filter = "Текстовые файлы (*.txt)|*.txt|Все файлы (*.*)|*.*"
        If (OpenFileDialog.ShowDialog(Me) = System.Windows.Forms.DialogResult.OK) Then
            Dim FileName As String = OpenFileDialog.FileName
            ' TODO: добавьте здесь код открытия файла.
        End If
    End Sub

    Private Sub SaveAsToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs)
        Dim SaveFileDialog As New SaveFileDialog
        SaveFileDialog.InitialDirectory = My.Computer.FileSystem.SpecialDirectories.MyDocuments
        SaveFileDialog.Filter = "Текстовые файлы (*.txt)|*.txt|Все файлы (*.*)|*.*"

        If (SaveFileDialog.ShowDialog(Me) = System.Windows.Forms.DialogResult.OK) Then
            Dim FileName As String = SaveFileDialog.FileName
            ' TODO: добавить код для сохранения содержимого формы в файл.
        End If
    End Sub


    Private Sub ExitToolsStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs)
        Me.Close()
    End Sub

    Private Sub CutToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs)
        ' Использовать My.Computer.Clipboard для помещения выбранного текста или изображений в буфер обмена
    End Sub

    Private Sub CopyToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs)
        ' Использовать My.Computer.Clipboard для помещения выбранного текста или изображений в буфер обмена
    End Sub

    Private Sub PasteToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs)
        'Использовать My.Computer.Clipboard.GetText() или My.Computer.Clipboard.GetData для получения информации из буфера обмена.
    End Sub

    Private Sub ToolBarToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs)
        'Me.ToolStrip.Visible = Me.ToolBarToolStripMenuItem.Checked
    End Sub

    Private Sub StatusBarToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs)
        'Me.StatusStrip.Visible = Me.StatusBarToolStripMenuItem.Checked
    End Sub

    Private Sub CascadeToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs)
        Me.LayoutMdi(MdiLayout.Cascade)
    End Sub

    Private Sub TileVerticalToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs)
        Me.LayoutMdi(MdiLayout.TileVertical)
    End Sub

    Private Sub TileHorizontalToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs)
        Me.LayoutMdi(MdiLayout.TileHorizontal)
    End Sub

    Private Sub ArrangeIconsToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs)
        Me.LayoutMdi(MdiLayout.ArrangeIcons)
    End Sub

    Private Sub CloseAllToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs)
        ' Закрыть все дочерние формы указанного родителя.
        For Each ChildForm As Form In Me.MdiChildren
            ChildForm.Close()
        Next
    End Sub

    Private m_ChildFormNumber As Integer

    Private Sub ДобавитьToolStripMenuItem_Click(sender As Object, e As EventArgs)

    End Sub


    Private Sub ДобавитьToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles ДобавитьToolStripMenuItem1.Click
        Try
            ActiveMdiChild.Close()
        Catch ex As Exception

        End Try
        Данные.MdiParent = Me
        Данные.Show() 'publish\РикмансГрузы\
    End Sub

    Private Sub ВыборкаToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles ВыборкаToolStripMenuItem1.Click

        'ActiveMdiChild.Hide()
        Dim f As Form
        f = ActiveMdiChild

        Выборка.MdiParent = Me
        'Dim f As New Task(New )
        Выборка.Show()
        'If Not f.Name = "Выборка" Then
        '    Try
        '        f.Close()
        '    Catch ex As Exception

        '    End Try

        'End If

        ПровФормы(f, "Выборка")
    End Sub
    Private Sub ПровФормы(ByVal f As Form, ByVal g As String)
        If f Is Nothing Then
            Exit Sub
        End If
        If Not f.Name = g Then
            Try
                f.Close()
            Catch ex As Exception

            End Try

        End If
    End Sub
    Private Sub НайтиToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles НайтиToolStripMenuItem.Click

        Dim f As Form
        f = ActiveMdiChild
        Dim f1 As New ПоискПеревозчиков
        f1.MdiParent = Me
        f1.Show()
        ПровФормы(f, "ПоискПеревозчиков")

    End Sub

    Private Sub ДобавитьToolStripMenuItem_Click_1(sender As Object, e As EventArgs) Handles ДобавитьToolStripMenuItem.Click



    End Sub
    Private Sub КалендарьНапоминание()
        Dim listes As New Dictionary(Of String, String)
        Using db As New dbAllDataContext(_cn3)

            Dim var2 = db.КалендарьНапоминание.Where(Function(x) x.ДатаНапоминания = Now.Date And x.Пользователь = Экспедитор).Select(Function(x) x).ToList()
            If var2.Count > 0 Then
                For Each r In var2
                    If listes.ContainsKey(r.ВремяНапоминания) Then Continue For
                    listes.Add(r.ВремяНапоминания, r.ТекстНапоминания)
                Next

            End If
        End Using
        If listes.Count > 0 Then
            Dim f As New PublicClass(listes)
        End If



    End Sub
    Public Async Sub КалендарьНапоминаниеAsync()
        Await Task.Run(Sub() КалендарьНапоминание())
    End Sub
    Private Sub Предзагрузка1()
        Dim mo As New AllUpd
        mo.ГрузыКлиентовAll()
        mo.ЖурналДатаAll()
        mo.ЖурналКлиентГрузAll()
        mo.ЖурналКлиентМаршрутAll()
        mo.ЖурналКлиентСписокAll()
        mo.ЖурналПеревозчикAll()
        mo.Календарь_ДатыAll()
        mo.КалендарьНапоминаниеAll()
        mo.КалендарьРезультатЗвонкаAll()
        mo.ПеревозчикиБазаAll()
        mo.ЧерныйСписокAll()
    End Sub
    Private Async Sub Предзагрузка1Async()
        Await Task.Run(Sub() Предзагрузка1())
    End Sub
    Private Async Sub ПредзагрузкаAsync()
        Await Task.Run(Sub() Предзагрузка())
    End Sub
    Private Sub Предзагрузка()
        Dim mo As New AllUpd
        mo.SkypeКлиентПредложениеAll()
        mo.SkypeПеревозчикПредложениеAll()
        mo.КлиентAll()
        mo.ПеревозчикиAll()
        mo.ПаролиAll()
        mo.РейсыКлиентаAll()
        mo.ОтчетРаботыСотрудникаAll()
        mo.ОтчетРаботыСотрудникаСводнаяAll()
        mo.ОплатыКлиентAll()
        mo.ОплатыПерAll()
        mo.РейсыПеревозчикаAll()
    End Sub
    'Private Sub getLoc()
    '    Dim f = Assembly.GetExecutingAssembly().Location

    '    Dim m = f
    'End Sub
    Public timer4 As Timer
    Private interVal4 As Long = 3600000

    Private Sub КонтрольОбновленияБазы()

        Using db As New dbAllDataContext(_cn3)
            Dim f = (From x In db.Календарь_Даты
                     Where x.Дата = Now
                     Select x).FirstOrDefault()
            If f IsNot Nothing Then
                КаленДаты = f
            End If
        End Using

        timer4 = New Timer(New TimerCallback(Sub() НапоминаниеКалендарь()), Nothing, 0, interVal4)

    End Sub

    Private Sub НапоминаниеКалендарь()
        If КаленДаты Is Nothing Then
            timer4.Dispose()
            Return
        End If
        If Экспедитор = "Олег" Then
            Dim m = Now.Hour
            If m >= "8" And m < "9" Then
                Dim ms = КаленДаты._8_00
                If ms IsNot Nothing Then
                    Dim m1 As New КалендарьВсплывФорма("8.00 - 9.00", ms)
                    m1.ShowDialog()

                End If
            End If

            If m >= "9" And m < "10" Then
                Dim ms = КаленДаты._9_00
                If ms IsNot Nothing Then
                    Dim m1 As New КалендарьВсплывФорма("9.00 - 10.00", ms)
                    m1.ShowDialog()
                End If
            End If

            If m >= "10" And m < "11" Then
                Dim ms = КаленДаты._10_00
                If ms IsNot Nothing Then
                    Dim m1 As New КалендарьВсплывФорма("10.00 - 11.00", ms)
                    m1.ShowDialog()
                End If
            End If

            If m >= "11" And m < "12" Then
                Dim ms = КаленДаты._11_00
                If ms IsNot Nothing Then
                    Dim m1 As New КалендарьВсплывФорма("11.00 - 12.00", ms)
                    m1.ShowDialog()
                End If
            End If

            If m >= "12" And m < "13" Then
                Dim ms = КаленДаты._12_00
                If ms IsNot Nothing Then
                    Dim m1 As New КалендарьВсплывФорма("12.00 - 13.00", ms)
                    m1.ShowDialog()
                End If
            End If

            If m >= "13" And m < "14" Then
                Dim ms = КаленДаты._13_00
                If ms IsNot Nothing Then
                    Dim m1 As New КалендарьВсплывФорма("13.00 - 14.00", ms)
                    m1.ShowDialog()
                End If
            End If

            If m >= "14" And m < "15" Then
                Dim ms = КаленДаты._14_00
                If ms IsNot Nothing Then
                    Dim m1 As New КалендарьВсплывФорма("14.00 - 15.00", ms)
                    m1.ShowDialog()
                End If
            End If

            If m >= "15" And m < "16" Then
                Dim ms = КаленДаты._15_00
                If ms IsNot Nothing Then
                    Dim m1 As New КалендарьВсплывФорма("15.00 - 16.00", ms)
                    m1.ShowDialog()
                End If
            End If

            If m >= "16" And m < "17" Then
                Dim ms = КаленДаты._16_00
                If ms IsNot Nothing Then
                    Dim m1 As New КалендарьВсплывФорма("16.00 - 17.00", ms)
                    m1.ShowDialog()
                End If
            End If

            If m >= "17" And m < "18" Then
                Dim ms = КаленДаты._17_00
                If ms IsNot Nothing Then
                    Dim m1 As New КалендарьВсплывФорма("17.00 - 18.00", ms)
                    m1.ShowDialog()
                End If
            End If

            If m >= "18" And m < "19" Then
                Dim ms = КаленДаты._18_00
                If ms IsNot Nothing Then
                    Dim m1 As New КалендарьВсплывФорма("18.00 - 19.00", ms)
                    m1.ShowDialog()
                End If
            End If

            If m >= "19" And m < "20" Then
                Dim ms = КаленДаты._19_00
                If ms IsNot Nothing Then
                    Dim m1 As New КалендарьВсплывФорма("19.00 - 20.00", ms)
                    m1.ShowDialog()
                End If
            End If

            If m >= "20" And m < "21" Then
                Dim ms = КаленДаты._20_00
                If ms IsNot Nothing Then
                    Dim m1 As New КалендарьВсплывФорма("20.00 - 21.00", ms)
                    m1.ShowDialog()
                End If
            End If

            If m >= "21" And m < "22" Then
                Dim ms = КаленДаты._21_00
                If ms IsNot Nothing Then
                    Dim m1 As New КалендарьВсплывФорма("21.00 - 22.00", ms)
                    m1.ShowDialog()
                End If
            End If

            If m >= "22" And m < "23" Then
                Dim ms = КаленДаты._22_00
                If ms IsNot Nothing Then
                    Dim m1 As New КалендарьВсплывФорма("22.00 - 23.00", ms)
                    m1.ShowDialog()
                End If
            End If

        End If

    End Sub
    Private Sub _проверкаСтрокиПодключения()
        XXX = ConfigurationManager.ConnectionStrings("OleVnytr").ConnectionString  'внутренний
        _cn3 = Global.Грузы.My.MySettings.Default.RickmansConnectionString   'внутренний
        ConString = Global.Грузы.My.MySettings.Default.RickmansConnectionString  'внутренний
        Try
            Using db As New dbAllDataContext(_cn3)
                Dim f = db.Пароли.FirstOrDefault()

            End Using
        Catch ex As Exception
            XXX = ConfigurationManager.ConnectionStrings("OleVneshn").ConnectionString  'внутренний
            _cn3 = Global.Грузы.My.MySettings.Default.RickmansConnectionString1
            ConString = Global.Грузы.My.MySettings.Default.RickmansConnectionString1
        End Try
    End Sub
    Private Sub MDIParent1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        'getLoc()

        _проверкаСтрокиПодключения()
        Me.Cursor = Cursors.WaitCursor
        'ALL()

        Предзагрузка1Async()
        ПредзагрузкаAsync()

        Dim f As New Пароль
        f.ShowDialog()

        If f.flg = False Then
            Return
        End If

        If Экспедитор = "Олег" Then
            ФинанасыTool.Enabled = True
        End If

        КонтрольОбновленияБазы()
        Getst()
        КалендарьНапоминаниеAsync()
        Reys()
        Me.Cursor = Cursors.Default
    End Sub
    Private Sub Reys()
        Dim f As Form
        f = ActiveMdiChild
        Dim f1 As New Рейс
        f1.MdiParent = Me
        f1.Show()
        ПровФормы(f, "Рейс")
    End Sub
    'Private Function too() As Task(Of String)
    '    'Dim d As New List(Of Task)()
    '    'd.Add(Task.Run(dtVyborka1, dtVyborka))
    '    '        Task.WhenAll(New Task(New Action(Sub() dtVyborka1())), New Task(New Action(Sub() Перевозчики())), New Task(New Action(Sub() Клиенты())),
    '    'New Task(New Action(Sub() ТипАвтоAll())))
    '    'Await dtVyborka()
    '    'Await dtVyborka1()
    '    'Await Перевозчики()
    '    'Await Клиенты()
    '    'Await ТипАвтоAll()

    '    'Dim m As Task
    '    'm.Run(New Action(Sub() dtVyborka()))
    '    'm.GetAwaiter()

    '    'Await Task.Run(New Action(Sub() dtVyborka1()))
    '    'Await Task.Run(New Action(Sub() Перевозчики()))
    '    'Await Task.Run(New Action(Sub() Клиенты()))
    '    'Await Task.Run(New Action(Sub() ТипАвтоAll()))




    'End Function
    Public Sub CheckForUpdate()
        Dim verFile As String = Application.StartupPath & "\version.txt"
        Dim curVer As String = My.Application.Info.Version.ToString
        If My.Computer.FileSystem.FileExists(verFile) Then My.Computer.FileSystem.DeleteFile(verFile)
        My.Computer.Network.DownloadFile("http://2trans.by/oleg/version.txt", verFile)
        Dim Lastver As String = My.Computer.FileSystem.ReadAllText(verFile)
        If Not curVer = Lastver Then
            If MessageBox.Show("Доступно обновление" & vbCrLf & "Скачать?", Рик, MessageBoxButtons.YesNo) = DialogResult.Yes Then
                'Shell("Updater.exe", vbNormalFocus)
                End
                Me.Close()
            End If
        End If
    End Sub

    Private Sub ПолныйToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ПолныйToolStripMenuItem.Click
        Dim f As Form
        'ActiveMdiChild.Hide()
        f = ActiveMdiChild
        Dim f1 As New ДобавитьПеревозчика
        f1.MdiParent = Me
        ДобПер = 0
        f1.Show()

        ПровФормы(f, "ДобавитьПеревозчика")

    End Sub

    Private Sub БыстроToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles БыстроToolStripMenuItem.Click
        ДобПер = 1
        Dim f As New ДобавитьПеревозчика
        f.ShowDialog()
    End Sub

    Private Sub ТранспортToolStripMenuItem_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub СоздатьToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles СоздатьToolStripMenuItem.Click

        Dim f As Form
        f = ActiveMdiChild
        Dim f1 As New Рейс
        f1.MdiParent = Me
        f1.Show()
        ПровФормы(f, "Рейс")

    End Sub

    Private Sub КлиентToolStripMenuItem_Click(sender As Object, e As EventArgs)
        Dim f As New НовыйКлиент
        f.ShowDialog()

    End Sub

    Private Sub ПеревозчикToolStripMenuItem_Click(sender As Object, e As EventArgs)
        Dim f As New НовыйПеревоз
        f.ShowDialog()
    End Sub

    Private Sub ПечатьToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ПечатьToolStripMenuItem.Click

    End Sub

    Private Sub РейсToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles РейсToolStripMenuItem1.Click
        Dim f As New ПечатьРейсы
        f.ShowDialog()
    End Sub

    Private Sub ДоговораToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles ДоговораToolStripMenuItem1.Click
        Dim f As New ПечатьДоговора
        f.ShowDialog()
    End Sub

    Private Sub ПереговорыToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ПереговорыToolStripMenuItem.Click

    End Sub

    Private Sub ЧерныйСписокToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ЧерныйСписокToolStripMenuItem.Click
        Dim f As New ЧерныйСписокForm
        f.ShowDialog()
    End Sub

    Public Function Провер() As List(Of ПереговорыКлиент)
        Dim var As List(Of ПереговорыКлиент)
        Using db As New dbAllDataContext(_cn3)
            var = db.ПереговорыКлиент.Where(Function(x) x.ДатаНапоминания = Now.Date And x.Экспедитор = Экспедитор).Select(Function(x) x).ToList()

        End Using
        If var.Count > 0 Then
            Return var
        Else
            Dim g As New List(Of ПереговорыКлиент)
            Return g
        End If


        'Dim DateEx As String = Format(Now.Date, "MM\/dd\/yyyy")
        'Dim strsql As String = "SELECT * FROM ПереговорыКлиент WHERE ДатаНапоминания=#" & DateEx & "# and Экспедитор='" & Экспедитор & "'"
        ''Dim ds As DataTable = Selects3(strsql)
        'Dim ds As DataTable = Selects3(strsql)
        'Return ds
    End Function
    Private Sub Getst()
        Dim df As List(Of ПереговорыКлиент) = Провер()
        If df.Count = 1 Then
            Dim f As New ВсплывФормаПереговоры()

            f.RichTextBox1.Text = ""
            f.RichTextBox2.Text = ""
            f.RichTextBox3.Text = ""
            f.RichTextBox1.Text = df(0).Клиент
            f.RichTextBox2.Text = df(0).ТекстНапоминания
            If df(0).ДатаНапоминания IsNot Nothing Then
                f.RichTextBox3.Text = Strings.Left(df(0).ДатаНапоминания.ToString, 10)
            End If
            f.ShowDialog()
        ElseIf df.Count > 1 Then
            Dim f As New ВсплывФормаПереговоры()
            f.GroupBox1.Visible = True
            f.Label4.Text = df.Count
            f.Label5.Text = df.Count
            f.Label5.Visible = True
            f.Загр(df)
            f.ShowDialog()
        Else
            Exit Sub
        End If
    End Sub

    Private Sub НапоминаниеToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles НапоминаниеToolStripMenuItem.Click
        Dim f As New Напоминание
        f.ShowDialog()
    End Sub

    Private Sub СвободныйТранспортToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles СвободныйТранспортToolStripMenuItem.Click
        Dim f As New ПеревозВПути
        f.ShowDialog()
    End Sub

    Private Sub ПоОрганизацииToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ПоОрганизацииToolStripMenuItem.Click
        Dim f As New ПереговорыКлиентФорма
        f.ShowDialog()

    End Sub

    Private Sub ВсеПереговорыToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ВсеПереговорыToolStripMenuItem.Click
        Try
            ActiveMdiChild.Close()
        Catch ex As Exception

        End Try
        Dim f As New ПереговорыВсе
        f.Show()

    End Sub



    Private Sub ОтчетToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ОтчетToolStripMenuItem.Click
        Dim f As New Паролик
        f.ShowDialog()
        If ПодтверждениеПароля = True Then
            Dim f1 As New ОтчетДляОлега
            f1.Show()
        End If
    End Sub


    Private Sub КалендарьToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles КалендарьToolStripMenuItem.Click
        If ActiveMdiChild IsNot Nothing Then
            ActiveMdiChild.Close()
        End If

        If My.Computer.Name.ToString = "OLEGLAPTOP" Then
            Dim f As New Календарь
            f.MdiParent = Me
            f.Show()
        End If

    End Sub

    Private Sub ОтчетToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles ОтчетToolStripMenuItem1.Click
        If ActiveMdiChild IsNot Nothing Then
            ActiveMdiChild.Close()
        End If
        Dim f As New ГотовыйОтчет
        f.Show()

    End Sub

    Private Sub ЖурналToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ЖурналToolStripMenuItem.Click
        If ActiveMdiChild IsNot Nothing Then
            ActiveMdiChild.Close()
        End If
        Dim f As New Журнал
        f.Show()

    End Sub

    Private Sub ExcelFilesAddToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExcelFilesAddToolStripMenuItem.Click
        Dim f As New ExcelAddToDatabase
        f.ShowDialog()
    End Sub

    Private Sub ПоискToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ПоискToolStripMenuItem.Click
        Dim f9 As New ПоискПолный()
        f9.MdiParent = Me
        f9.WindowState = FormWindowState.Maximized
        f9.Show()

    End Sub

    Private Sub SQLDepToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SQLDepToolStripMenuItem.Click
        Dim f As New SQLDependency2
        f.ShowDialog()
    End Sub

    Private Sub ImageToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ImageToolStripMenuItem.Click
        Dim f As New ImageForm
        f.ShowDialog()
    End Sub

    Private Sub БраузерToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles БраузерToolStripMenuItem.Click
        Dim f As New Браузер
        f.MdiParent = Me
        f.Show()
    End Sub

    Private Sub ФинанасыTool_Click(sender As Object, e As EventArgs) Handles ФинанасыTool.Click
        If Экспедитор = "Олег" Then
            Dim f As New Финансы
            f.Show()
        End If

    End Sub

    Private Sub СводнаяОплатToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles СводнаяОплатToolStripMenuItem.Click
        Dim f As New Сводная
        f.Show()
    End Sub



    Private Sub StartToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles StartToolStripMenuItem.Click
        Dim f As New Radio
        f.Show()
    End Sub

    Private Sub StopToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles StopToolStripMenuItem.Click

        Frec(0).controls.stop()


        'Dim f = Process.GetProcessesByName("wmplayer")
        'If f.Length > 0 Then
        '    f(f.Length - 1).Kill()
        'End If

    End Sub
End Class
