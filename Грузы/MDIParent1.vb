Imports System.Windows.Forms
Imports System.Threading
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
        ПоискПеревозчиков.MdiParent = Me
        ПоискПеревозчиков.Show()
        ПровФормы(f, "ПоискПеревозчиков")

    End Sub

    Private Sub ДобавитьToolStripMenuItem_Click_1(sender As Object, e As EventArgs) Handles ДобавитьToolStripMenuItem.Click



    End Sub

    Private Sub MDIParent1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Me.Cursor = Cursors.WaitCursor
        ALL()


        'Dim str As String = Await too()
        'MsgBox(str)

        Пароль.ShowDialog()
        Getst()


        Me.Cursor = Cursors.Default
    End Sub
    Private Async Function too() As Task(Of String)
        'Dim d As New List(Of Task)()
        'd.Add(Task.Run(dtVyborka1, dtVyborka))
        '        Task.WhenAll(New Task(New Action(Sub() dtVyborka1())), New Task(New Action(Sub() Перевозчики())), New Task(New Action(Sub() Клиенты())),
        'New Task(New Action(Sub() ТипАвтоAll())))
        'Await dtVyborka()
        'Await dtVyborka1()
        'Await Перевозчики()
        'Await Клиенты()
        'Await ТипАвтоAll()

        'Dim m As Task
        'm.Run(New Action(Sub() dtVyborka()))
        'm.GetAwaiter()

        'Await Task.Run(New Action(Sub() dtVyborka1()))
        'Await Task.Run(New Action(Sub() Перевозчики()))
        'Await Task.Run(New Action(Sub() Клиенты()))
        'Await Task.Run(New Action(Sub() ТипАвтоAll()))




    End Function
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

        ДобавитьПеревозчика.MdiParent = Me
        ДобПер = 0

        ДобавитьПеревозчика.Show()

        ПровФормы(f, "ДобавитьПеревозчика")

    End Sub

    Private Sub БыстроToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles БыстроToolStripMenuItem.Click
        ДобПер = 1
        ДобавитьПеревозчика.ShowDialog()
    End Sub

    Private Sub ТранспортToolStripMenuItem_Click(sender As Object, e As EventArgs) 

    End Sub

    Private Sub СоздатьToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles СоздатьToolStripMenuItem.Click

        Dim f As Form
        f = ActiveMdiChild
        Рейс.MdiParent = Me
        Рейс.Show()
        ПровФормы(f, "Рейс")

    End Sub

    Private Sub КлиентToolStripMenuItem_Click(sender As Object, e As EventArgs)

        НовыйКлиент.ShowDialog()
    End Sub

    Private Sub ПеревозчикToolStripMenuItem_Click(sender As Object, e As EventArgs)
        НовыйПеревоз.ShowDialog()
    End Sub

    Private Sub ПечатьToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ПечатьToolStripMenuItem.Click

    End Sub

    Private Sub РейсToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles РейсToolStripMenuItem1.Click
        ПечатьРейсы.ShowDialog()
    End Sub

    Private Sub ДоговораToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles ДоговораToolStripMenuItem1.Click
        ПечатьДоговора.ShowDialog()
    End Sub

    Private Sub ПереговорыToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ПереговорыToolStripMenuItem.Click

    End Sub

    Private Sub ЧерныйСписокToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ЧерныйСписокToolStripMenuItem.Click
        ЧерныйСписок.ShowDialog()
    End Sub

    Public Function Провер() As DataTable
        Dim DateEx As String = Format(Now.Date, "MM\/dd\/yyyy")
        Dim strsql As String = "SELECT * FROM ПереговорыКлиент WHERE ДатаНапоминания=#" & DateEx & "# and Экспедитор='" & Экспедитор & "'"
        'Dim ds As DataTable = Selects3(strsql)
        Dim ds As DataTable = Selects3(strsql)
        Return ds
    End Function
    Private Sub Getst()
        Dim df As DataTable = Провер()
        If df.Rows.Count = 1 Then
            ВсплывФормаПереговоры.RichTextBox1.Text = ""
            ВсплывФормаПереговоры.RichTextBox2.Text = ""
            ВсплывФормаПереговоры.RichTextBox3.Text = ""
            ВсплывФормаПереговоры.RichTextBox1.Text = df.Rows(0).Item(1).ToString
            ВсплывФормаПереговоры.RichTextBox2.Text = df.Rows(0).Item(5).ToString
            ВсплывФормаПереговоры.RichTextBox3.Text = Strings.Left(df.Rows(0).Item(4).ToString, 10)
            ВсплывФормаПереговоры.ShowDialog()
        ElseIf df.Rows.Count > 1 Then
            ВсплывФормаПереговоры.GroupBox1.Visible = True
            ВсплывФормаПереговоры.Label4.Text = df.Rows.Count.ToString
            ВсплывФормаПереговоры.Label5.Text = df.Rows.Count.ToString
            ВсплывФормаПереговоры.Label5.Visible = True
            ВсплывФормаПереговоры.Загр(df)
            ВсплывФормаПереговоры.ShowDialog()
        Else
            Exit Sub
        End If
    End Sub

    Private Sub НапоминаниеToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles НапоминаниеToolStripMenuItem.Click
        Напоминание.ShowDialog()
    End Sub

    Private Sub СвободныйТранспортToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles СвободныйТранспортToolStripMenuItem.Click
        ПеревозВПути.ShowDialog()
    End Sub

    Private Sub ПоОрганизацииToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ПоОрганизацииToolStripMenuItem.Click
        ПереговорыКлиент.ShowDialog()
    End Sub

    Private Sub ВсеПереговорыToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ВсеПереговорыToolStripMenuItem.Click
        Try
            ActiveMdiChild.Close()
        Catch ex As Exception

        End Try
        ПереговорыВсе.Show()
    End Sub

    Private Sub ЛИToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ЛИToolStripMenuItem.Click
        'ЛардиТранс.ShowDialog()
    End Sub

    Private Sub ДоговораToolStripMenuItem_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub ОтчетToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ОтчетToolStripMenuItem.Click
        Паролик.ShowDialog()
        If ПодтверждениеПароля = True Then
            ОтчетДляОлега.Show()
        End If
    End Sub

    Private Sub РейсToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles РейсToolStripMenuItem.Click

    End Sub
End Class
