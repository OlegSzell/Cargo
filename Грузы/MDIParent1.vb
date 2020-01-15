Imports System.Windows.Forms

Public Class MDIParent1

    Dim dgh As DataTable

    Private Sub ShowNewForm(ByVal sender As Object, ByVal e As EventArgs) Handles NewToolStripButton.Click
        ' Создать новый экземпляр дочерней формы.
        Dim ChildForm As New System.Windows.Forms.Form
        ' Сделать ее дочерней для данной формы MDI перед отображением.
        ChildForm.MdiParent = Me

        m_ChildFormNumber += 1
        ChildForm.Text = "Окно " & m_ChildFormNumber

        ChildForm.Show()
    End Sub

    Private Sub OpenFile(ByVal sender As Object, ByVal e As EventArgs) Handles OpenToolStripButton.Click
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

        Данные.Show() 'publish\РикмансГрузы\
    End Sub

    Private Sub ВыборкаToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles ВыборкаToolStripMenuItem1.Click
        Try
            ActiveMdiChild.Close()
        Catch ex As Exception

        End Try
        Выборка.Show()
    End Sub

    Private Sub НайтиToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles НайтиToolStripMenuItem.Click
        Try
            ActiveMdiChild.Close()
        Catch ex As Exception

        End Try
        ПоискПеревозчиков.Show()


    End Sub

    Private Sub ДобавитьToolStripMenuItem_Click_1(sender As Object, e As EventArgs) Handles ДобавитьToolStripMenuItem.Click



    End Sub

    Private Sub MDIParent1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'CheckForUpdate()

        'MessageBox.Show("3")



        Пароль.ShowDialog()
        Getst()

        If Экспедитор <> "" Then
            ПоискПеревозчиков.Show()
        Else
            Me.Close()
        End If

    End Sub
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
    Private Sub MenuStrip_ItemClicked(sender As Object, e As ToolStripItemClickedEventArgs) Handles MenuStrip.ItemClicked

    End Sub

    Private Sub ПолныйToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ПолныйToolStripMenuItem.Click
        Try
            ActiveMdiChild.Close()
        Catch ex As Exception

        End Try
        ДобПер = 0
        ДобавитьПеревозчика.Show()
    End Sub

    Private Sub БыстроToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles БыстроToolStripMenuItem.Click
        ДобПер = 1
        ДобавитьПеревозчика.ShowDialog()
    End Sub

    Private Sub ТранспортToolStripMenuItem_Click(sender As Object, e As EventArgs) 

    End Sub

    Private Sub СоздатьToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles СоздатьToolStripMenuItem.Click
        Try
            ActiveMdiChild.Close()
        Catch ex As Exception

        End Try
        Рейс.Show()
    End Sub

    Private Sub КлиентToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles КлиентToolStripMenuItem.Click
        НовыйКлиент.ShowDialog()
    End Sub

    Private Sub ПеревозчикToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ПеревозчикToolStripMenuItem.Click
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
        ПереговорыКлиент.ShowDialog()
    End Sub

    Private Sub ЧерныйСписокToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ЧерныйСписокToolStripMenuItem.Click
        ЧерныйСписок.ShowDialog()
    End Sub

    Private Sub ToolStrip_ItemClicked(sender As Object, e As ToolStripItemClickedEventArgs) Handles ToolStrip.ItemClicked

    End Sub
    Public Async Function Провер() As Task(Of DataTable)
        Dim DateEx As String = Format(Now.Date, "MM\/dd\/yyyy")
        Dim strsql As String = "SELECT * FROM ПереговорыКлиент WHERE ДатаНапоминания=#" & DateEx & "# and Экспедитор='" & Экспедитор & "'"
        Dim ds As DataTable = Selects(strsql)
        Return ds
    End Function
    Private Async Sub Getst()
        Dim df As DataTable = Await Провер()
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
End Class
