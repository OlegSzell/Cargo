<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class MDIParent1
    Inherits System.Windows.Forms.Form

    'Форма переопределяет dispose для очистки списка компонентов.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub


    'Является обязательной для конструктора форм Windows Forms
    Private components As System.ComponentModel.IContainer

    'Примечание: следующая процедура является обязательной для конструктора форм Windows Forms
    'Для ее изменения используйте конструктор форм Windows Form.  
    'Не изменяйте ее в редакторе исходного кода.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(MDIParent1))
        Me.MenuStrip = New System.Windows.Forms.MenuStrip()
        Me.ГрузыToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ДобавитьToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem()
        Me.ВыборкаToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem()
        Me.ПеревозчикиToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ДобавитьToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ПолныйToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.БыстроToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.НайтиToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.КлиентToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem()
        Me.ПереговорыToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.НапоминаниеToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.РейсToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.СоздатьToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ДоговораToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.КлиентToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ПеревозчикToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ПечатьToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.РейсToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem()
        Me.ДоговораToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem()
        Me.ЧерныйСписокToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStrip = New System.Windows.Forms.ToolStrip()
        Me.NewToolStripButton = New System.Windows.Forms.ToolStripButton()
        Me.OpenToolStripButton = New System.Windows.Forms.ToolStripButton()
        Me.SaveToolStripButton = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripSeparator1 = New System.Windows.Forms.ToolStripSeparator()
        Me.PrintToolStripButton = New System.Windows.Forms.ToolStripButton()
        Me.PrintPreviewToolStripButton = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripSeparator2 = New System.Windows.Forms.ToolStripSeparator()
        Me.HelpToolStripButton = New System.Windows.Forms.ToolStripButton()
        Me.StatusStrip = New System.Windows.Forms.StatusStrip()
        Me.ToolStripStatusLabel = New System.Windows.Forms.ToolStripStatusLabel()
        Me.ToolTip = New System.Windows.Forms.ToolTip(Me.components)
        Me.СвободныйТранспортToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.MenuStrip.SuspendLayout()
        Me.ToolStrip.SuspendLayout()
        Me.StatusStrip.SuspendLayout()
        Me.SuspendLayout()
        '
        'MenuStrip
        '
        Me.MenuStrip.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ГрузыToolStripMenuItem, Me.ПеревозчикиToolStripMenuItem, Me.КлиентToolStripMenuItem1, Me.РейсToolStripMenuItem, Me.ЧерныйСписокToolStripMenuItem})
        Me.MenuStrip.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip.Name = "MenuStrip"
        Me.MenuStrip.Padding = New System.Windows.Forms.Padding(8, 3, 0, 3)
        Me.MenuStrip.Size = New System.Drawing.Size(1395, 25)
        Me.MenuStrip.TabIndex = 5
        Me.MenuStrip.Text = "MenuStrip"
        '
        'ГрузыToolStripMenuItem
        '
        Me.ГрузыToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ДобавитьToolStripMenuItem1, Me.ВыборкаToolStripMenuItem1})
        Me.ГрузыToolStripMenuItem.Name = "ГрузыToolStripMenuItem"
        Me.ГрузыToolStripMenuItem.Size = New System.Drawing.Size(52, 19)
        Me.ГрузыToolStripMenuItem.Text = "Грузы"
        '
        'ДобавитьToolStripMenuItem1
        '
        Me.ДобавитьToolStripMenuItem1.Name = "ДобавитьToolStripMenuItem1"
        Me.ДобавитьToolStripMenuItem1.Size = New System.Drawing.Size(126, 22)
        Me.ДобавитьToolStripMenuItem1.Text = "Добавить"
        '
        'ВыборкаToolStripMenuItem1
        '
        Me.ВыборкаToolStripMenuItem1.Name = "ВыборкаToolStripMenuItem1"
        Me.ВыборкаToolStripMenuItem1.Size = New System.Drawing.Size(126, 22)
        Me.ВыборкаToolStripMenuItem1.Text = "Выборка"
        '
        'ПеревозчикиToolStripMenuItem
        '
        Me.ПеревозчикиToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ДобавитьToolStripMenuItem, Me.НайтиToolStripMenuItem})
        Me.ПеревозчикиToolStripMenuItem.Name = "ПеревозчикиToolStripMenuItem"
        Me.ПеревозчикиToolStripMenuItem.Size = New System.Drawing.Size(92, 19)
        Me.ПеревозчикиToolStripMenuItem.Text = "Перевозчики"
        '
        'ДобавитьToolStripMenuItem
        '
        Me.ДобавитьToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ПолныйToolStripMenuItem, Me.БыстроToolStripMenuItem})
        Me.ДобавитьToolStripMenuItem.Name = "ДобавитьToolStripMenuItem"
        Me.ДобавитьToolStripMenuItem.Size = New System.Drawing.Size(126, 22)
        Me.ДобавитьToolStripMenuItem.Text = "Добавить"
        '
        'ПолныйToolStripMenuItem
        '
        Me.ПолныйToolStripMenuItem.Name = "ПолныйToolStripMenuItem"
        Me.ПолныйToolStripMenuItem.Size = New System.Drawing.Size(120, 22)
        Me.ПолныйToolStripMenuItem.Text = "Полный"
        '
        'БыстроToolStripMenuItem
        '
        Me.БыстроToolStripMenuItem.Name = "БыстроToolStripMenuItem"
        Me.БыстроToolStripMenuItem.Size = New System.Drawing.Size(120, 22)
        Me.БыстроToolStripMenuItem.Text = "Быстро"
        '
        'НайтиToolStripMenuItem
        '
        Me.НайтиToolStripMenuItem.Name = "НайтиToolStripMenuItem"
        Me.НайтиToolStripMenuItem.Size = New System.Drawing.Size(126, 22)
        Me.НайтиToolStripMenuItem.Text = "Найти"
        '
        'КлиентToolStripMenuItem1
        '
        Me.КлиентToolStripMenuItem1.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ПереговорыToolStripMenuItem, Me.НапоминаниеToolStripMenuItem, Me.СвободныйТранспортToolStripMenuItem})
        Me.КлиентToolStripMenuItem1.Name = "КлиентToolStripMenuItem1"
        Me.КлиентToolStripMenuItem1.Size = New System.Drawing.Size(88, 19)
        Me.КлиентToolStripMenuItem1.Text = "Переговоры"
        '
        'ПереговорыToolStripMenuItem
        '
        Me.ПереговорыToolStripMenuItem.Name = "ПереговорыToolStripMenuItem"
        Me.ПереговорыToolStripMenuItem.Size = New System.Drawing.Size(180, 22)
        Me.ПереговорыToolStripMenuItem.Text = "Клиент"
        '
        'НапоминаниеToolStripMenuItem
        '
        Me.НапоминаниеToolStripMenuItem.Name = "НапоминаниеToolStripMenuItem"
        Me.НапоминаниеToolStripMenuItem.Size = New System.Drawing.Size(180, 22)
        Me.НапоминаниеToolStripMenuItem.Text = "Напоминание"
        '
        'РейсToolStripMenuItem
        '
        Me.РейсToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.СоздатьToolStripMenuItem, Me.ДоговораToolStripMenuItem, Me.ПечатьToolStripMenuItem})
        Me.РейсToolStripMenuItem.Name = "РейсToolStripMenuItem"
        Me.РейсToolStripMenuItem.Size = New System.Drawing.Size(45, 19)
        Me.РейсToolStripMenuItem.Text = "Рейс"
        '
        'СоздатьToolStripMenuItem
        '
        Me.СоздатьToolStripMenuItem.Name = "СоздатьToolStripMenuItem"
        Me.СоздатьToolStripMenuItem.Size = New System.Drawing.Size(176, 22)
        Me.СоздатьToolStripMenuItem.Text = "Создать/Изменить"
        '
        'ДоговораToolStripMenuItem
        '
        Me.ДоговораToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.КлиентToolStripMenuItem, Me.ПеревозчикToolStripMenuItem})
        Me.ДоговораToolStripMenuItem.Name = "ДоговораToolStripMenuItem"
        Me.ДоговораToolStripMenuItem.Size = New System.Drawing.Size(176, 22)
        Me.ДоговораToolStripMenuItem.Text = "Новый"
        '
        'КлиентToolStripMenuItem
        '
        Me.КлиентToolStripMenuItem.Name = "КлиентToolStripMenuItem"
        Me.КлиентToolStripMenuItem.Size = New System.Drawing.Size(140, 22)
        Me.КлиентToolStripMenuItem.Text = "Клиент"
        '
        'ПеревозчикToolStripMenuItem
        '
        Me.ПеревозчикToolStripMenuItem.Name = "ПеревозчикToolStripMenuItem"
        Me.ПеревозчикToolStripMenuItem.Size = New System.Drawing.Size(140, 22)
        Me.ПеревозчикToolStripMenuItem.Text = "Перевозчик"
        '
        'ПечатьToolStripMenuItem
        '
        Me.ПечатьToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.РейсToolStripMenuItem1, Me.ДоговораToolStripMenuItem1})
        Me.ПечатьToolStripMenuItem.Name = "ПечатьToolStripMenuItem"
        Me.ПечатьToolStripMenuItem.Size = New System.Drawing.Size(176, 22)
        Me.ПечатьToolStripMenuItem.Text = "Печать"
        '
        'РейсToolStripMenuItem1
        '
        Me.РейсToolStripMenuItem1.Name = "РейсToolStripMenuItem1"
        Me.РейсToolStripMenuItem1.Size = New System.Drawing.Size(137, 22)
        Me.РейсToolStripMenuItem1.Text = "Рейс"
        '
        'ДоговораToolStripMenuItem1
        '
        Me.ДоговораToolStripMenuItem1.Name = "ДоговораToolStripMenuItem1"
        Me.ДоговораToolStripMenuItem1.Size = New System.Drawing.Size(137, 22)
        Me.ДоговораToolStripMenuItem1.Text = "Документы"
        '
        'ЧерныйСписокToolStripMenuItem
        '
        Me.ЧерныйСписокToolStripMenuItem.Name = "ЧерныйСписокToolStripMenuItem"
        Me.ЧерныйСписокToolStripMenuItem.Size = New System.Drawing.Size(105, 19)
        Me.ЧерныйСписокToolStripMenuItem.Text = "Черный список"
        '
        'ToolStrip
        '
        Me.ToolStrip.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.NewToolStripButton, Me.OpenToolStripButton, Me.SaveToolStripButton, Me.ToolStripSeparator1, Me.PrintToolStripButton, Me.PrintPreviewToolStripButton, Me.ToolStripSeparator2, Me.HelpToolStripButton})
        Me.ToolStrip.Location = New System.Drawing.Point(0, 25)
        Me.ToolStrip.Name = "ToolStrip"
        Me.ToolStrip.Size = New System.Drawing.Size(1395, 25)
        Me.ToolStrip.TabIndex = 6
        Me.ToolStrip.Text = "ToolStrip"
        '
        'NewToolStripButton
        '
        Me.NewToolStripButton.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.NewToolStripButton.Image = CType(resources.GetObject("NewToolStripButton.Image"), System.Drawing.Image)
        Me.NewToolStripButton.ImageTransparentColor = System.Drawing.Color.Black
        Me.NewToolStripButton.Name = "NewToolStripButton"
        Me.NewToolStripButton.Size = New System.Drawing.Size(23, 22)
        Me.NewToolStripButton.Text = "Создать"
        '
        'OpenToolStripButton
        '
        Me.OpenToolStripButton.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.OpenToolStripButton.Image = CType(resources.GetObject("OpenToolStripButton.Image"), System.Drawing.Image)
        Me.OpenToolStripButton.ImageTransparentColor = System.Drawing.Color.Black
        Me.OpenToolStripButton.Name = "OpenToolStripButton"
        Me.OpenToolStripButton.Size = New System.Drawing.Size(23, 22)
        Me.OpenToolStripButton.Text = "Открыть"
        '
        'SaveToolStripButton
        '
        Me.SaveToolStripButton.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.SaveToolStripButton.Image = CType(resources.GetObject("SaveToolStripButton.Image"), System.Drawing.Image)
        Me.SaveToolStripButton.ImageTransparentColor = System.Drawing.Color.Black
        Me.SaveToolStripButton.Name = "SaveToolStripButton"
        Me.SaveToolStripButton.Size = New System.Drawing.Size(23, 22)
        Me.SaveToolStripButton.Text = "Сохранить"
        '
        'ToolStripSeparator1
        '
        Me.ToolStripSeparator1.Name = "ToolStripSeparator1"
        Me.ToolStripSeparator1.Size = New System.Drawing.Size(6, 25)
        '
        'PrintToolStripButton
        '
        Me.PrintToolStripButton.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.PrintToolStripButton.Image = CType(resources.GetObject("PrintToolStripButton.Image"), System.Drawing.Image)
        Me.PrintToolStripButton.ImageTransparentColor = System.Drawing.Color.Black
        Me.PrintToolStripButton.Name = "PrintToolStripButton"
        Me.PrintToolStripButton.Size = New System.Drawing.Size(23, 22)
        Me.PrintToolStripButton.Text = "Печать"
        '
        'PrintPreviewToolStripButton
        '
        Me.PrintPreviewToolStripButton.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.PrintPreviewToolStripButton.Image = CType(resources.GetObject("PrintPreviewToolStripButton.Image"), System.Drawing.Image)
        Me.PrintPreviewToolStripButton.ImageTransparentColor = System.Drawing.Color.Black
        Me.PrintPreviewToolStripButton.Name = "PrintPreviewToolStripButton"
        Me.PrintPreviewToolStripButton.Size = New System.Drawing.Size(23, 22)
        Me.PrintPreviewToolStripButton.Text = "Предварительный просмотр"
        '
        'ToolStripSeparator2
        '
        Me.ToolStripSeparator2.Name = "ToolStripSeparator2"
        Me.ToolStripSeparator2.Size = New System.Drawing.Size(6, 25)
        '
        'HelpToolStripButton
        '
        Me.HelpToolStripButton.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.HelpToolStripButton.Image = CType(resources.GetObject("HelpToolStripButton.Image"), System.Drawing.Image)
        Me.HelpToolStripButton.ImageTransparentColor = System.Drawing.Color.Black
        Me.HelpToolStripButton.Name = "HelpToolStripButton"
        Me.HelpToolStripButton.Size = New System.Drawing.Size(23, 22)
        Me.HelpToolStripButton.Text = "Справка"
        '
        'StatusStrip
        '
        Me.StatusStrip.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripStatusLabel})
        Me.StatusStrip.Location = New System.Drawing.Point(0, 692)
        Me.StatusStrip.Name = "StatusStrip"
        Me.StatusStrip.Padding = New System.Windows.Forms.Padding(1, 0, 19, 0)
        Me.StatusStrip.Size = New System.Drawing.Size(1395, 22)
        Me.StatusStrip.TabIndex = 7
        Me.StatusStrip.Text = "StatusStrip"
        '
        'ToolStripStatusLabel
        '
        Me.ToolStripStatusLabel.Name = "ToolStripStatusLabel"
        Me.ToolStripStatusLabel.Size = New System.Drawing.Size(66, 17)
        Me.ToolStripStatusLabel.Text = "Состояние"
        '
        'СвободныйТранспортToolStripMenuItem
        '
        Me.СвободныйТранспортToolStripMenuItem.Name = "СвободныйТранспортToolStripMenuItem"
        Me.СвободныйТранспортToolStripMenuItem.Size = New System.Drawing.Size(198, 22)
        Me.СвободныйТранспортToolStripMenuItem.Text = "Свободный транспорт"
        '
        'MDIParent1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 17.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1395, 714)
        Me.Controls.Add(Me.ToolStrip)
        Me.Controls.Add(Me.MenuStrip)
        Me.Controls.Add(Me.StatusStrip)
        Me.Font = New System.Drawing.Font("Times New Roman", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.IsMdiContainer = True
        Me.MainMenuStrip = Me.MenuStrip
        Me.Margin = New System.Windows.Forms.Padding(4)
        Me.Name = "MDIParent1"
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Show
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Главная"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.MenuStrip.ResumeLayout(False)
        Me.MenuStrip.PerformLayout()
        Me.ToolStrip.ResumeLayout(False)
        Me.ToolStrip.PerformLayout()
        Me.StatusStrip.ResumeLayout(False)
        Me.StatusStrip.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents HelpToolStripButton As System.Windows.Forms.ToolStripButton
    Friend WithEvents ToolStripSeparator2 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents PrintPreviewToolStripButton As System.Windows.Forms.ToolStripButton
    Friend WithEvents ToolTip As System.Windows.Forms.ToolTip
    Friend WithEvents ToolStripStatusLabel As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents StatusStrip As System.Windows.Forms.StatusStrip
    Friend WithEvents PrintToolStripButton As System.Windows.Forms.ToolStripButton
    Friend WithEvents NewToolStripButton As System.Windows.Forms.ToolStripButton
    Friend WithEvents ToolStrip As System.Windows.Forms.ToolStrip
    Friend WithEvents OpenToolStripButton As System.Windows.Forms.ToolStripButton
    Friend WithEvents SaveToolStripButton As System.Windows.Forms.ToolStripButton
    Friend WithEvents ToolStripSeparator1 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents MenuStrip As System.Windows.Forms.MenuStrip
    Friend WithEvents ГрузыToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ДобавитьToolStripMenuItem1 As ToolStripMenuItem
    Friend WithEvents ВыборкаToolStripMenuItem1 As ToolStripMenuItem
    Friend WithEvents ПеревозчикиToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ДобавитьToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents НайтиToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ПолныйToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents БыстроToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents РейсToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents СоздатьToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ДоговораToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents КлиентToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ПеревозчикToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ПечатьToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents РейсToolStripMenuItem1 As ToolStripMenuItem
    Friend WithEvents ДоговораToolStripMenuItem1 As ToolStripMenuItem
    Friend WithEvents КлиентToolStripMenuItem1 As ToolStripMenuItem
    Friend WithEvents ПереговорыToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ЧерныйСписокToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents НапоминаниеToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents СвободныйТранспортToolStripMenuItem As ToolStripMenuItem
End Class
