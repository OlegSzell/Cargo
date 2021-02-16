<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Календарь
    Inherits System.Windows.Forms.Form

    'Форма переопределяет dispose для очистки списка компонентов.
    <System.Diagnostics.DebuggerNonUserCode()> _
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
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.Calendar1 = New System.Windows.Forms.MonthCalendar()
        Me.Grid2 = New System.Windows.Forms.DataGridView()
        Me.ContextMenuStrip1 = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.ВыполненоToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ContextMenuStrip2 = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.РезультатЗвонкаToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ИзменитьРезультатToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.Grid1 = New Грузы.DoubleBuferGrid()
        CType(Me.Grid2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.ContextMenuStrip1.SuspendLayout()
        Me.ContextMenuStrip2.SuspendLayout()
        CType(Me.Grid1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Calendar1
        '
        Me.Calendar1.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.Calendar1.CalendarDimensions = New System.Drawing.Size(1, 2)
        Me.Calendar1.Location = New System.Drawing.Point(8, 3)
        Me.Calendar1.Margin = New System.Windows.Forms.Padding(12)
        Me.Calendar1.MaximumSize = New System.Drawing.Size(400, 415)
        Me.Calendar1.Name = "Calendar1"
        Me.Calendar1.ShowWeekNumbers = True
        Me.Calendar1.TabIndex = 1
        '
        'Grid2
        '
        Me.Grid2.AllowUserToAddRows = False
        Me.Grid2.AllowUserToDeleteRows = False
        Me.Grid2.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Grid2.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill
        Me.Grid2.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.Grid2.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.Grid2.Location = New System.Drawing.Point(8, 311)
        Me.Grid2.Margin = New System.Windows.Forms.Padding(4)
        Me.Grid2.Name = "Grid2"
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.Grid2.RowsDefaultCellStyle = DataGridViewCellStyle1
        Me.Grid2.Size = New System.Drawing.Size(1516, 300)
        Me.Grid2.TabIndex = 3
        '
        'ContextMenuStrip1
        '
        Me.ContextMenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ВыполненоToolStripMenuItem})
        Me.ContextMenuStrip1.Name = "ContextMenuStrip1"
        Me.ContextMenuStrip1.Size = New System.Drawing.Size(139, 26)
        '
        'ВыполненоToolStripMenuItem
        '
        Me.ВыполненоToolStripMenuItem.Name = "ВыполненоToolStripMenuItem"
        Me.ВыполненоToolStripMenuItem.Size = New System.Drawing.Size(138, 22)
        Me.ВыполненоToolStripMenuItem.Text = "Выполнено"
        '
        'ContextMenuStrip2
        '
        Me.ContextMenuStrip2.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.РезультатЗвонкаToolStripMenuItem, Me.ИзменитьРезультатToolStripMenuItem})
        Me.ContextMenuStrip2.Name = "ContextMenuStrip2"
        Me.ContextMenuStrip2.Size = New System.Drawing.Size(185, 48)
        '
        'РезультатЗвонкаToolStripMenuItem
        '
        Me.РезультатЗвонкаToolStripMenuItem.Name = "РезультатЗвонкаToolStripMenuItem"
        Me.РезультатЗвонкаToolStripMenuItem.Size = New System.Drawing.Size(184, 22)
        Me.РезультатЗвонкаToolStripMenuItem.Text = "Результат звонка"
        '
        'ИзменитьРезультатToolStripMenuItem
        '
        Me.ИзменитьРезультатToolStripMenuItem.Name = "ИзменитьРезультатToolStripMenuItem"
        Me.ИзменитьРезультатToolStripMenuItem.Size = New System.Drawing.Size(184, 22)
        Me.ИзменитьРезультатToolStripMenuItem.Text = "Изменить результат"
        '
        'Grid1
        '
        Me.Grid1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Grid1.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill
        Me.Grid1.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.Grid1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.Grid1.Location = New System.Drawing.Point(197, 3)
        Me.Grid1.Name = "Grid1"
        DataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.Grid1.RowsDefaultCellStyle = DataGridViewCellStyle2
        Me.Grid1.Size = New System.Drawing.Size(1327, 301)
        Me.Grid1.TabIndex = 4
        '
        'Календарь
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 18.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1537, 624)
        Me.Controls.Add(Me.Grid1)
        Me.Controls.Add(Me.Grid2)
        Me.Controls.Add(Me.Calendar1)
        Me.DoubleBuffered = True
        Me.Font = New System.Drawing.Font("Calibri", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.Margin = New System.Windows.Forms.Padding(4)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "Календарь"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Календарь"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        CType(Me.Grid2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ContextMenuStrip1.ResumeLayout(False)
        Me.ContextMenuStrip2.ResumeLayout(False)
        CType(Me.Grid1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents Calendar1 As MonthCalendar
    Friend WithEvents Grid2 As DataGridView
    Friend WithEvents ContextMenuStrip1 As ContextMenuStrip
    Friend WithEvents ВыполненоToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ContextMenuStrip2 As ContextMenuStrip
    Friend WithEvents РезультатЗвонкаToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ИзменитьРезультатToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents Grid1 As DoubleBuferGrid
End Class
