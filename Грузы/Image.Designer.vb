<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class ImageForm
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(ImageForm))
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.Button3 = New System.Windows.Forms.Button()
        Me.Button4 = New System.Windows.Forms.Button()
        Me.Button5 = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Button6 = New System.Windows.Forms.Button()
        Me.TabControl1 = New System.Windows.Forms.TabControl()
        Me.Фото_1 = New System.Windows.Forms.TabPage()
        Me.Окно2 = New System.Windows.Forms.TabPage()
        Me.RichTextBox1 = New System.Windows.Forms.RichTextBox()
        Me.RichTextBox2 = New System.Windows.Forms.RichTextBox()
        Me.TabControl1.SuspendLayout()
        Me.Фото_1.SuspendLayout()
        Me.Окно2.SuspendLayout()
        Me.SuspendLayout()
        '
        'Button1
        '
        Me.Button1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Button1.Location = New System.Drawing.Point(1274, 754)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(127, 25)
        Me.Button1.TabIndex = 2
        Me.Button1.Text = "Записать"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Button2
        '
        Me.Button2.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.Button2.Enabled = False
        Me.Button2.Location = New System.Drawing.Point(115, 12)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(97, 25)
        Me.Button2.TabIndex = 3
        Me.Button2.Text = "Прочитать"
        Me.Button2.UseVisualStyleBackColor = False
        '
        'Button3
        '
        Me.Button3.BackColor = System.Drawing.Color.FromArgb(CType(CType(128, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.Button3.Location = New System.Drawing.Point(12, 12)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(97, 25)
        Me.Button3.TabIndex = 4
        Me.Button3.Text = "Записать"
        Me.Button3.UseVisualStyleBackColor = False
        '
        'Button4
        '
        Me.Button4.BackColor = System.Drawing.Color.FromArgb(CType(CType(128, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.Button4.Location = New System.Drawing.Point(514, 12)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(97, 25)
        Me.Button4.TabIndex = 5
        Me.Button4.Text = "Очистить"
        Me.Button4.UseVisualStyleBackColor = False
        '
        'Button5
        '
        Me.Button5.BackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.Button5.Location = New System.Drawing.Point(411, 12)
        Me.Button5.Name = "Button5"
        Me.Button5.Size = New System.Drawing.Size(97, 25)
        Me.Button5.TabIndex = 6
        Me.Button5.Text = "Вставить"
        Me.Button5.UseVisualStyleBackColor = False
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(632, 17)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(66, 14)
        Me.Label1.TabIndex = 7
        Me.Label1.Text = "Дата фото:"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Calibri", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.Label2.Location = New System.Drawing.Point(704, 17)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(39, 14)
        Me.Label2.TabIndex = 8
        Me.Label2.Text = "Label2"
        '
        'Button6
        '
        Me.Button6.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.Button6.Location = New System.Drawing.Point(218, 12)
        Me.Button6.Name = "Button6"
        Me.Button6.Size = New System.Drawing.Size(156, 25)
        Me.Button6.TabIndex = 9
        Me.Button6.Text = "Добавить к действующей"
        Me.Button6.UseVisualStyleBackColor = False
        '
        'TabControl1
        '
        Me.TabControl1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TabControl1.Controls.Add(Me.Фото_1)
        Me.TabControl1.Controls.Add(Me.Окно2)
        Me.TabControl1.Location = New System.Drawing.Point(12, 43)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(1364, 679)
        Me.TabControl1.TabIndex = 10
        '
        'Фото_1
        '
        Me.Фото_1.Controls.Add(Me.RichTextBox1)
        Me.Фото_1.ForeColor = System.Drawing.Color.Transparent
        Me.Фото_1.Location = New System.Drawing.Point(4, 23)
        Me.Фото_1.Name = "Фото_1"
        Me.Фото_1.Padding = New System.Windows.Forms.Padding(3)
        Me.Фото_1.Size = New System.Drawing.Size(1356, 652)
        Me.Фото_1.TabIndex = 0
        Me.Фото_1.Text = "Фото 1"
        Me.Фото_1.UseVisualStyleBackColor = True
        '
        'Окно2
        '
        Me.Окно2.Controls.Add(Me.RichTextBox2)
        Me.Окно2.Location = New System.Drawing.Point(4, 23)
        Me.Окно2.Name = "Окно2"
        Me.Окно2.Padding = New System.Windows.Forms.Padding(3)
        Me.Окно2.Size = New System.Drawing.Size(1356, 652)
        Me.Окно2.TabIndex = 1
        Me.Окно2.Text = "Фото 2"
        Me.Окно2.UseVisualStyleBackColor = True
        '
        'RichTextBox1
        '
        Me.RichTextBox1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.RichTextBox1.BackColor = System.Drawing.SystemColors.InactiveBorder
        Me.RichTextBox1.Location = New System.Drawing.Point(16, 17)
        Me.RichTextBox1.Name = "RichTextBox1"
        Me.RichTextBox1.Size = New System.Drawing.Size(1324, 619)
        Me.RichTextBox1.TabIndex = 1
        Me.RichTextBox1.Text = ""
        '
        'RichTextBox2
        '
        Me.RichTextBox2.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.RichTextBox2.BackColor = System.Drawing.SystemColors.InactiveBorder
        Me.RichTextBox2.Location = New System.Drawing.Point(16, 17)
        Me.RichTextBox2.Name = "RichTextBox2"
        Me.RichTextBox2.Size = New System.Drawing.Size(1324, 619)
        Me.RichTextBox2.TabIndex = 2
        Me.RichTextBox2.Text = ""
        '
        'ImageForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 14.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.DarkKhaki
        Me.ClientSize = New System.Drawing.Size(1388, 734)
        Me.Controls.Add(Me.TabControl1)
        Me.Controls.Add(Me.Button6)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Button5)
        Me.Controls.Add(Me.Button4)
        Me.Controls.Add(Me.Button3)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.Button1)
        Me.DoubleBuffered = True
        Me.Font = New System.Drawing.Font("Calibri", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "ImageForm"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Изображение"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.TabControl1.ResumeLayout(False)
        Me.Фото_1.ResumeLayout(False)
        Me.Окно2.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Button1 As Button
    Friend WithEvents Button2 As Button
    Friend WithEvents Button3 As Button
    Friend WithEvents Button4 As Button
    Friend WithEvents Button5 As Button
    Friend WithEvents Label1 As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents Button6 As Button
    Friend WithEvents TabControl1 As TabControl
    Friend WithEvents Фото_1 As TabPage
    Friend WithEvents RichTextBox1 As RichTextBox
    Friend WithEvents Окно2 As TabPage
    Friend WithEvents RichTextBox2 As RichTextBox
End Class
