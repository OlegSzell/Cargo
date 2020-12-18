<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class НовыйПеревоз
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(НовыйПеревоз))
        Me.Button2 = New System.Windows.Forms.Button()
        Me.Label18 = New System.Windows.Forms.Label()
        Me.ComboBox3 = New System.Windows.Forms.ComboBox()
        Me.MaskedTextBox1 = New System.Windows.Forms.MaskedTextBox()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.ComboBox2 = New System.Windows.Forms.ComboBox()
        Me.CheckBox4 = New System.Windows.Forms.CheckBox()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.CheckBox2 = New System.Windows.Forms.CheckBox()
        Me.CheckBox1 = New System.Windows.Forms.CheckBox()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Label16 = New System.Windows.Forms.Label()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.Label17 = New System.Windows.Forms.Label()
        Me.RichTextBox14 = New System.Windows.Forms.RichTextBox()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.RichTextBox10 = New System.Windows.Forms.RichTextBox()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.RichTextBox11 = New System.Windows.Forms.RichTextBox()
        Me.Label14 = New System.Windows.Forms.Label()
        Me.RichTextBox12 = New System.Windows.Forms.RichTextBox()
        Me.Label15 = New System.Windows.Forms.Label()
        Me.RichTextBox13 = New System.Windows.Forms.RichTextBox()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.RichTextBox7 = New System.Windows.Forms.RichTextBox()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.RichTextBox8 = New System.Windows.Forms.RichTextBox()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.RichTextBox9 = New System.Windows.Forms.RichTextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.RichTextBox4 = New System.Windows.Forms.RichTextBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.RichTextBox5 = New System.Windows.Forms.RichTextBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.RichTextBox6 = New System.Windows.Forms.RichTextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.RichTextBox3 = New System.Windows.Forms.RichTextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.RichTextBox2 = New System.Windows.Forms.RichTextBox()
        Me.ComboBox1 = New System.Windows.Forms.ComboBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.RichTextBox1 = New System.Windows.Forms.RichTextBox()
        Me.ListBox1 = New System.Windows.Forms.ListBox()
        Me.ComboBox4 = New System.Windows.Forms.ComboBox()
        Me.CheckBox3 = New System.Windows.Forms.CheckBox()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Button2
        '
        Me.Button2.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.Button2.Font = New System.Drawing.Font("Times New Roman", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.Button2.Location = New System.Drawing.Point(1068, 6)
        Me.Button2.Margin = New System.Windows.Forms.Padding(4)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(84, 29)
        Me.Button2.TabIndex = 57
        Me.Button2.Text = "Удалить"
        Me.Button2.UseVisualStyleBackColor = False
        '
        'Label18
        '
        Me.Label18.AutoSize = True
        Me.Label18.Location = New System.Drawing.Point(374, 6)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(56, 18)
        Me.Label18.TabIndex = 56
        Me.Label18.Text = "Кол.экз"
        '
        'ComboBox3
        '
        Me.ComboBox3.Font = New System.Drawing.Font("Times New Roman", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.ComboBox3.FormattingEnabled = True
        Me.ComboBox3.Items.AddRange(New Object() {"1", "2"})
        Me.ComboBox3.Location = New System.Drawing.Point(434, 3)
        Me.ComboBox3.Name = "ComboBox3"
        Me.ComboBox3.Size = New System.Drawing.Size(53, 23)
        Me.ComboBox3.TabIndex = 55
        '
        'MaskedTextBox1
        '
        Me.MaskedTextBox1.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.MaskedTextBox1.Location = New System.Drawing.Point(876, 22)
        Me.MaskedTextBox1.Mask = "00/00/0000"
        Me.MaskedTextBox1.Name = "MaskedTextBox1"
        Me.MaskedTextBox1.Size = New System.Drawing.Size(88, 23)
        Me.MaskedTextBox1.TabIndex = 54
        Me.MaskedTextBox1.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.MaskedTextBox1.ValidatingType = GetType(Date)
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(704, 1)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(97, 18)
        Me.Label11.TabIndex = 53
        Me.Label11.Text = "Наш расч.счет"
        '
        'ComboBox2
        '
        Me.ComboBox2.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.ComboBox2.FormattingEnabled = True
        Me.ComboBox2.Items.AddRange(New Object() {"Евро", "Доллар", "Российский рубль"})
        Me.ComboBox2.Location = New System.Drawing.Point(707, 22)
        Me.ComboBox2.Name = "ComboBox2"
        Me.ComboBox2.Size = New System.Drawing.Size(150, 23)
        Me.ComboBox2.TabIndex = 52
        '
        'CheckBox4
        '
        Me.CheckBox4.AutoSize = True
        Me.CheckBox4.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.CheckBox4.Location = New System.Drawing.Point(151, 2)
        Me.CheckBox4.Name = "CheckBox4"
        Me.CheckBox4.Size = New System.Drawing.Size(133, 19)
        Me.CheckBox4.TabIndex = 50
        Me.CheckBox4.Text = "Оформить договор"
        Me.CheckBox4.UseVisualStyleBackColor = True
        '
        'GroupBox2
        '
        Me.GroupBox2.BackColor = System.Drawing.SystemColors.ActiveCaption
        Me.GroupBox2.Controls.Add(Me.CheckBox2)
        Me.GroupBox2.Location = New System.Drawing.Point(493, 3)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(192, 42)
        Me.GroupBox2.TabIndex = 51
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Перевозчик"
        '
        'CheckBox2
        '
        Me.CheckBox2.AutoSize = True
        Me.CheckBox2.Location = New System.Drawing.Point(85, 16)
        Me.CheckBox2.Name = "CheckBox2"
        Me.CheckBox2.Size = New System.Drawing.Size(102, 22)
        Me.CheckBox2.TabIndex = 36
        Me.CheckBox2.Text = "Экспедитор"
        Me.CheckBox2.UseVisualStyleBackColor = True
        '
        'CheckBox1
        '
        Me.CheckBox1.AutoSize = True
        Me.CheckBox1.Location = New System.Drawing.Point(980, 10)
        Me.CheckBox1.Name = "CheckBox1"
        Me.CheckBox1.Size = New System.Drawing.Size(78, 22)
        Me.CheckBox1.TabIndex = 49
        Me.CheckBox1.Text = "Очистка"
        Me.CheckBox1.UseVisualStyleBackColor = True
        '
        'Button1
        '
        Me.Button1.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.Button1.Font = New System.Drawing.Font("Times New Roman", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.Button1.Location = New System.Drawing.Point(13, 6)
        Me.Button1.Margin = New System.Windows.Forms.Padding(4)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(112, 33)
        Me.Button1.TabIndex = 48
        Me.Button1.Text = "Записать"
        Me.Button1.UseVisualStyleBackColor = False
        '
        'Label16
        '
        Me.Label16.AutoSize = True
        Me.Label16.Location = New System.Drawing.Point(873, 0)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(38, 18)
        Me.Label16.TabIndex = 47
        Me.Label16.Text = "Дата"
        '
        'GroupBox1
        '
        Me.GroupBox1.BackColor = System.Drawing.SystemColors.AppWorkspace
        Me.GroupBox1.Controls.Add(Me.Label17)
        Me.GroupBox1.Controls.Add(Me.RichTextBox14)
        Me.GroupBox1.Controls.Add(Me.Label12)
        Me.GroupBox1.Controls.Add(Me.RichTextBox10)
        Me.GroupBox1.Controls.Add(Me.Label13)
        Me.GroupBox1.Controls.Add(Me.RichTextBox11)
        Me.GroupBox1.Controls.Add(Me.Label14)
        Me.GroupBox1.Controls.Add(Me.RichTextBox12)
        Me.GroupBox1.Controls.Add(Me.Label15)
        Me.GroupBox1.Controls.Add(Me.RichTextBox13)
        Me.GroupBox1.Controls.Add(Me.Label8)
        Me.GroupBox1.Controls.Add(Me.RichTextBox7)
        Me.GroupBox1.Controls.Add(Me.Label9)
        Me.GroupBox1.Controls.Add(Me.RichTextBox8)
        Me.GroupBox1.Controls.Add(Me.Label10)
        Me.GroupBox1.Controls.Add(Me.RichTextBox9)
        Me.GroupBox1.Controls.Add(Me.Label5)
        Me.GroupBox1.Controls.Add(Me.RichTextBox4)
        Me.GroupBox1.Controls.Add(Me.Label6)
        Me.GroupBox1.Controls.Add(Me.RichTextBox5)
        Me.GroupBox1.Controls.Add(Me.Label7)
        Me.GroupBox1.Controls.Add(Me.RichTextBox6)
        Me.GroupBox1.Controls.Add(Me.Label4)
        Me.GroupBox1.Controls.Add(Me.RichTextBox3)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.RichTextBox2)
        Me.GroupBox1.Controls.Add(Me.ComboBox1)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Controls.Add(Me.RichTextBox1)
        Me.GroupBox1.Location = New System.Drawing.Point(12, 49)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(783, 682)
        Me.GroupBox1.TabIndex = 46
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Организация"
        '
        'Label17
        '
        Me.Label17.Location = New System.Drawing.Point(14, 494)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(170, 43)
        Me.Label17.TabIndex = 32
        Me.Label17.Text = "Должность ответс. лица в род.падеж *"
        '
        'RichTextBox14
        '
        Me.RichTextBox14.Location = New System.Drawing.Point(190, 501)
        Me.RichTextBox14.Name = "RichTextBox14"
        Me.RichTextBox14.Size = New System.Drawing.Size(580, 37)
        Me.RichTextBox14.TabIndex = 31
        Me.RichTextBox14.Text = ""
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Location = New System.Drawing.Point(11, 634)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(203, 18)
        Me.Label12.TabIndex = 30
        Me.Label12.Text = "На основании чего действует *"
        '
        'RichTextBox10
        '
        Me.RichTextBox10.Location = New System.Drawing.Point(220, 631)
        Me.RichTextBox10.Name = "RichTextBox10"
        Me.RichTextBox10.Size = New System.Drawing.Size(550, 37)
        Me.RichTextBox10.TabIndex = 29
        Me.RichTextBox10.Text = ""
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Location = New System.Drawing.Point(11, 591)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(135, 18)
        Me.Label13.TabIndex = 28
        Me.Label13.Text = "ФИО в Род.падеже*"
        '
        'RichTextBox11
        '
        Me.RichTextBox11.Location = New System.Drawing.Point(190, 588)
        Me.RichTextBox11.Name = "RichTextBox11"
        Me.RichTextBox11.Size = New System.Drawing.Size(580, 37)
        Me.RichTextBox11.TabIndex = 27
        Me.RichTextBox11.Text = ""
        '
        'Label14
        '
        Me.Label14.AutoSize = True
        Me.Label14.Location = New System.Drawing.Point(14, 547)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(130, 18)
        Me.Label14.TabIndex = 26
        Me.Label14.Text = "ФИО ответс. лица *"
        '
        'RichTextBox12
        '
        Me.RichTextBox12.Location = New System.Drawing.Point(190, 544)
        Me.RichTextBox12.Name = "RichTextBox12"
        Me.RichTextBox12.Size = New System.Drawing.Size(580, 37)
        Me.RichTextBox12.TabIndex = 25
        Me.RichTextBox12.Text = ""
        '
        'Label15
        '
        Me.Label15.Location = New System.Drawing.Point(14, 461)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(170, 23)
        Me.Label15.TabIndex = 24
        Me.Label15.Text = "Должность ответс. лица *"
        '
        'RichTextBox13
        '
        Me.RichTextBox13.Location = New System.Drawing.Point(190, 457)
        Me.RichTextBox13.Name = "RichTextBox13"
        Me.RichTextBox13.Size = New System.Drawing.Size(580, 37)
        Me.RichTextBox13.TabIndex = 23
        Me.RichTextBox13.Text = ""
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(11, 417)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(168, 18)
        Me.Label8.TabIndex = 19
        Me.Label8.Text = "Контакт. лицо, телефон *"
        '
        'RichTextBox7
        '
        Me.RichTextBox7.Location = New System.Drawing.Point(190, 414)
        Me.RichTextBox7.Name = "RichTextBox7"
        Me.RichTextBox7.Size = New System.Drawing.Size(580, 37)
        Me.RichTextBox7.TabIndex = 18
        Me.RichTextBox7.Text = ""
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(11, 374)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(167, 18)
        Me.Label9.TabIndex = 17
        Me.Label9.Text = "Банк, Адрес банка, БИК *"
        '
        'RichTextBox8
        '
        Me.RichTextBox8.Location = New System.Drawing.Point(190, 371)
        Me.RichTextBox8.Name = "RichTextBox8"
        Me.RichTextBox8.Size = New System.Drawing.Size(580, 37)
        Me.RichTextBox8.TabIndex = 16
        Me.RichTextBox8.Text = ""
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(14, 330)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(144, 18)
        Me.Label10.TabIndex = 15
        Me.Label10.Text = "Расчетный счет (RUB)"
        '
        'RichTextBox9
        '
        Me.RichTextBox9.Font = New System.Drawing.Font("Times New Roman", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.RichTextBox9.Location = New System.Drawing.Point(190, 327)
        Me.RichTextBox9.Name = "RichTextBox9"
        Me.RichTextBox9.Size = New System.Drawing.Size(580, 37)
        Me.RichTextBox9.TabIndex = 14
        Me.RichTextBox9.Text = ""
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(14, 287)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(143, 18)
        Me.Label5.TabIndex = 13
        Me.Label5.Text = "Расчетный счет (EUR)"
        '
        'RichTextBox4
        '
        Me.RichTextBox4.Font = New System.Drawing.Font("Times New Roman", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.RichTextBox4.Location = New System.Drawing.Point(190, 284)
        Me.RichTextBox4.Name = "RichTextBox4"
        Me.RichTextBox4.Size = New System.Drawing.Size(580, 37)
        Me.RichTextBox4.TabIndex = 12
        Me.RichTextBox4.Text = ""
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(14, 244)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(144, 18)
        Me.Label6.TabIndex = 11
        Me.Label6.Text = "Расчетный счет (USD)"
        '
        'RichTextBox5
        '
        Me.RichTextBox5.Font = New System.Drawing.Font("Times New Roman", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.RichTextBox5.Location = New System.Drawing.Point(190, 240)
        Me.RichTextBox5.Name = "RichTextBox5"
        Me.RichTextBox5.Size = New System.Drawing.Size(580, 37)
        Me.RichTextBox5.TabIndex = 10
        Me.RichTextBox5.Text = ""
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(14, 200)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(144, 18)
        Me.Label7.TabIndex = 9
        Me.Label7.Text = "Расчетный счет (BYN)"
        '
        'RichTextBox6
        '
        Me.RichTextBox6.Font = New System.Drawing.Font("Times New Roman", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.RichTextBox6.Location = New System.Drawing.Point(190, 199)
        Me.RichTextBox6.MaxLength = 28
        Me.RichTextBox6.Name = "RichTextBox6"
        Me.RichTextBox6.Size = New System.Drawing.Size(580, 37)
        Me.RichTextBox6.TabIndex = 8
        Me.RichTextBox6.Text = ""
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(14, 157)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(121, 18)
        Me.Label4.TabIndex = 7
        Me.Label4.Text = "Почтовый адрес *"
        '
        'RichTextBox3
        '
        Me.RichTextBox3.Location = New System.Drawing.Point(190, 154)
        Me.RichTextBox3.Name = "RichTextBox3"
        Me.RichTextBox3.Size = New System.Drawing.Size(580, 37)
        Me.RichTextBox3.TabIndex = 6
        Me.RichTextBox3.Text = ""
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(11, 113)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(174, 18)
        Me.Label3.TabIndex = 5
        Me.Label3.Text = "Адрес организации, УНП *"
        '
        'RichTextBox2
        '
        Me.RichTextBox2.Location = New System.Drawing.Point(190, 110)
        Me.RichTextBox2.Name = "RichTextBox2"
        Me.RichTextBox2.Size = New System.Drawing.Size(580, 37)
        Me.RichTextBox2.TabIndex = 4
        Me.RichTextBox2.Text = ""
        '
        'ComboBox1
        '
        Me.ComboBox1.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend
        Me.ComboBox1.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.CustomSource
        Me.ComboBox1.FormattingEnabled = True
        Me.ComboBox1.Location = New System.Drawing.Point(190, 24)
        Me.ComboBox1.Name = "ComboBox1"
        Me.ComboBox1.Size = New System.Drawing.Size(580, 26)
        Me.ComboBox1.TabIndex = 3
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(14, 33)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(157, 18)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "Форма собственности *"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(14, 70)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(162, 18)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Название организации *"
        '
        'RichTextBox1
        '
        Me.RichTextBox1.Location = New System.Drawing.Point(190, 66)
        Me.RichTextBox1.Name = "RichTextBox1"
        Me.RichTextBox1.Size = New System.Drawing.Size(580, 37)
        Me.RichTextBox1.TabIndex = 0
        Me.RichTextBox1.Text = ""
        '
        'ListBox1
        '
        Me.ListBox1.BackColor = System.Drawing.Color.Khaki
        Me.ListBox1.FormattingEnabled = True
        Me.ListBox1.ItemHeight = 18
        Me.ListBox1.Location = New System.Drawing.Point(801, 100)
        Me.ListBox1.Name = "ListBox1"
        Me.ListBox1.ScrollAlwaysVisible = True
        Me.ListBox1.Size = New System.Drawing.Size(354, 634)
        Me.ListBox1.TabIndex = 45
        '
        'ComboBox4
        '
        Me.ComboBox4.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend
        Me.ComboBox4.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems
        Me.ComboBox4.BackColor = System.Drawing.SystemColors.GradientInactiveCaption
        Me.ComboBox4.FormattingEnabled = True
        Me.ComboBox4.Location = New System.Drawing.Point(801, 67)
        Me.ComboBox4.Name = "ComboBox4"
        Me.ComboBox4.Size = New System.Drawing.Size(354, 26)
        Me.ComboBox4.TabIndex = 58
        '
        'CheckBox3
        '
        Me.CheckBox3.AutoSize = True
        Me.CheckBox3.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.CheckBox3.Location = New System.Drawing.Point(151, 27)
        Me.CheckBox3.Name = "CheckBox3"
        Me.CheckBox3.Size = New System.Drawing.Size(332, 19)
        Me.CheckBox3.TabIndex = 59
        Me.CheckBox3.Text = "Оформить новый договор, с действующим перевозом."
        Me.CheckBox3.UseVisualStyleBackColor = True
        '
        'НовыйПеревоз
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 18.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1169, 742)
        Me.Controls.Add(Me.CheckBox3)
        Me.Controls.Add(Me.ComboBox4)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.Label18)
        Me.Controls.Add(Me.ComboBox3)
        Me.Controls.Add(Me.MaskedTextBox1)
        Me.Controls.Add(Me.Label11)
        Me.Controls.Add(Me.ComboBox2)
        Me.Controls.Add(Me.CheckBox4)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.CheckBox1)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.Label16)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.ListBox1)
        Me.Font = New System.Drawing.Font("Calibri", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Margin = New System.Windows.Forms.Padding(4)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "НовыйПеревоз"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Новый перевозчик"
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Button2 As Button
    Friend WithEvents Label18 As Label
    Friend WithEvents ComboBox3 As ComboBox
    Friend WithEvents MaskedTextBox1 As MaskedTextBox
    Friend WithEvents Label11 As Label
    Friend WithEvents ComboBox2 As ComboBox
    Friend WithEvents CheckBox4 As CheckBox
    Friend WithEvents GroupBox2 As GroupBox
    Friend WithEvents CheckBox1 As CheckBox
    Friend WithEvents Button1 As Button
    Friend WithEvents Label16 As Label
    Friend WithEvents GroupBox1 As GroupBox
    Friend WithEvents Label17 As Label
    Friend WithEvents RichTextBox14 As RichTextBox
    Friend WithEvents Label12 As Label
    Friend WithEvents RichTextBox10 As RichTextBox
    Friend WithEvents Label13 As Label
    Friend WithEvents RichTextBox11 As RichTextBox
    Friend WithEvents Label14 As Label
    Friend WithEvents RichTextBox12 As RichTextBox
    Friend WithEvents Label15 As Label
    Friend WithEvents RichTextBox13 As RichTextBox
    Friend WithEvents Label8 As Label
    Friend WithEvents RichTextBox7 As RichTextBox
    Friend WithEvents Label9 As Label
    Friend WithEvents RichTextBox8 As RichTextBox
    Friend WithEvents Label10 As Label
    Friend WithEvents RichTextBox9 As RichTextBox
    Friend WithEvents Label5 As Label
    Friend WithEvents RichTextBox4 As RichTextBox
    Friend WithEvents Label6 As Label
    Friend WithEvents RichTextBox5 As RichTextBox
    Friend WithEvents Label7 As Label
    Friend WithEvents RichTextBox6 As RichTextBox
    Friend WithEvents Label4 As Label
    Friend WithEvents RichTextBox3 As RichTextBox
    Friend WithEvents Label3 As Label
    Friend WithEvents RichTextBox2 As RichTextBox
    Friend WithEvents ComboBox1 As ComboBox
    Friend WithEvents Label2 As Label
    Friend WithEvents Label1 As Label
    Friend WithEvents RichTextBox1 As RichTextBox
    Friend WithEvents ListBox1 As ListBox
    Friend WithEvents CheckBox2 As CheckBox
    Friend WithEvents ComboBox4 As ComboBox
    Friend WithEvents CheckBox3 As CheckBox
End Class
