Imports System.IO
Imports System.Runtime.CompilerServices
Imports System.Text
Imports AxSHDocVw

Public Class ImageForm

    Private ФормаЗапуска As Boolean = False
    Public Property Imagese As Byte() = Nothing
    Public Property Imagese2 As Byte() = Nothing
    Private IDp As Integer
    Private Read As Boolean = False

    Sub New()

        ' Этот вызов является обязательным для конструктора.
        InitializeComponent()

        ' Добавить код инициализации после вызова InitializeComponent().

    End Sub

    Sub New(ByVal _ФормаЗапуска As Boolean, Optional ID As Integer = 0, Optional _Read As Boolean = False)
        ФормаЗапуска = _ФормаЗапуска
        IDp = ID
        Read = _Read
        ' Этот вызов является обязательным для конструктора.
        InitializeComponent()

        ' Добавить код инициализации после вызова InitializeComponent().

    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        ReadImg()
    End Sub
    Private Sub ReadImg()
        Using db As New dbAllDataContext()
            Dim f = db.ПеревозчикиБаза.Where(Function(x) x.ID = IDp).FirstOrDefault()
            If f IsNot Nothing Then
                If f.ФотоДанные IsNot Nothing Then
                    Dim df = byteArrayToImage(f.ФотоДанные.ToArray)

                    If df IsNot Nothing Then
                        Clipboard.Clear()
                        Clipboard.SetImage(df)
                        RichTextBox1.Text = String.Empty
                        RichTextBox1.Paste()
                        Label2.Text = f.ДатаИзменения
                    End If

                End If

                If f.ФотоДанные2 IsNot Nothing Then
                    Dim df = byteArrayToImage(f.ФотоДанные2.ToArray)

                    If df IsNot Nothing Then
                        Clipboard.Clear()
                        Clipboard.SetImage(df)
                        RichTextBox2.Text = String.Empty
                        RichTextBox2.Paste()
                    End If

                End If

                'IO.File.WriteAllBytes(RichTextBox1.Text, f.ФотоДанные.ToArray)
            End If
        End Using
    End Sub

    Public Shared Function imageToByteArray(ByVal imageIn As System.Drawing.Image) As Byte()
        Dim ms As MemoryStream = New MemoryStream()
        imageIn.Save(ms, System.Drawing.Imaging.ImageFormat.Gif)
        Return ms.ToArray()
    End Function

    Public Shared Function byteArrayToImage(ByVal byteArrayIn As Byte()) As Image
        Dim ms As MemoryStream = New MemoryStream(byteArrayIn)
        Dim returnImage As Image
        Try
            returnImage = Image.FromStream(ms)
        Catch ex As Exception
            returnImage = Nothing
        End Try

        Return returnImage
    End Function

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

        'Dim msd = Encoding.Default.GetBytes(RichTextBox1.Text)

        'Dim msd As New ImageConverter
        'CType(msd.ConvertTo(RichTextBox1.Text, GetType(Byte())), Byte())

        'Dim iData As IDataObject = Clipboard.GetDataObject()
        'Dim fg = iData.GetFormats

        If Clipboard.ContainsImage() = False Then Return
        Dim fg = TabControl1.SelectedTab
        Dim returnImage As Image = Nothing
        returnImage = My.Computer.Clipboard.GetImage()
        'Dim bt = IO.File.ReadAllBytes(RichTextBox1.Text)
        Dim mo As New AllUpd
        If ФормаЗапуска = False Then
            Using db As New dbAllDataContext()
                Dim f = db.ПеревозчикиБаза.Where(Function(x) x.ID = IDp).FirstOrDefault()
                If f IsNot Nothing Then
                    If fg.TabIndex = 0 Then
                        f.ФотоДанные = imageToByteArray(returnImage)
                        RichTextBox1.Text = String.Empty
                    Else
                        f.ФотоДанные2 = imageToByteArray(returnImage)
                        RichTextBox2.Text = String.Empty
                    End If

                    f.ДатаИзменения = Now
                    db.SubmitChanges()
                    mo.ПеревозчикиБазаAllAsync()
                    MessageBox.Show("Фото принято!", Рик)

                    Clipboard.Clear()
                End If
            End Using

        Else
            If fg.TabIndex = 0 Then
                Imagese = imageToByteArray(returnImage)
                RichTextBox1.Text = String.Empty
            Else
                Imagese2 = imageToByteArray(returnImage)
                RichTextBox2.Text = String.Empty
            End If

            MessageBox.Show("Фото принято!", Рик)

            Clipboard.Clear()
        End If


    End Sub








    Private Sub ImageForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If IDp > 0 Then
            Button2.Enabled = True
        End If
        If Read = True Then
            ReadImg()
        End If
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        If Clipboard.ContainsImage() = False Then
            MessageBox.Show("Нет данных для вставки!", Рик)
            Return
        End If
        RichTextBox1.Text = String.Empty
        RichTextBox1.Paste()

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        RichTextBox1.Text = String.Empty

    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        If Clipboard.ContainsImage() = False Then Return

        Dim returnImage As Image = Nothing
        returnImage = My.Computer.Clipboard.GetImage()
        'Dim bt = IO.File.ReadAllBytes(RichTextBox1.Text)
        Dim mo As New AllUpd
        If ФормаЗапуска = False Then
            Using db As New dbAllDataContext()
                Dim f = db.ПеревозчикиБаза.Where(Function(x) x.ID = IDp).FirstOrDefault()
                If f IsNot Nothing Then
                    Dim k1 = imageToByteArray(returnImage)
                    Dim k = f.ФотоДанные.ToArray
                    Dim bytnew = Combine2(k, k1)
                    f.ФотоДанные = bytnew
                    f.ДатаИзменения = Now
                    db.SubmitChanges()
                    mo.ПеревозчикиБазаAllAsync()
                    MessageBox.Show("Фото принято!", Рик)
                    RichTextBox1.Text = String.Empty
                    Clipboard.Clear()
                End If
            End Using

        End If
    End Sub
    Public Shared Function Combine2(ByVal first As Byte(), ByVal second As Byte()) As Byte()
        Dim ret As Byte() = New Byte(first.Length + second.Length - 1) {}
        Buffer.BlockCopy(first, 0, ret, 0, first.Length)
        Buffer.BlockCopy(second, 0, ret, first.Length, second.Length)
        Return ret
    End Function
End Class