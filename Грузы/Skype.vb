Imports System.Threading

Public Class Skype
    Private NameOrg As String = Nothing
    Private DateSkype As String = Nothing
    Private Skype As String
    Private Clas As Grid1ЖурналClass = Nothing
    Private Flag As Boolean = False

    Sub New()

        ' Этот вызов является обязательным для конструктора.
        InitializeComponent()

        ' Добавить код инициализации после вызова InitializeComponent().

    End Sub
    Sub New(ByVal _NameOrg As String, ByVal _DateSkype As String, Optional _Skype As Long = Nothing)
        NameOrg = _NameOrg
        DateSkype = _DateSkype
        Skype = _Skype
        ' Этот вызов является обязательным для конструктора.
        InitializeComponent()

        ' Добавить код инициализации после вызова InitializeComponent().

    End Sub
    Sub New(ByVal _Clas As Grid1ЖурналClass, ByVal _Flag As Boolean)
        Clas = _Clas
        Flag = _Flag
        ' Этот вызов является обязательным для конструктора.
        InitializeComponent()

        ' Добавить код инициализации после вызова InitializeComponent().

    End Sub
    Private Sub Skype_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If Clas Is Nothing Then
            Label3.Text = NameOrg
            Label2.Text = DateSkype
            RichTextBox1.Text = Skype
        Else
            If Flag = True Then
                Button1.Enabled = True
            Else
                Button1.Enabled = False
            End If

            Label3.Text = Clas.Перевозчик
            Label2.Text = Clas.SkypeDate
            RichTextBox1.Text = Clas.Skype
        End If

    End Sub


    Private Sub UpdRich(ByVal Ric As String)
        Using db As New dbAllDataContext(_cn3)
            Dim f = db.ЖурналПеревозчик.Where(Function(x) x.Код = Clas.КодЖурналПеревозчик).FirstOrDefault()
            If f IsNot Nothing Then
                f.Skype = Ric
                f.SkypeDate = Now
                db.SubmitChanges()
                Dim mo As New AllUpd
                mo.ЖурналПеревозчикAll()
            End If
        End Using
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim f1 = RichTextBox1.Text
        Dim f As New Thread(New ThreadStart(Sub() UpdRich(f1)))
        f.Start()

    End Sub
End Class