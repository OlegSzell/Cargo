Imports EASendMail
Public Class SendEmail

    Public Sub insert(d As String)
        Dim oMail As New SmtpMail("TryIt")
        Dim oSmtp As New SmtpClient()
        oMail.From = EmailPass(0).Addres
        oMail.To = EmailPass(0).Addres
        oMail.Subject = "Создан новый рейс"
        oMail.TextBody = d
        Dim oServer As New SmtpServer("")
        oServer.Server = "smtp.yandex.ru"
        oServer.User = EmailPass(0).Addres
        oServer.Password = EmailPass(0).Pass
        oServer.Port = 465
        oServer.ConnectType = SmtpConnectType.ConnectSSLAuto 'если требуется обязательно ssl
        Try
            oSmtp.SendMail(oServer, oMail)
            'MessageBox.Show("Письмо отправлено!", Рик)
        Catch ex As Exception
            MessageBox.Show("Ошибка передачи " & ex.ToString)
        End Try




        'If TextBox1.Text = "6289925@mail.ru" Then
        '    'oServer.Server = "smtp.mail.ru"
        '    'oServer.User = "6289925@mail.ru"
        '    'oServer.Password = "6807057a"
        '    'oServer.Port = 465
        '    'oServer.ConnectType = SmtpConnectType.ConnectSSLAuto 'если требуется обязательно ssl
        'ElseIf TextBox1.Text = "1389925@gmail.com" Then
        '    'oServer.Server = "smtp.gmail.com"
        '    'oServer.User = "1389925@gmail.com"
        '    'oServer.Password = "oleg110403"
        '    'oServer.Port = 465
        '    'oServer.ConnectType = SmtpConnectType.ConnectSSLAuto 'если требуется обязательно ssl
        'ElseIf TextBox1.Text = "os@2trans.by" Then
        '    'oServer.Server = "smtp.yandex.ru"
        '    'oServer.User = "os@2trans.by"
        '    'oServer.Password = "oleg61351127441"
        '    'oServer.Port = 465
        '    'oServer.ConnectType = SmtpConnectType.ConnectSSLAuto 'если требуется обязательно ssl
        'Else

        'End If

        'Dim oServer As New SmtpServer("smtp.emailarchitect.net")
        'oServer.User = "test@emailarchitect.net"
        'oServer.Password = "testpassword"
        'Try
        '    oSmtp.SendMail(oServer, oMail)
        '    MessageBox.Show("Письмо отправлено!", Рик)
        'Catch ex As Exception
        '    MessageBox.Show("Ошибка передачи " & ex.ToString)
        'End Try
    End Sub
    Public Sub delete(d As String)
        Dim oMail As New SmtpMail("TryIt")
        Dim oSmtp As New SmtpClient()
        oMail.From = EmailPass(0).Addres
        oMail.To = EmailPass(0).Addres
        oMail.Subject = "Удален рейс"
        oMail.TextBody = d
        Dim oServer As New SmtpServer("")
        oServer.Server = "smtp.yandex.ru"
        oServer.User = EmailPass(0).Addres
        oServer.Password = EmailPass(0).Pass
        oServer.Port = 465
        oServer.ConnectType = SmtpConnectType.ConnectSSLAuto 'если требуется обязательно ssl
        Try
            oSmtp.SendMail(oServer, oMail)
            'MessageBox.Show("Письмо отправлено!", Рик)
        Catch ex As Exception
            MessageBox.Show("Ошибка передачи " & ex.ToString)
        End Try
    End Sub
    Public Sub update(d As String)
        Dim oMail As New SmtpMail("TryIt")
        Dim oSmtp As New SmtpClient()
        oMail.From = EmailPass(0).Addres
        oMail.To = EmailPass(0).Addres
        oMail.Subject = "Изменен рейс"
        oMail.TextBody = d
        Dim oServer As New SmtpServer("")
        oServer.Server = "smtp.yandex.ru"
        oServer.User = EmailPass(0).Addres
        oServer.Password = EmailPass(0).Pass
        oServer.Port = 465
        oServer.ConnectType = SmtpConnectType.ConnectSSLAuto 'если требуется обязательно ssl
        Try
            oSmtp.SendMail(oServer, oMail)
            'MessageBox.Show("Письмо отправлено!", Рик)
        Catch ex As Exception
            MessageBox.Show("Ошибка передачи " & ex.ToString)
        End Try
    End Sub
    Public Sub SendRnb(d As String, _to As String)
        Dim oMail As New SmtpMail("TryIt")
        Dim oSmtp As New SmtpClient()
        oMail.From = EmailPass(0).Addres
        oMail.To = _to
        oMail.Subject = "Введите данные в поле 'Подтверждения'"
        oMail.TextBody = d
        Dim oServer As New SmtpServer("")
        oServer.Server = "smtp.yandex.ru"
        oServer.User = EmailPass(0).Addres
        oServer.Password = EmailPass(0).Pass
        oServer.Port = 465
        oServer.ConnectType = SmtpConnectType.ConnectSSLAuto 'если требуется обязательно ssl
        Try
            oSmtp.SendMail(oServer, oMail)
            'MessageBox.Show("Письмо отправлено!", Рик)
        Catch ex As Exception
            MessageBox.Show("Ошибка передачи " & ex.ToString)
        End Try
    End Sub

End Class
