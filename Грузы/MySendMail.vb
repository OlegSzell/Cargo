Imports System.Net
Imports System.Net.Mail

Public Class MySendMail
    Private Txt As String
    Private Subj As String
    Private EmailTo As String
    Private EmailFrom As String
    Private PasswFrom As String
    Sub New(_txt As String, _subj As String)
        Txt = _txt
        Subj = _subj
    End Sub

    Sub New(_txt As String, _subj As String, _emailTo As String)
        Txt = _txt
        Subj = _subj
        EmailTo = _emailTo

    End Sub
    Sub New(_txt As String, _subj As String, _emailTo As String, _emailFrom As String, _passwFrom As String)
        Txt = _txt
        Subj = _subj
        EmailTo = _emailTo
        PasswFrom = _passwFrom
        EmailFrom = _emailFrom

    End Sub

    Public Async Sub Mail()
        Dim f As EmailTb
        Using db As New DbAllDataContext(_cn3)
            f = (From x In db.EmailTb
                 Where x.ID = 2
                 Select x).FirstOrDefault()
        End Using
        If f IsNot Nothing Then
            '// отправитель - устанавливаем адрес и отображаемое в письме имя
            Dim from As New MailAddress(f.Addres, "Rickmans")
            '// кому отправляем
            Dim tos As New MailAddress(EmailPass(0).Addres)
            '// создаем объект сообщения
            Dim m As New MailMessage(from, tos)
            '// тема письма
            m.Subject = Subj
            '// текст письма
            m.Body = Txt
            '// письмо представляет код html
            m.IsBodyHtml = False
            '// адрес smtp-сервера и порт, с которого будем отправлять письмо
            Dim smtp As New SmtpClient("smtp.gmail.com", 587)
            '// логин и пароль
            smtp.Credentials = New NetworkCredential(f.Addres, f.Pass)
            smtp.EnableSsl = True
            Try
                Await smtp.SendMailAsync(m)
            Catch ex As Exception
                MessageBox.Show(ex.ToString)
            End Try

        End If


        'Using mm As New MailMessage(EmailPass(0).Addres, EmailPass(0).Addres, Subj, Txt)
        '    mm.IsBodyHtml = False
        '    Using sc As New SmtpClient("smtp.yandex.ru", 25)
        '        sc.EnableSsl = True
        '        sc.DeliveryMethod = SmtpDeliveryMethod.Network
        '        sc.UseDefaultCredentials = False
        '        sc.Credentials = New NetworkCredential(EmailPass(0).Addres, EmailPass(0).Pass)
        '        Try
        '            'smtp.Send(mm)
        '            Await sc.SendMailAsync(mm)
        '            'sc.SendMailAsync(mm)
        '            'MessageBox.Show("Письмо отправлено!", Рик)
        '        Catch ex As Exception
        '            MessageBox.Show("Ошибка передачи " & ex.ToString)
        '        End Try


        '    End Using
        'End Using



    End Sub

    Public Sub Mail2()
        '// отправитель - устанавливаем адрес и отображаемое в письме имя
        Dim from As New MailAddress(EmailFrom, "Rickmans")
        '// кому отправляем
        Dim tos As New MailAddress(EmailTo)
        '// создаем объект сообщения
        Dim m As New MailMessage(from, tos)
        '// тема письма
        m.Subject = Subj
        '// текст письма
        m.Body = Txt
        '// письмо представляет код html
        m.IsBodyHtml = False
        '// адрес smtp-сервера и порт, с которого будем отправлять письмо
        Dim smtp As New SmtpClient("smtp.gmail.com", 587)
        '// логин и пароль
        smtp.Credentials = New NetworkCredential(EmailFrom, PasswFrom)
        smtp.EnableSsl = True
        Try
            If Not EmailTo.Contains("9") Then
                smtp.Send(m)
            End If
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try

    End Sub


End Class
