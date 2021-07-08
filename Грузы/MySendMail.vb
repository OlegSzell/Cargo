Imports System.Net
Imports System.Net.Mail

Public Class MySendMail
    Private Txt As String
    Private Subj As String
    Sub New(_txt As String, _subj As String)
        Txt = _txt
        Subj = _subj
    End Sub

    Public Async Sub Mail()


        Using mm As New MailMessage(EmailPass(0).Addres, EmailPass(0).Addres, Subj, Txt)
            mm.IsBodyHtml = False
            Using sc As New SmtpClient("smtp.yandex.ru", 25)
                sc.EnableSsl = True
                sc.DeliveryMethod = SmtpDeliveryMethod.Network
                sc.UseDefaultCredentials = False
                sc.Credentials = New NetworkCredential(EmailPass(0).Addres, EmailPass(0).Pass)
                Try
                    'smtp.Send(mm)
                    Await sc.SendMailAsync(mm)
                    'MessageBox.Show("Письмо отправлено!", Рик)
                Catch ex As Exception
                    MessageBox.Show("Ошибка передачи " & ex.ToString)
                End Try


            End Using
        End Using











        'Dim mm As New MailMessage
        'Dim smtp As New SmtpClient

        'mm.From = New MailAddress(EmailPass(1).Addres, Environment.UserDomainName, Text.Encoding.UTF8)
        'mm.To.Add(New MailAddress(EmailPass(1).Addres))
        'mm.Subject = Subj
        'mm.Body = "<h2>" & Txt & "</h2>"
        'mm.IsBodyHtml = True

        'smtp.Host = "smtp.gmail.com"


        'If EmailPass(1).Addres IsNot Nothing Then
        '    mm.CC.Add(EmailPass(1).Addres)
        'End If

        'smtp.EnableSsl = True
        'Dim NetworkCred As New NetworkCredential
        'NetworkCred.UserName = EmailPass(1).Addres '//gmail user name
        'NetworkCred.Password = EmailPass(1).Pass '// password
        'smtp.UseDefaultCredentials = True
        'smtp.Credentials = NetworkCred
        'smtp.Port = 587 ' //Gmail port For e-mail 465 Or 587

        'Try
        '    'smtp.Send(mm)
        '    Await smtp.SendMailAsync(mm)
        '    MessageBox.Show("Письмо отправлено!", Рик)
        'Catch ex As Exception
        '    MessageBox.Show("Ошибка передачи " & ex.ToString)
        'End Try







        'Dim from As New MailAddress(EmailPass.Addres, Environment.UserDomainName)
        'Dim tos As New MailAddress(EmailPass.Addres)
        'Dim m As New MailMessage(from, tos)
        'm.Subject =Subj
        'm.Body = "<h2>" & Txt & "</h2>"
        'm.IsBodyHtml = True
        'Dim smtp As New SmtpClient("smtp.yandex.ru", 465)
        'smtp.Credentials = New NetworkCredential(EmailPass.Addres, EmailPass.Pass)
        'smtp.EnableSsl = True
        'Try
        '    Await smtp.SendMailAsync(m)
        '    MessageBox.Show("Письмо отправлено!", Рик)
        'Catch ex As Exception
        '    MessageBox.Show("Ошибка передачи " & ex.ToString)
        'End Try


    End Sub

End Class
