Imports System.Net.Mail
Module EnvioMail
    Dim mensaje As String
    Dim asunto As String
    Dim enviado As Boolean
    Public Function enviomail(ByVal para As String, ByVal tipoaviso As String)

        If tipoaviso = "RegistroCliente" Then
            mensaje = "Le agradecemos su suscripcion, se ah realizado el registro de sus datos en el sistema"
            asunto = "Suscripcion"
        End If

        If tipoaviso = "RegistroInmueble" Then
            mensaje = FormPrincipal.TxtFormPrincipalIngresoPropiedadesNotificacion.Text
            asunto = "Publicación"
        End If


        If tipoaviso = "Recordatorio" Then
            mensaje = "Le recordamos que el dia de mañana esta pactado su visita al inmueble"
            asunto = "Recordatorio"
        End If



        Try
            Dim Smtp_Server As New SmtpClient
            Dim e_mail As New MailMessage()
            Smtp_Server.UseDefaultCredentials = True
            Smtp_Server.Credentials = New Net.NetworkCredential("enigmagrupo3ic@gmail.com", "2019mmaaenigma")
            Smtp_Server.Port = 25
            Smtp_Server.EnableSsl = True
            Smtp_Server.Host = "smtp.gmail.com"

            e_mail = New MailMessage()
            e_mail.From = New MailAddress("enigmagrupo3ic@gmail.com")
            e_mail.To.Add(para)
            e_mail.Subject = asunto
            e_mail.IsBodyHtml = False
            e_mail.Body = mensaje
            Smtp_Server.Send(e_mail)
            MsgBox("Mail enviado correctamente")

            enviado = True

        Catch error_t As Exception
            MsgBox(error_t.Message)
            enviado = False
        End Try


        Return enviado
    End Function









End Module
