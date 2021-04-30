Imports System.Net
Imports System.Net.Mail
Imports System.Text
Public Class cEnviaCorreo
    Dim oMail As New MailMessage()
    Dim oSmtpServer As New SmtpClient()


    Private sCorreoDe As String
    Private sCorreoDeContraseña As String
    Private sCorreoPara As String
    Private sCorreoCopia As String
    Private sCorreoCopiaoculta As String
    Private sCorreoAsunto As String
    Private sCorreoAlias As String
    Private sSmtpPort As String
    Private sSmtpHost As String
    Private sStrB As New StringBuilder
    Private sSSL As String
    Private oArchivoAdjunto As New Object
    Private sMensajeError As String

    Public Property CorreoDe As String
        Get
            Return sCorreoDe
        End Get
        Set(value As String)
            sCorreoDe = value
        End Set
    End Property

    Public Property CorreoDeContraseña As String
        Get
            Return sCorreoDeContraseña
        End Get
        Set(value As String)
            sCorreoDeContraseña = value
        End Set
    End Property

    Public Property CorreoPara As String
        Get
            Return sCorreoPara
        End Get
        Set(value As String)
            sCorreoPara = value
        End Set
    End Property

    Public Property CorreoCopia As String
        Get
            Return sCorreoCopia
        End Get
        Set(value As String)
            sCorreoCopia = value
        End Set
    End Property

    Public Property CorreoCopiaoculta As String
        Get
            Return sCorreoCopiaoculta
        End Get
        Set(value As String)
            sCorreoCopiaoculta = value
        End Set
    End Property

    Public Property CorreoAsunto As String
        Get
            Return sCorreoAsunto
        End Get
        Set(value As String)
            sCorreoAsunto = value
        End Set
    End Property

    Public Property CorreoAlias As String
        Get
            Return sCorreoAlias
        End Get
        Set(value As String)
            sCorreoAlias = value
        End Set
    End Property

    Public Property SmtpPort As String
        Get
            Return sSmtpPort
        End Get
        Set(value As String)
            sSmtpPort = value
        End Set
    End Property

    Public Property SmtpHost As String
        Get
            Return sSmtpHost
        End Get
        Set(value As String)
            sSmtpHost = value
        End Set
    End Property

    Public Property StrB As StringBuilder
        Get
            Return sStrB
        End Get
        Set(value As StringBuilder)
            sStrB = value
        End Set
    End Property

    Public Property SSL As String
        Get
            Return sSSL
        End Get
        Set(value As String)
            sSSL = value
        End Set
    End Property

    Public Property ArchivoAdjunto As Object
        Get
            Return oArchivoAdjunto
        End Get
        Set(value As Object)
            oArchivoAdjunto = value
        End Set
    End Property

    Public Property MensajeError As String
        Get
            Return sMensajeError
        End Get
        Set(value As String)
            sMensajeError = value
        End Set
    End Property

    Public Function Envia(oDatosCorreo As cEnviaCorreo)
        Try


            oMail.From = New MailAddress(oDatosCorreo.CorreoDe, oDatosCorreo.CorreoAlias)
            oMail.To.Clear()
            oMail.CC.Clear()
            oMail.Bcc.Clear()


            oMail.To.Add(CorreoPara)
            If CorreoCopia <> "" Then
                oMail.CC.Add(CorreoCopia)
            End If

            'oMail.Bcc.Add(CorreoCopiaoculta)
            oMail.Subject = CorreoAsunto
            'If ArchivoAdjunto IsNot Nothing Then
            '    oMail.Attachments.Add(ArchivoAdjunto)
            'End If


            oMail.Body = StrB.ToString
            oMail.Priority = MailPriority.Normal
            oMail.BodyEncoding = Encoding.UTF8
            oMail.IsBodyHtml = True
            oSmtpServer.EnableSsl = True


            oSmtpServer.Port = oDatosCorreo.SmtpPort
            oSmtpServer.Host = oDatosCorreo.SmtpHost
            oSmtpServer.Credentials = New Net.NetworkCredential(oDatosCorreo.CorreoDe, oDatosCorreo.CorreoDeContraseña)
            oSmtpServer.Send(oMail)

            Return True

        Catch ex As Exception
            oDatosCorreo.MensajeError = ex.Message
            Return False
        End Try
    End Function
End Class
