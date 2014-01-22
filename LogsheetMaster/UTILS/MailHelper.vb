Imports System.Net.Mail

Public Class MailHelper
    ''' <summary>
    ''' Sends an mail message
    ''' </summary>
    ''' <param name="from">Sender address</param>
    ''' <param name="recepient">Recepient address</param>
    ''' <param name="bcc">Bcc recepient</param>
    ''' <param name="cc">Cc recepient</param>
    ''' <param name="subject">Subject of mail message</param>
    ''' <param name="body">Body of mail message</param>
    Public Shared Sub SendMailMessage(ByVal from As String, ByVal recepient As String, ByVal bcc As String, ByVal cc As String, ByVal subject As String, ByVal body As String, ByVal pAttachment As String)
        ' Instantiate a new instance of MailMessage
        Dim mMailMessage As New MailMessage()
        Dim Attach As System.Net.Mail.Attachment

        Attach = New Attachment(pAttachment)

        ' Set the sender address of the mail message
        mMailMessage.From = New MailAddress(from)
        ' Set the recepient address of the mail message
        mMailMessage.To.Add(New MailAddress(recepient))

        ' Check if the bcc value is nothing or an empty string
        If Not bcc Is Nothing And bcc <> String.Empty Then
            ' Set the Bcc address of the mail message
            mMailMessage.Bcc.Add(New MailAddress(bcc))
        End If

        ' Check if the cc value is nothing or an empty value
        If Not cc Is Nothing And cc <> String.Empty Then
            ' Set the CC address of the mail message
            mMailMessage.CC.Add(New MailAddress(cc))
        End If

        ' Set the subject of the mail message
        mMailMessage.Subject = subject
        ' Set the body of the mail message
        mMailMessage.Body = body

        ' Set the format of the mail message body as HTML
        mMailMessage.IsBodyHtml = True
        ' Set the priority of the mail message to normal
        mMailMessage.Priority = MailPriority.Normal

        mMailMessage.Attachments.Add(Attach)

        ' Instantiate a new instance of SmtpClient
        Dim mSmtpClient As New SmtpClient()
        mSmtpClient.UseDefaultCredentials = False
        mSmtpClient.Credentials = New Net.NetworkCredential("ofpweb@gmail.com", "ofpweb4321")
        mSmtpClient.Port = 587
        mSmtpClient.Host = "smtp.gmail.com"
        mSmtpClient.DeliveryMethod = SmtpDeliveryMethod.Network
        mSmtpClient.EnableSsl = True
        mSmtpClient.Timeout = 300000 'milliseconds = 5 minutes
        ' Send the mail message
        mSmtpClient.Send(mMailMessage)
    End Sub
End Class