Imports System.Net.Mail

Public Class FormEmail
    Private Sub FormEmail_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Me.TopMost = True
        CenterToScreen()
        Timer1.Enabled = True



    End Sub


    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Try

            'GlobalVariables.Email = "aig1066@hotmail.co.uk"

            GlobalVariables.EmailSending = True
            Dim Smtp_Server As New SmtpClient
            Dim e_mail As New MailMessage()
            Smtp_Server.UseDefaultCredentials = False
            Smtp_Server.Credentials = New Net.NetworkCredential(GlobalVariables.EmailUsername, GlobalVariables.EmailPassword)
            Smtp_Server.Port = 587
            Smtp_Server.EnableSsl = True
            Smtp_Server.Host = GlobalVariables.EmailServer

            e_mail = New MailMessage()
            e_mail.From = New MailAddress(GlobalVariables.EmailUsername)
            If GlobalVariables.Email = "" Or GlobalVariables.Email = "sample@farm.com" Then
                Me.Close()
                Timer1.Enabled = False
                MsgBox("Email recipient has to been entered, go to settings and correct.", vbOKOnly & vbCritical, "Error")
                Exit Sub
            End If
            e_mail.To.Add(GlobalVariables.Email)
            e_mail.Subject = GlobalVariables.Subject
            e_mail.IsBodyHtml = False
            e_mail.Body = GlobalVariables.Body
            If GlobalVariables.Attachment <> "" Then
                Dim AttachmentFile As String = GlobalVariables.Attachment
                e_mail.Attachments.Add(New Attachment(AttachmentFile))
            End If
            Smtp_Server.Send(e_mail)

            e_mail.Attachments.Dispose()



        Catch ex As Exception
            Me.Close()
            Dim st As New StackTrace(True)
            st = New StackTrace(ex, True)
            GlobalVariables.Linenumber = st.GetFrame(0).GetFileLineNumber().ToString()
            Dim f As New StackFrame
            FormMain.ErrorHandler(ex, System.Reflection.MethodBase.GetCurrentMethod().Name, GlobalVariables.Linenumber, f)
            GlobalVariables.EmailSending = False
            Exit Sub
        End Try
        GlobalVariables.EmailSending = False
        Me.Close()
    End Sub
End Class