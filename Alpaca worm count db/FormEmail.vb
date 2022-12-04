Imports System.Net.Mail
Imports System.IO
Imports System.Threading
Public Class FormEmail
    Private Sub FormEmail_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Me.TopMost = True
        CenterToScreen()
        Dim myThread As New Thread(AddressOf email)
        myThread.IsBackground = True
        myThread.Start()
        Me.Close()

    End Sub


    Private Sub email()
        Try
            NotifyIcon1.Icon = My.Resources.WDB
            NotifyIcon1.Visible = True
            Dim msg As String = "Sending E Mail" & vbNewLine & "to " & GlobalVariables.Email & vbNewLine & GlobalVariables.Subject
            NotifyIcon1.ShowBalloonTip(10000, "Alpaca worm count database", msg, ToolTipIcon.Info)
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
                MsgBox("Email recipient has to been entered, go to settings and correct.", vbOKOnly & vbCritical, "Error")
                Exit Sub
            End If
            e_mail.To.Add(GlobalVariables.Email)
            e_mail.Subject = GlobalVariables.Subject
            e_mail.IsBodyHtml = True

            'e_mail.Body = GlobalVariables.Body
            Dim body As String = File.ReadAllText("C:\WormCountDATA\body.html")
            body = body.Replace("<%Tag01%>", GlobalVariables.AlpacaName)
            e_mail.Body = body
            If GlobalVariables.Attachment <> "" Then
                Dim AttachmentFile As String = GlobalVariables.Attachment
                e_mail.Attachments.Add(New Attachment(AttachmentFile))
            End If
            Smtp_Server.Send(e_mail)
            e_mail.Attachments.Dispose()
            NotifyIcon1.ShowBalloonTip(7000, "Alpaca worm count database", "Sent", ToolTipIcon.Info)
        Catch ex As Exception

            Dim st As New StackTrace(True)
            st = New StackTrace(ex, True)
            GlobalVariables.Linenumber = st.GetFrame(0).GetFileLineNumber().ToString()
            Dim f As New StackFrame
            FormMain.ErrorHandler(ex, System.Reflection.MethodBase.GetCurrentMethod().Name, GlobalVariables.Linenumber, f)
            GlobalVariables.EmailSending = False
            Exit Sub
        End Try
        GlobalVariables.EmailSending = False

    End Sub
End Class