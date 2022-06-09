Imports System.ComponentModel
Imports System.IO.DirectoryInfo

Public Class FormBackup

    Private Sub ToolStrip1_MouseHover(sender As Object, e As EventArgs) Handles ToolStrip1.MouseHover
        Me.Cursor = Cursors.Hand '
    End Sub
    Private Sub ToolStrip1_MouseLeave(sender As Object, e As EventArgs) Handles ToolStrip1.MouseLeave
        Me.Cursor = Cursors.Arrow
    End Sub



    Private Sub FormBackup_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load

        Me.TopMost = True
        'Check if the directory is there

        ProgressBar1.Style = ProgressBarStyle.Marquee
        PictureBox1.Visible = True
        PictureBox1.Image = My.Resources.backup

        CenterToScreen()
        Me.Icon = My.Resources.Form


    End Sub


    Public Sub BackupFile()

    End Sub



    Public Sub Backup()
        ToolStripButton1.Enabled = False

        Try


            GlobalVariables.bar = 0
            ProgressBar1.Visible = True
            ProgressBar1.Enabled = True
            BackgroundWorker1.RunWorkerAsync()


        Catch ex As Exception
            Dim st As New StackTrace(True)
            st = New StackTrace(ex, True)
            GlobalVariables.Linenumber = st.GetFrame(0).GetFileLineNumber().ToString()
            Dim f As New StackFrame
            FormMain.ErrorHandler(ex, System.Reflection.MethodBase.GetCurrentMethod().Name, GlobalVariables.Linenumber, f)
        End Try



    End Sub







    Private Sub BackgroundWorker1_DoWork(sender As Object, e As ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork

        BackupFile()

    End Sub





    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        If ProgressBar1.Enabled = True Then
            If GlobalVariables.bar < ProgressBar1.Maximum - 1 Then
                GlobalVariables.bar = GlobalVariables.bar + 1
                ProgressBar1.Value = GlobalVariables.bar
            End If
        End If
    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(sender As Object, e As RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        ToolStripButton1.Enabled = True
        Close()
    End Sub

    Private Sub ToolStripButton1_Click(sender As Object, e As EventArgs) Handles ToolStripButton1.Click
        Backup()
    End Sub

    Private Sub ToolStripButton2_Click(sender As Object, e As EventArgs) Handles ToolStripButton2.Click
        Me.Close()
    End Sub
End Class