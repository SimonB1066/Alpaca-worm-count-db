Public Class FormNote


    Private Sub Form1_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Me.TopMost = True
        Dim x As Integer
        Dim y As Integer
        x = Screen.PrimaryScreen.WorkingArea.Width
        y = Screen.PrimaryScreen.WorkingArea.Height - Me.Height - 30

        Do Until x = Screen.PrimaryScreen.WorkingArea.Width - Me.Width - 30
            x = x - 1
            Me.Location = New Point(x, y)
        Loop

        MonthCalendar1.MinDate = DateTime.Now


        Me.Refresh()
    End Sub



    Private Sub MonthCalendar1_DateSelected(sender As Object, e As System.Windows.Forms.DateRangeEventArgs) Handles MonthCalendar1.DateSelected
        Dim dtNow As DateTime = DateTime.Now
        GlobalVariables.ReTestAddedTime = MonthCalendar1.SelectionRange.Start.Subtract(dtNow).Days + 1
        GlobalVariables.ReminderState = True
        FormMain.ToolStripButton11.Visible = True
        FormMain.SetXMLData()
        Me.Close()
    End Sub


    Private Sub RectangleShape11_MouseEnter(sender As Object, e As System.EventArgs) Handles RectangleShape11.MouseEnter
        RectangleShape11.BackColor = Color.LightSeaGreen
        Label5.BackColor = Color.LightSeaGreen
    End Sub


    Private Sub RectangleShape11_MouseLeave(sender As Object, e As System.EventArgs) Handles RectangleShape11.MouseLeave
        RectangleShape11.BackColor = Color.Teal
        Label5.BackColor = Color.Teal
    End Sub



    Private Sub RectangleShape11_Click(sender As System.Object, e As System.EventArgs) Handles RectangleShape11.Click
        MonthCalendar1.Visible = True
    End Sub

    Private Sub FormNote_MouseClick(sender As Object, e As System.Windows.Forms.MouseEventArgs) Handles Me.MouseClick
        MonthCalendar1.Visible = False
    End Sub

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        Try
            Close()
        Catch
        End Try
    End Sub

    Private Sub Button1_MouseEnter(sender As Object, e As System.EventArgs) Handles Button1.MouseEnter
        Button1.BackColor = Color.FromArgb(127, Color.Teal)
    End Sub


    Private Sub Button1_MouseLeave(sender As Object, e As System.EventArgs) Handles Button1.MouseLeave
        Button1.BackColor = Color.Teal
    End Sub
End Class