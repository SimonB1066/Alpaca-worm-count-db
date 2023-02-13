

Imports System.Xml
Imports System.Globalization
'Imports Microsoft.Win32

Public Class FormSettings
    Dim Save As Boolean = False

    Private Sub ToolStrip1_MouseHover(sender As Object, e As EventArgs) Handles ToolStrip1.MouseHover
        Me.Cursor = Cursors.Hand '
    End Sub
    Private Sub ToolStrip1_MouseLeave(sender As Object, e As EventArgs) Handles ToolStrip1.MouseLeave
        Me.Cursor = Cursors.Arrow
    End Sub

    Private Sub FormSettings_FormClosed(sender As Object, e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed



        GlobalVariables.EggType(1).Bias = TrackBar1.Value / 10
        GlobalVariables.EggType(2).Bias = TrackBar2.Value / 10
        GlobalVariables.EggType(3).Bias = TrackBar3.Value / 10
        GlobalVariables.EggType(4).Bias = TrackBar4.Value / 10
        GlobalVariables.EggType(5).Bias = TrackBar5.Value / 10
        GlobalVariables.EggType(6).Bias = TrackBar6.Value / 10
        GlobalVariables.EggType(7).Bias = TrackBar7.Value / 10
        GlobalVariables.EggType(8).Bias = TrackBar8.Value / 10
        GlobalVariables.EggType(9).Bias = TrackBar9.Value / 10
        GlobalVariables.EggType(10).Bias = TrackBar10.Value / 10


        If TextBox6.Text = "" Or TextBox6.Text = "New sample Farm" Then
            MessageBox.Show("You did not set the farm name please go back into setting and set it.", "Message")
        End If

        If Not IsValidEmail(GlobalVariables.Email) Then
            MessageBox.Show("Not a valid EMail address", "Error")
        End If

        FormMain.SetXMLData()
        If Save Then
            FormMain.DataGridViewRefresh()
        End If
    End Sub



    Private Sub FormSettings_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        CenterToScreen()

        Me.Icon = My.Resources.Form
        Save = False


    End Sub

    Private Sub FormSettings_Shown(sender As Object, e As EventArgs) Handles Me.Shown


        Label50.Text = GlobalVariables.EggType(1).Name
        Label51.Text = GlobalVariables.EggType(2).Name
        Label52.Text = GlobalVariables.EggType(3).Name
        Label53.Text = GlobalVariables.EggType(4).Name
        Label54.Text = GlobalVariables.EggType(5).Name
        Label55.Text = GlobalVariables.EggType(6).Name
        Label56.Text = GlobalVariables.EggType(7).Name
        Label57.Text = GlobalVariables.EggType(8).Name
        Label58.Text = GlobalVariables.EggType(9).Name
        Label59.Text = GlobalVariables.EggType(10).Name
        TrackBar1.Value = GlobalVariables.EggType(1).Bias * 10
        TrackBar2.Value = GlobalVariables.EggType(2).Bias * 10
        TrackBar3.Value = GlobalVariables.EggType(3).Bias * 10
        TrackBar4.Value = GlobalVariables.EggType(4).Bias * 10
        TrackBar5.Value = GlobalVariables.EggType(5).Bias * 10
        TrackBar6.Value = GlobalVariables.EggType(6).Bias * 10
        TrackBar7.Value = GlobalVariables.EggType(7).Bias * 10
        TrackBar8.Value = GlobalVariables.EggType(8).Bias * 10
        TrackBar9.Value = GlobalVariables.EggType(9).Bias * 10
        TrackBar10.Value = GlobalVariables.EggType(10).Bias * 10



        ToolStripButton3.Text = GlobalVariables.pass


        CheckBox2.Checked = GlobalVariables.BiasEnable
        vis()


        If GlobalVariables.StartAtBoot = True Then
            ToolStripButton4.Image = My.Resources.Windows1
        Else
            ToolStripButton4.Image = My.Resources.Windows11
        End If



        'New settings
        TextBox2.Text = GlobalVariables.Email
        TextBox6.Text = GlobalVariables.FarmName
        TextBox13.Text = GlobalVariables.EmailServer
        TextBox14.Text = GlobalVariables.EmailUsername
        TextBox15.Text = GlobalVariables.EmailPassword


        If IO.File.Exists(GlobalVariables.DbDriveLocation & "\Header.png") Then
            Dim fs As System.IO.FileStream
            ' Specify a valid picture file path on your computer.
            fs = New System.IO.FileStream(GlobalVariables.DbDriveLocation & "\Header.png",
                 IO.FileMode.Open, IO.FileAccess.Read)
            PictureBox1.Image = System.Drawing.Image.FromStream(fs)
            fs.Close()
        End If

        Refresh()


    End Sub



    Private Function GetxmlSMTP(ByVal Tag As String) As String

        Dim ret As String = ""
        Try
            Using Reader As XmlReader = XmlReader.Create(GlobalVariables.DbDriveLocation & "\MailServer.xml")

                While Reader.Read()
                    ' Check for start elements.
                    If Reader.IsStartElement() Then
                        If Reader.Name = Tag Then

                            ' Text data.
                            If Reader.Read() Then
                                ret = Reader.Value.Trim()
                            End If
                        End If
                    End If
                End While
                'Reader.Close()
            End Using

        Catch ex As Exception
            If ex.Message <> "Application is not installed." Then
                Dim st As New StackTrace(True)
                st = New StackTrace(ex, True)
                GlobalVariables.Linenumber = st.GetFrame(0).GetFileLineNumber().ToString()
                Dim f As New StackFrame
                FormMain.ErrorHandler(ex, System.Reflection.MethodBase.GetCurrentMethod().Name, GlobalVariables.Linenumber, f)
            End If
        End Try
        Return ret
    End Function
    Private Sub TextBox2_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox2.TextChanged

        GlobalVariables.Email = TextBox2.Text

    End Sub

    Public Function IsValidEmail(ByVal email As String) As Boolean
        'regular exp<b></b>ression pattern for valid email
        'addresses, allows for the following domains:
        'com,edu,info,gov,int,mil,net,org,biz,name,museum,coop,aero,pro,tv
        Dim pattern As String = "^[-a-zA-Z0-9][-.a-zA-Z0-9]*@[-.a-zA-Z0-9]+(.[-.a-zA-Z0-9]+)*." &
        "(com|edu|info|gov|int|mil|net|org|biz|name|museum|coop|aero|pro|tv|[a-zA-Z]{2})$"
        'Regular exp<b></b>ression object
        Dim check As New System.Text.RegularExpressions.Regex(pattern, System.Text.RegularExpressions.RegexOptions.IgnorePatternWhitespace)

        'boolean variable to return to calling method
        Dim valid As Boolean = False

        'make sure an email address was provided
        If String.IsNullOrEmpty(email) Then
            valid = False
        Else
            'use IsMatch to validate the address
            valid = check.IsMatch(email)
        End If
        'return the value to the calling method
        Return valid
    End Function



    Private Sub TextBox6_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox6.TextChanged
        GlobalVariables.FarmName = TextBox6.Text
    End Sub


    Private Sub ToolStripButton2_Click(sender As Object, e As EventArgs) Handles ToolStripButton2.Click
        If TextBox6.Text = "" Or TextBox6.Text = "New sample Farm" Then
            MessageBox.Show("Please enter a farm name", "Message")
            Exit Sub
        End If
        Me.Close()

    End Sub

    Private Sub ToolStripButton1_Click(sender As Object, e As EventArgs) Handles ToolStripButton1.Click
        FormMain.Email("Worm count database - Test email", "Test email", "")
    End Sub

    Private Sub ToolStripButton3_Click(sender As Object, e As EventArgs) Handles ToolStripButton3.Click
        FormPass.ShowDialog()
        ToolStripButton3.Text = GlobalVariables.pass

    End Sub

    Private Sub ToolStripButton4_Click(sender As Object, e As EventArgs) Handles ToolStripButton4.Click


        If GlobalVariables.StartAtBoot = True Then
            GlobalVariables.StartAtBoot = False
            ToolStripButton4.Image = My.Resources.Windows11
            FormMain.SetXMLData()
            BootStateChanged()
            Exit Sub
        End If
        If GlobalVariables.StartAtBoot = False Then
            GlobalVariables.StartAtBoot = True
            ToolStripButton4.Image = My.Resources.Windows1
            FormMain.SetXMLData()
            BootStateChanged()
        End If


        Refresh()
    End Sub

    Private Sub BootStateChanged()
        If GlobalVariables.StartAtBoot = True Then
            Dim applicationName As String = Application.ProductName
            Dim applicationPath As String = Application.ExecutablePath
            Dim regKey As Microsoft.Win32.RegistryKey
            regKey = Microsoft.Win32.Registry.CurrentUser.OpenSubKey("SOFTWARE\Microsoft\Windows\CurrentVersion\Run", True)
            regKey.SetValue(applicationName, """" & applicationPath & """")
            regKey.Close()
        Else
            Dim applicationName As String = Application.ProductName
            Dim regKey As Microsoft.Win32.RegistryKey
            regKey = Microsoft.Win32.Registry.CurrentUser.OpenSubKey("SOFTWARE\Microsoft\Windows\CurrentVersion\Run", True)
            regKey.DeleteValue(applicationName, False)
            regKey.Close()

        End If


    End Sub

    Private Sub ToolStripButton5_Click(sender As Object, e As EventArgs) Handles ToolStripButton5.Click
        Dim fd As OpenFileDialog = New OpenFileDialog()
        Dim strFileName As String
        Try
            fd.Title = "Open File Dialog"
            fd.InitialDirectory = My.Computer.FileSystem.SpecialDirectories.MyPictures
            fd.Filter = "All files (*.*)|*.*|Image files |*.png"
            fd.FilterIndex = 2
            fd.RestoreDirectory = True

            If fd.ShowDialog() = DialogResult.OK Then
                strFileName = fd.FileName

                PictureBox1.Image = FormMain.ResizeImage(Image.FromFile(strFileName))
                PictureBox1.SizeMode = PictureBoxSizeMode.StretchImage


                If IO.File.Exists(GlobalVariables.DbDriveLocation & "Header.png") Then
                    My.Computer.FileSystem.DeleteFile(GlobalVariables.DbDriveLocation & "Header.png")
                End If

                System.IO.File.Copy(strFileName, GlobalVariables.DbDriveLocation & "Header.png")

                If IO.File.Exists(GlobalVariables.DbDriveLocation & "Header.png") Then
                    Dim fs As System.IO.FileStream
                    ' Specify a valid picture file path on your computer.
                    fs = New System.IO.FileStream(GlobalVariables.DbDriveLocation & "Header.png",
                         IO.FileMode.Open, IO.FileAccess.Read)
                    PictureBox1.Image = System.Drawing.Image.FromStream(fs)

                    fs.Close()
                End If

            End If

        Catch
            MessageBox.Show("Error saving new image")
        End Try
    End Sub


    Private Sub Button1_Click(sender As Object, e As EventArgs)
        ' Call ShowDialog.
        Dim result As DialogResult = OpenFileDialog1.ShowDialog()

        ' Test result.
        If result = Windows.Forms.DialogResult.OK Then
            GlobalVariables.PacaManpath = OpenFileDialog1.FileName
        End If
    End Sub

    Private Sub CheckBox1_Click(sender As Object, e As EventArgs)

        Save = True
        vis()
    End Sub
    Private Sub CheckBox2_Click(sender As Object, e As EventArgs) Handles CheckBox2.Click
        GlobalVariables.BiasEnable = CheckBox2.Checked
        Save = True
        visBias()
    End Sub

    Public Sub vis()


    End Sub

    Public Sub visBias()
        If CheckBox2.Checked Then

            Label50.Visible = True
            Label51.Visible = True
            Label52.Visible = True
            Label53.Visible = True
            Label54.Visible = True
            Label55.Visible = True
            Label56.Visible = True
            Label57.Visible = True
            Label58.Visible = True
            Label59.Visible = True
            Label17.Visible = True
            Label18.Visible = True
            Label19.Visible = True
            Label20.Visible = True
            TrackBar1.Visible = True
            TrackBar2.Visible = True
            TrackBar3.Visible = True
            TrackBar4.Visible = True
            TrackBar5.Visible = True
            TrackBar6.Visible = True
            TrackBar6.Visible = True
            TrackBar7.Visible = True
            TrackBar8.Visible = True
            TrackBar9.Visible = True
            TrackBar10.Visible = True
        Else
            Label50.Visible = False
            Label51.Visible = False
            Label52.Visible = False
            Label53.Visible = False
            Label54.Visible = False
            Label55.Visible = False
            Label56.Visible = False
            Label57.Visible = False
            Label58.Visible = False
            Label59.Visible = False
            Label17.Visible = False
            Label18.Visible = False
            Label19.Visible = False
            Label20.Visible = False
            TrackBar1.Visible = False
            TrackBar2.Visible = False
            TrackBar3.Visible = False
            TrackBar4.Visible = False
            TrackBar5.Visible = False
            TrackBar6.Visible = False
            TrackBar6.Visible = False
            TrackBar7.Visible = False
            TrackBar8.Visible = False
            TrackBar9.Visible = False
            TrackBar10.Visible = False
        End If


    End Sub

    Private Sub TextBox13_TextChanged(sender As Object, e As EventArgs) Handles TextBox13.TextChanged
        GlobalVariables.EmailServer = TextBox13.Text
        FormMain.SetXMLData()

    End Sub

    Private Sub TextBox14_TextChanged(sender As Object, e As EventArgs) Handles TextBox14.TextChanged
        GlobalVariables.EmailUsername = TextBox14.Text
        FormMain.SetXMLData()
    End Sub

    Private Sub TextBox15_TextChanged(sender As Object, e As EventArgs) Handles TextBox15.TextChanged
        GlobalVariables.EmailPassword = TextBox15.Text
        FormMain.SetXMLData()

    End Sub

    Private Sub DateTimePicker1_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker1.ValueChanged
        GlobalVariables.ArchiveDate = DateTime.ParseExact(DateTimePicker1.Value.ToShortDateString, "dd/MM/yyyy", CultureInfo.CurrentCulture)
    End Sub
End Class