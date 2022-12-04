Imports System.Net
Imports System.IO
Imports System.Net.Mail
Imports System.Xml
Imports System.Security.Cryptography
Imports NetFwTypeLib
Imports System.Deployment.Application
Imports System.Reflection
Imports System.ComponentModel
Imports System.Globalization

Public Class FormMain

    Public ConnectedDB As DataBaseFunctions

    Public PacaMan As DataSet

    Dim MyMsgBox1 As New MyMsgBox
    Dim Crash As Boolean = False
    Dim dv As DataView
    Dim dvTestResults As DataView
    Dim Gridloaded As Integer = 0
    Dim Loaded As Boolean
    Dim customToolTip As New ToolTip

    Private Sub ToolStrip1_MouseHover(sender As Object, e As EventArgs) Handles ToolStrip1.MouseHover
        Me.Cursor = Cursors.Hand '
    End Sub
    Private Sub ToolStrip1_MouseLeave(sender As Object, e As EventArgs) Handles ToolStrip1.MouseLeave
        Me.Cursor = Cursors.Arrow
    End Sub
    Public Sub ErrorHandler(ex As Exception, m As String, st As String, f As StackFrame)
        Systemlog("Class: " & Replace(f.GetMethod().DeclaringType.FullName, "WindowsApplication1.", ""), "   Module: " & m, "    Line number: " & st, ex.Message)
        Actionlog("Class: " & Replace(f.GetMethod().DeclaringType.FullName, "WindowsApplication1.", "") & "   Module: " & m & "    Line number: " & st & ex.Message)
        MessageBox.Show("Class: " & Replace(f.GetMethod().DeclaringType.FullName, "WindowsApplication1.", "") & vbNewLine & "Module: " & m & vbNewLine & "Line number: " & st & vbNewLine & "Error: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)

    End Sub
    Private Sub FormMain_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        If Not Crash Then
            Try

                'Save the settings to a file
                Actionlog("Closeing save settings.xml")
                SetXMLData()

            Catch ex As Exception

                Dim st As New StackTrace(True)
                st = New StackTrace(ex, True)
                GlobalVariables.Linenumber = st.GetFrame(0).GetFileLineNumber().ToString()
                Dim f As New StackFrame
                ErrorHandler(ex, System.Reflection.MethodBase.GetCurrentMethod().Name, GlobalVariables.Linenumber, f)
            End Try
        End If
    End Sub
    Private Sub AddApp(ProfileType As NET_FW_PROFILE_TYPE_)
        Try
            Dim app As INetFwAuthorizedApplication = DirectCast(Activator.CreateInstance(Type.GetTypeFromProgID("HnetCfg.FwAuthorizedApplication")), INetFwAuthorizedApplication)
            app.Name = Application.ProductName
            app.ProcessImageFileName = Application.ExecutablePath
            app.Enabled = True
            Dim fwMgr As INetFwMgr = DirectCast(Activator.CreateInstance(Type.GetTypeFromProgID("HnetCfg.FwMgr")), INetFwMgr)
            fwMgr.LocalPolicy.GetProfileByType(ProfileType).AuthorizedApplications.Add(app)
        Catch
        End Try
    End Sub
    Public Sub CheckWebForUpdates()

        Try

            FtpFolderCreate("ftp://www.mullacottalpacas.com/wwwroot/aig/WDB_Updates/", GlobalVariables.User, GlobalVariables.Password)

            'Send error log
            Dim request As System.Net.FtpWebRequest = DirectCast(System.Net.WebRequest.Create("ftp://www.mullacottalpacas.com/wwwroot/aig/"), System.Net.FtpWebRequest)
            request.EnableSsl = False
            request.UsePassive = False
            request.Credentials = New System.Net.NetworkCredential(GlobalVariables.User, GlobalVariables.Password)
            request.Method = System.Net.WebRequestMethods.Ftp.ListDirectory
            request.KeepAlive = False

            Dim reader As New StreamReader(request.GetResponse().GetResponseStream())
            Dim line = reader.ReadLine()
            Dim lines As New List(Of String)

            Do Until line Is Nothing
                lines.Add(line)
                If CDbl(Val(Replace(Replace(line, "_", ""), ".txt", ""))) > CDbl(Val(Replace(GlobalVariables.Version, ".", ""))) Then
                    GlobalVariables.LatestUpdate = line
                    UpdateMsgBox.ShowDialog()
                    Exit Sub
                End If
                line = reader.ReadLine

            Loop
            MsgBox("No update avalable")
            Exit Sub

            Dim info As String = My.Application.Info.Version.ToString

        Catch ex As Exception
            Dim st As New StackTrace(True)
            st = New StackTrace(ex, True)
            GlobalVariables.Linenumber = st.GetFrame(0).GetFileLineNumber().ToString()
            Dim f As New StackFrame
            ErrorHandler(ex, System.Reflection.MethodBase.GetCurrentMethod().Name, GlobalVariables.Linenumber, f)
        End Try
    End Sub
    Private Sub CheckUpdates()
        Try
            'Check if this system already has a database


            Dim dir As New IO.DirectoryInfo("C:\WormCountDATA")
            If Not dir.Exists Then
                Actionlog("Backend missing")
                Dim result = MessageBox.Show("Back end database missing this may be" & vbNewLine & "due to it being moved or may be" & vbNewLine & "due to this being a new installation." & vbNewLine & "If you have moved it, press Cancel and restore the directory." & vbNewLine & "If this is a new installation press OK", "ERROR", MessageBoxButtons.OKCancel)
                If result = Windows.Forms.DialogResult.OK Then
                    'Create the directory and save the location
                    Directory.CreateDirectory("C:\WormCountDATA")
                    Directory.CreateDirectory("C:\WormCountDATA\Service")
                    SetAttr("C:\WormCountDATA\Service", vbHidden)
                    GlobalVariables.DbDriveLocation = "C:\WormCountDATA\"
                    'Now copy the system working files to this location

                    Dim myDir = Directory.GetDirectories("C:\WormCountDatabase\Application Files\").OrderByDescending(Function(f) New DirectoryInfo(f).LastWriteTime).First()


                    Dim myFile = myDir & "\" 'Directory.GetDirectories("C:\WormCountDatabase\Application Files\").OrderByDescending(Function(f) New FileInfo(f).LastWriteTime).First() & "\"
                    My.Computer.FileSystem.CopyFile(myFile & "WormCount.mdb", "C:\WormCountDATA\WormCount.mdb")
                    My.Computer.FileSystem.CopyFile(myFile & "Settings.xml", "C:\WormCountDATA\Settings.xml")
                    My.Computer.FileSystem.CopyFile(myFile & "MailServer.xml", "C:\WormCountDATA\MailServer.xml")
                    My.Computer.FileSystem.CopyFile(myFile & "Header.png", "C:\WormCountDATA\Header.png")
                    My.Computer.FileSystem.CopyFile(myFile & "Service\Install.bat", "C:\WormCountDATA\Service\Install.bat")
                    My.Computer.FileSystem.CopyFile(myFile & "InstallUtil.exe", "C:\WormCountDATA\Service\InstallUtil.exe")
                    My.Computer.FileSystem.CopyFile(myFile & "InstallUtilLib.dll", "C:\WormCountDATA\Service\InstallUtilLib.dll")
                    My.Computer.FileSystem.CopyFile(myFile & "InstallUtil.exe.config", "C:\WormCountDATA\Service\InstallUtil.exe.config")
                    My.Computer.FileSystem.CopyFile(myFile & "Service\AWC.exe", "C:\WormCountDATA\Service\AWC.exe")
                    My.Computer.FileSystem.CopyFile(myFile & "Service\AWC.exe.config", "C:\WormCountDATA\Service\AWC.exe.config")
                    My.Computer.FileSystem.CopyFile(myFile & "Service\AWC.pdb", "C:\WormCountDATA\Service\AWC.pdb")
                    My.Computer.FileSystem.CopyFile(myFile & "Service\AWC.xml", "C:\WormCountDATA\Service\AWC.xml")
                    My.Computer.FileSystem.CopyFile(myFile & "Service\AWC.exe.manifest", "C:\WormCountDATA\Service\AWC.exe.manifest")
                    My.Computer.FileSystem.CopyFile(myFile & "Service\UnInstall.bat", "C:\WormCountDATA\Service\UnInstall.bat")


                Else
                    Actionlog("Backend checked and OK")
                    Application.Exit()
                End If
            End If


        Catch ex As Exception
            Dim st As New StackTrace(True)
            st = New StackTrace(ex, True)
            GlobalVariables.Linenumber = st.GetFrame(0).GetFileLineNumber().ToString()
            Dim f As New StackFrame
            ErrorHandler(ex, System.Reflection.MethodBase.GetCurrentMethod().Name, GlobalVariables.Linenumber, f)
        End Try
    End Sub
    Public Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load


        Threading.Thread.Sleep(5000)
        Me.Opacity = 100
        Systemlog("System: " & "Load", "   Module: " & "Form1_Load", "    Line number:  " & "0", "System started")


    End Sub
    Private Sub FormMain_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        Actionlog("Program, Worm Count Database - started")

        AddApp(NET_FW_PROFILE_TYPE_.NET_FW_PROFILE_CURRENT)  'public
        AddApp(NET_FW_PROFILE_TYPE_.NET_FW_PROFILE_STANDARD) 'private  
        AddApp(NET_FW_PROFILE_TYPE_.NET_FW_PROFILE_DOMAIN)

        'GetUserAndPassword()
        Me.WindowState = FormWindowState.Maximized

        GlobalVariables.Version = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.ToString


        CheckUpdates()

        GlobalVariables.DbDriveLocation = "C:\WormCountDATA\"
        Try


            Actionlog("Program, Worm Count Database - started")
            'Read the settings from a file
            Actionlog("Open settings.xml")
            GlobalVariables.OnFarm = GetXMLData("OnFarm")
            If GlobalVariables.OnFarm < 0 Or GlobalVariables.OnFarm > 1 Then
                GlobalVariables.OnFarm = 1
            End If
            GlobalVariables.OnFarm = 1



            GlobalVariables.pass = GetXMLData("PassRate")
            GlobalVariables.Email = GetXMLData("Email")
            GlobalVariables.FarmName = GetXMLData("FarmName")
            GlobalVariables.EmailUsername = GetXMLData("EmailUsername")
            GlobalVariables.BackupNumber = GetXMLData("BackupNumber")
            GlobalVariables.ReTestTime = GetXMLData("ReTestPeriod")
            GlobalVariables.ReTestAddedTime = GetXMLData("ReTestAddedTime")
            GlobalVariables.ReminderState = GetXMLData("ReminderState")
            GlobalVariables.StartAtBoot = GetXMLData("Boot")
            GlobalVariables.PacaManpath = GetXMLData("PacaMan")
            GlobalVariables.Stoll = GetXMLData("Stoll")

            If GetXMLData("Stoll") = "True" Then
                GlobalVariables.Stoll = True
            Else
                GlobalVariables.Stoll = False
            End If
            If GlobalVariables.Stoll = True Then
                ShowStoll.Image = My.Resources.Stoll2
            Else
                ShowStoll.Image = My.Resources.Stoll1
            End If


            If GetXMLData("PacaManInstalled") = "True" Then
                GlobalVariables.PacaManInstalled = True
            Else
                GlobalVariables.PacaManInstalled = False
            End If

            If GetXMLData("ColourEnable") = "True" Then
                GlobalVariables.ColourEnable = True
            Else
                GlobalVariables.ColourEnable = False
            End If

            If GetXMLData("Messages") = "True" Then
                GlobalVariables.Messages = True
            Else
                GlobalVariables.Messages = False
            End If

            If GetXMLData("EmailPassword") <> "" Then
                Dim UnsafePassword As String = Decrypt(System.Text.Encoding.Unicode.GetBytes(GetXMLData("EmailPassword")), "Skye")
                GlobalVariables.EmailPassword = UnsafePassword
            End If
            GlobalVariables.EmailServer = GetXMLData("EmailServer")
            ToolStripButton5.Text = GlobalVariables.pass
            Try
                For i = 1 To 10
                    GlobalVariables.ColourLevel(i) = GetXMLData("ColourLevel" & i.ToString)
                Next

                If GetXMLData("BiasEnable") = "True" Then
                    GlobalVariables.BiasEnable = True
                Else
                    GlobalVariables.BiasEnable = False
                End If

                For i = 1 To 10
                    GlobalVariables.EggType(i).Name = GetXMLData("EggTypeName" & i.ToString)
                    GlobalVariables.EggType(i).Bias = GetXMLData("EggTypeBias" & i.ToString)
                Next
            Catch
            End Try

            SetXMLData()



            If Process.GetProcessesByName(Process.GetCurrentProcess.ProcessName).Length > 1 Then

                MessageBox.Show("Another Instance of this process is already running, click the icon on the task bar", "Multiple Instances Forbidden", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)

                Crash = True
                Application.Exit()

            End If

            If GetAttr("C:\WormCountDATA\Service") And Not vbHidden Then
                SetAttr("C:\WormCountDATA\Service", vbHidden)
            End If


        Catch ex As Exception

            If ex.Message <> "Application is not installed." Then
                Dim st As New StackTrace(True)
                st = New StackTrace(ex, True)
                GlobalVariables.Linenumber = st.GetFrame(0).GetFileLineNumber().ToString()
                Dim f As New StackFrame
                ErrorHandler(ex, System.Reflection.MethodBase.GetCurrentMethod().Name, GlobalVariables.Linenumber, f)
                Application.Exit()
            End If
        End Try
        'FIRST CHECK THAT THIS IS NOT  NEW INSTALLATION

        'Check that the background service is running
        Actionlog("Check service is running")
        ServiceCheck()

        'Check there is at least one animal in the Alpaca table
        ConnectedDB = New DataBaseFunctions()       'Open the database connection
        GlobalVariables.ds = ConnectedDB.PopulateDataSet()          'Put the data into a dataset
        Dim str As String = ""
        If GlobalVariables.ds.Tables("Alpaca").Rows.Count <= 0 Then
            Dim drow As DataRow = GlobalVariables.ds.Tables("Alpaca").NewRow

            str = InputBox("Enter at least one animals name." & vbNewLine & "Later more can be added with the 'Add animal' button.", "Animal name")
            If str = "" Then
                str = "Alpaca001"
            End If
            drow.Item(1) = str
            drow.Item(2) = True
            GlobalVariables.ds.Tables("Alpaca").Rows.Add(drow)
        End If


        OpenTheDataBase()




        BackgroundWorker1.RunWorkerAsync()
    End Sub
    Public Sub SendErrorReport()

        'Try
        'Dim web_upload As UploadFileToWeb = New UploadFileToWeb
        'web_upload.Upload("ftp://www.mullacottalpacas.com/wwwroot/aig/WDB/", GlobalVariables.FarmName.Replace(" ", ""))

        ' Catch ex As Exception

        'Dim st As New StackTrace(True)
        'st = New StackTrace(ex, True)
        'GlobalVariables.Linenumber = st.GetFrame(0).GetFileLineNumber().ToString()
        'Dim f As New StackFrame
        'ErrorHandler(ex, System.Reflection.MethodBase.GetCurrentMethod().Name, GlobalVariables.Linenumber, f)
        'Application.Exit()

        'End Try
    End Sub
    Public Sub OpenTheDataBase()


        'Build the data and fill the newly created datagridview
        loadFormWithAnimals()
        DataGridViewRefresh()


    End Sub
    Public Sub Systemlog(ByVal strFunction As String, ByVal strModule As String, ByVal strLine As String, ByVal strError As String)
        Try

            Using writer As New StreamWriter(GlobalVariables.DbDriveLocation & "\systemlog.txt", True)
                Dim str As String = ""
                str = DateTime.Now & vbTab
                str = str & strFunction & vbTab
                str = str & strModule & vbTab
                str = str & strLine & vbTab
                str = str & Replace(Replace(strError, vbCr, ""), vbLf, "") & vbTab
                str = str & GlobalVariables.FarmName
                writer.WriteLine(str)
                writer.Flush()
            End Using
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error")
        End Try
    End Sub
    Public Sub Actionlog(ByVal strLine As String)
        Try

            Using writer As New StreamWriter(GlobalVariables.DbDriveLocation & "\Action.log", True)
                Dim str As String = ""
                str = DateTime.Now & vbTab
                str = str & strLine.PadRight("80", " ") & vbTab
                str = str & GlobalVariables.FarmName
                writer.WriteLine(str)
                writer.Flush()
            End Using
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error")
        End Try
    End Sub
    Public Sub CreateTable()

        ConnectedDB = New DataBaseFunctions()       'Open the database connection


        Try
            ConnectedDB.UpdateDatabase("CREATE TABLE Worm (ID AUTOINCREMENT,Worm Char(40),Bias Integer)")
            ConnectedDB.UpdateDatabase("INSERT INTO Worm (Worm,Bias) VALUES ('Trichostrongyles',20)")
            ConnectedDB.UpdateDatabase("INSERT INTO Worm (Worm,Bias) VALUES ('Trichurius',50)")
            ConnectedDB.UpdateDatabase("INSERT INTO Worm (Worm,Bias) VALUES ('Nematordirus',50)")
            ConnectedDB.UpdateDatabase("INSERT INTO Worm (Worm,Bias) VALUES ('Capillarid',50)")
            ConnectedDB.UpdateDatabase("INSERT INTO Worm (Worm,Bias) VALUES ('Moniezid',50)")
            ConnectedDB.UpdateDatabase("INSERT INTO Worm (Worm,Bias) VALUES ('Emac',200)")
            ConnectedDB.UpdateDatabase("INSERT INTO Worm (Worm,Bias) VALUES ('Eivitaesis',20)")
            ConnectedDB.UpdateDatabase("INSERT INTO Worm (Worm,Bias) VALUES ('Ealpacae',10)")
            ConnectedDB.UpdateDatabase("INSERT INTO Worm (Worm,Bias) VALUES ('Elamae',10)")
            ConnectedDB.UpdateDatabase("INSERT INTO Worm (Worm,Bias) VALUES ('Epunoensis',20)")


        Catch ex As Exception

        End Try
    End Sub
    Public Sub loadFormWithAnimals()

        'Clear the datagridView's rows of data
        GlobalVariables.OnFarm = GetXMLData("OnFarm")



        Try

            If GlobalVariables.ds.Tables("TestResults").Rows.Count > 0 Then
                DataGridView1.DataSource = Nothing
            End If

        Catch ex As Exception
            Dim st As New StackTrace(True)
            st = New StackTrace(ex, True)
            GlobalVariables.Linenumber = st.GetFrame(0).GetFileLineNumber().ToString()
            Dim f As New StackFrame
            ErrorHandler(ex, System.Reflection.MethodBase.GetCurrentMethod().Name, GlobalVariables.Linenumber, f)
            Exit Sub
        End Try

        'Get the latest test group name
        ConnectedDB = New DataBaseFunctions()       'Open the database connection
        GlobalVariables.ds = ConnectedDB.PopulateDataSet()          'Put the data into a dataset
        GlobalVariables.TestGroupName = GlobalVariables.ds.Tables("TestGroup")(0)("GroupName")

        'Build the coloum names in the datagridview from the test group table
        Try
            Dim col As New DataGridViewTextBoxColumn
            col.DataPropertyName = "Name"
            col.HeaderText = "Name"
            col.Name = "Name"
            col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            col.DefaultCellStyle.Font = New Font("Tahoma", 12, FontStyle.Regular)
            col.Frozen = True
            DataGridView1.Columns.Add(col)

            'If no columns at all add at least one.
            If GlobalVariables.ds.Tables("TestGroup").Rows.Count <= 0 Then
                Dim drow1 As DataRow = GlobalVariables.ds.Tables("TestGroup").NewRow
                Dim iMonth As Integer = Month(DateTime.Now)
                drow1.Item(1) = MonthName(iMonth) + " " + Year(DateTime.Now).ToString
                drow1.Item(2) = CDate(Now.ToShortDateString)
                GlobalVariables.ds.Tables("TestGroup").Rows.Add(drow1)

                Dim colTestGroup As New DataGridViewTextBoxColumn
                colTestGroup.DataPropertyName = drow1.Item(1)
                colTestGroup.HeaderText = drow1.Item(1)
                colTestGroup.Name = drow1.Item(1)
                colTestGroup.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                DataGridView1.Columns.Add(colTestGroup)
                Exit Sub
            End If


            For Each GroupRow As DataRow In GlobalVariables.ds.Tables("TestGroup").Rows
                GlobalVariables.ds.Tables("GridView").Columns.Add(New DataColumn(GroupRow.Item(1)))
                Dim colGroupRow As New DataGridViewTextBoxColumn
                colGroupRow.DataPropertyName = GroupRow.Item(1)
                colGroupRow.HeaderText = GroupRow.Item(1)
                colGroupRow.Name = GroupRow.Item(1)
                colGroupRow.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                DataGridView1.Columns.Add(colGroupRow)
            Next




            ' ToolStripStatusLabel2.Text = "Version number " & My.Application.Deployment.CurrentVersion.ToString() & "    | "
        Catch ex As Exception
            If ex.Message <> "Application is not installed." Then
                Dim st As New StackTrace(True)
                st = New StackTrace(ex, True)
                GlobalVariables.Linenumber = st.GetFrame(0).GetFileLineNumber().ToString()
                Dim f As New StackFrame
                ErrorHandler(ex, System.Reflection.MethodBase.GetCurrentMethod().Name, GlobalVariables.Linenumber, f)
            End If
        End Try

        FillDataView()


    End Sub
    Public Sub FillDataView()
        Dim str As String = ""
        Dim max As Integer
        Dim insertValue As String = ""
        Dim CurrentDate As String = ""
        Dim lastDate As String = "01/01/1970 00:00:00"
        Try
            'Find out how many tests have been done
            Dim NumberOfTestDone As Integer = 0
            Dim NumberOfGroup As String = ""

            'Check if each animal has had a test and if there is a group for it if not add a group test coloum in the datagridview
            For Each AnimalRow As DataRow In GlobalVariables.ds.Tables("Alpaca").Rows
                NumberOfTestDone = 0

                For Each TestRow As DataRow In GlobalVariables.ds.Tables("TestResults").Rows
                    'Check to see if this is a new test if so create a new coloumn
                    If Not GlobalVariables.ds.Tables("GridView").Columns.Contains(TestRow.Item(5)) Then
                        'Create the new test coloumn
                        GlobalVariables.ds.Tables("GridView").Columns.Add(New DataColumn(TestRow.Item(5)))
                        Dim col As New DataGridViewTextBoxColumn
                        col.DataPropertyName = TestRow.Item(5)
                        col.HeaderText = TestRow.Item(5)
                        col.Name = TestRow.Item(5)
                        col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                        DataGridView1.Columns.Add(col)
                    End If


                    'Check to see if the entry is for the current aniamal if so fill the test result
                    If TestRow.Item(2) = AnimalRow(1) Then
                        'now run through the database again and see if the animal has more that one entry per group date.
                        NumberOfTestDone = 0
                        Dim foundRow() As DataRow
                        foundRow = GlobalVariables.ds.Tables("TestResults").Select("TestName='" & AnimalRow(1) & "' and TestNumberName='" & TestRow.Item(5) & "'")
                        NumberOfTestDone = foundRow.Length

                        'Find the max result for this period then add a Chr(2) at the end to be used latter in the formatting
                        max = 0
                        For i = 0 To foundRow.Length - 1
                            If Convert.ToInt16((foundRow(i).ItemArray(6)) + Convert.ToInt16(foundRow(i).ItemArray(7))) > max Then
                                max = Convert.ToInt16(foundRow(i).ItemArray(6)) + Convert.ToInt16(foundRow(i).ItemArray(7))
                            End If
                        Next

                        Try
                            If TestRow.Item(2) = "Krystal" And AnimalRow(1) = "Krystal" Then
                                AnimalRow(1) = AnimalRow(1)
                            End If
                            'If TestRow.Item(1) > lastDate Then 'Check that this is the latest test
                            If GlobalVariables.ds.Tables("GridView").Rows(AnimalRow.Table.Rows.IndexOf(AnimalRow)).Item(TestRow.Item(5)).ToString = "" Then
                                lastDate = TestRow.Item(1)
                                insertValue = TestRow(6) & " " & TestRow(7)
                                If NumberOfTestDone > 1 Then
                                    insertValue = insertValue & "*"
                                End If
                                If max >= GlobalVariables.pass Then
                                    insertValue = insertValue & Chr(2)
                                End If
                                GlobalVariables.ds.Tables("GridView").Rows(AnimalRow.Table.Rows.IndexOf(AnimalRow)).Item(TestRow.Item(5)) = insertValue
                            End If


                        Catch ex As Exception
                            Dim st As New StackTrace(True)
                            st = New StackTrace(ex, True)
                            GlobalVariables.Linenumber = st.GetFrame(0).GetFileLineNumber().ToString()
                            Dim f As New StackFrame
                            ErrorHandler(ex, System.Reflection.MethodBase.GetCurrentMethod().Name, GlobalVariables.Linenumber, f)
                        End Try
                    End If
                Next
                lastDate = "01/01/1970 00:00:00"
            Next

        Catch ex As Exception

            If ex.Message <> "Application is not installed." Then
                Dim st As New StackTrace(True)
                st = New StackTrace(ex, True)
                GlobalVariables.Linenumber = st.GetFrame(0).GetFileLineNumber().ToString()
                Dim f As New StackFrame
                ErrorHandler(ex, System.Reflection.MethodBase.GetCurrentMethod().Name, GlobalVariables.Linenumber, f)
            End If
        End Try


        'DataGridViewRefresh()

    End Sub
    Private Function ColourCell(row As Integer)
        Dim EggBiased As Integer = 0
        Dim Cell1 As Integer = 0
        Dim Cell2 As Integer = 0
        If row > 0 Then

            Dim ID As String = GlobalVariables.ds.Tables("TestResults").Rows(row).Item(0)
            Dim Animal As String = GlobalVariables.ds.Tables("TestResults").Rows(row).Item(2)
            Dim dt As String = GlobalVariables.ds.Tables("TestResults").Rows(row).Item(1)
            Dim Trichostrongyles As String = GlobalVariables.ds.Tables("TestResults").Rows(row).Item(8)
            Dim Trichurius As String = GlobalVariables.ds.Tables("TestResults").Rows(row).Item(9)
            Dim Nematordirus As String = GlobalVariables.ds.Tables("TestResults").Rows(row).Item(10)
            Dim Capillarid As String = GlobalVariables.ds.Tables("TestResults").Rows(row).Item(11)
            Dim Moniezid As String = GlobalVariables.ds.Tables("TestResults").Rows(row).Item(12)
            Dim Emac As String = GlobalVariables.ds.Tables("TestResults").Rows(row).Item(14)
            Dim Eivitaesis As String = GlobalVariables.ds.Tables("TestResults").Rows(row).Item(15)
            Dim Ealpacae As String = GlobalVariables.ds.Tables("TestResults").Rows(row).Item(16)
            Dim Elamae As String = GlobalVariables.ds.Tables("TestResults").Rows(row).Item(17)
            Dim Epunoensis As String = GlobalVariables.ds.Tables("TestResults").Rows(row).Item(18)

            EggBiased = GetEggs(Trichostrongyles)
            EggBiased = EggBiased + GetEggs(Trichurius)
            EggBiased = EggBiased + GetEggs(Nematordirus)
            EggBiased = EggBiased + GetEggs(Capillarid)
            EggBiased = EggBiased + GetEggs(Moniezid)
            EggBiased = EggBiased + GetEggs(Emac)
            EggBiased = EggBiased + GetEggs(Eivitaesis)
            EggBiased = EggBiased + GetEggs(Ealpacae)
            EggBiased = EggBiased + GetEggs(Elamae)
            EggBiased = EggBiased + GetEggs(Epunoensis)


            If (EggBiased * 50) >= GlobalVariables.pass Then
                ColourCell = Color.Pink
            Else
                ColourCell = Color.PaleGreen
            End If


            Exit Function
        End If

        ColourCell = Color.PaleGreen
    End Function
    Public Function GetEggs(egg As String)
        While egg > 0
            GetEggs += egg Mod 10
            egg = (egg / 10) - ((egg Mod 10) / 10)
        End While

    End Function
    Public Function GetXMLData(ByVal Tag As String) As String
        Dim ret As String = ""
        Try
            Using Reader As XmlReader = XmlReader.Create(GlobalVariables.DbDriveLocation & "Settings.xml")
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
                ErrorHandler(ex, System.Reflection.MethodBase.GetCurrentMethod().Name, GlobalVariables.Linenumber, f)
            End If
        End Try

        Return ret
    End Function
    Public Function Encrypt(ByVal plainText As String, ByVal secretKey As String) As Byte()
        Dim encryptedPassword As Byte()
        Using outputStream As MemoryStream = New MemoryStream()
            Dim algorithm As RijndaelManaged = getAlgorithm(secretKey)
            Using cryptoStream As CryptoStream = New CryptoStream(outputStream, algorithm.CreateEncryptor(), CryptoStreamMode.Write)
                Dim inputBuffer() As Byte = System.Text.Encoding.Unicode.GetBytes(plainText)
                cryptoStream.Write(inputBuffer, 0, inputBuffer.Length)
                cryptoStream.FlushFinalBlock()
                encryptedPassword = outputStream.ToArray()
            End Using
        End Using
        Return encryptedPassword
    End Function
    Public Function Decrypt(ByVal encryptedBytes As Byte(), ByVal secretKey As String) As String
        Dim plainText As String = Nothing
        Using inputStream As MemoryStream = New MemoryStream(encryptedBytes)
            Dim algorithm As RijndaelManaged = getAlgorithm(secretKey)
            Using cryptoStream As CryptoStream = New CryptoStream(inputStream, algorithm.CreateDecryptor(), CryptoStreamMode.Read)
                Dim outputBuffer(0 To CType(inputStream.Length - 1, Integer)) As Byte
                Dim readBytes As Integer = cryptoStream.Read(outputBuffer, 0, CType(inputStream.Length, Integer))
                plainText = System.Text.Encoding.Unicode.GetString(outputBuffer, 0, readBytes)
            End Using
        End Using
        Return plainText
    End Function
    Private Function getAlgorithm(ByVal secretKey As String) As RijndaelManaged
        Const salt As String = "sdgfsdhflfhiljf"
        Const keySize As Integer = 256

        Dim keyBuilder As Rfc2898DeriveBytes = New Rfc2898DeriveBytes(secretKey, System.Text.Encoding.Unicode.GetBytes(salt))
        Dim algorithm As RijndaelManaged = New RijndaelManaged()
        algorithm.KeySize = keySize
        algorithm.IV = keyBuilder.GetBytes(CType(algorithm.BlockSize / 8, Integer))
        algorithm.Key = keyBuilder.GetBytes(CType(algorithm.KeySize / 8, Integer))
        algorithm.Padding = PaddingMode.PKCS7
        Return algorithm
    End Function
    Public Function SetXMLData() As String
        Dim ret As String = ""


        Try
            Dim settings As XmlWriterSettings = New XmlWriterSettings()
            settings.Indent = True

            ' Create XmlWriter.
            Using writer As XmlWriter = XmlWriter.Create(GlobalVariables.DbDriveLocation & "Settings.xml", settings)
                ' Begin writing.
                writer.WriteStartDocument()
                writer.WriteStartElement("Settings") ' Root.

                writer.WriteStartElement("Global")
                If GlobalVariables.FarmName = "" Then
                    GlobalVariables.FarmName = "New sample Farm"
                End If
                writer.WriteElementString("FarmName", GlobalVariables.FarmName)
                writer.WriteElementString("OnFarm", GlobalVariables.OnFarm)
                writer.WriteElementString("PassRate", GlobalVariables.pass)
                writer.WriteElementString("PacaMan", GlobalVariables.PacaManpath)
                writer.WriteElementString("PacaManInstalled", GlobalVariables.PacaManInstalled)
                writer.WriteElementString("Stoll", GlobalVariables.Stoll)
                writer.WriteElementString("ColourEnable", GlobalVariables.ColourEnable)
                If GlobalVariables.Email = "" Then
                    GlobalVariables.Email = "sample@farm.com"
                End If
                writer.WriteElementString("Email", GlobalVariables.Email)
                If GlobalVariables.EmailUsername = "" Then
                    GlobalVariables.EmailUsername = "User@hotmail.com"
                End If
                writer.WriteElementString("EmailUsername", GlobalVariables.EmailUsername)

                If GlobalVariables.EmailPassword <> "" Then
                    Dim SafePassword() As Byte = Encrypt(GlobalVariables.EmailPassword, "Skye")
                    writer.WriteElementString("EmailPassword", System.Text.Encoding.Unicode.GetString(SafePassword))
                End If
                If GlobalVariables.EmailServer = "" Then
                    GlobalVariables.EmailServer = "smtp-mail.outlook.com"
                End If


                writer.WriteElementString("EmailServer", GlobalVariables.EmailServer)
                writer.WriteElementString("BackupNumber", GlobalVariables.BackupNumber)
                writer.WriteElementString("ReTestPeriod", GlobalVariables.ReTestTime)
                writer.WriteElementString("ReTestAddedTime", GlobalVariables.ReTestAddedTime)
                writer.WriteElementString("Messages", GlobalVariables.Messages)
                writer.WriteElementString("ReminderState", GlobalVariables.ReminderState)
                writer.WriteElementString("Max", GlobalVariables.StartMax)
                writer.WriteElementString("Boot", GlobalVariables.StartAtBoot)
                writer.WriteElementString("BackEndLocation", GlobalVariables.DbDriveLocation)
                For i = 1 To 10
                    writer.WriteElementString("ColourLevel" & CInt(i), GlobalVariables.ColourLevel(i))
                Next

                writer.WriteElementString("BiasEnable", GlobalVariables.BiasEnable)
                writer.WriteElementString("EggTypeName1", "Trichostrongyles")
                writer.WriteElementString("EggTypeName2", "Trichurius")
                writer.WriteElementString("EggTypeName3", "Nematordirus")
                writer.WriteElementString("EggTypeName4", "Capillarid")
                writer.WriteElementString("EggTypeName5", "Moniezid")
                writer.WriteElementString("EggTypeName6", "E-mac")
                writer.WriteElementString("EggTypeName7", "E-ivitaesis")
                writer.WriteElementString("EggTypeName8", "E-alpacae")
                writer.WriteElementString("EggTypeName9", "E-lamae")
                writer.WriteElementString("EggTypeName10", "E-punoensis")

                For i = 1 To 10
                    writer.WriteElementString("EggTypeBias" & CInt(i), GlobalVariables.EggType(i).Bias)
                Next


                writer.WriteEndElement()




                ' End document.
                writer.WriteEndElement()
                writer.WriteEndDocument()
            End Using
        Catch ex As Exception
            Dim st As New StackTrace(True)
            st = New StackTrace(ex, True)
            GlobalVariables.Linenumber = st.GetFrame(0).GetFileLineNumber().ToString()
            Dim f As New StackFrame
            ErrorHandler(ex, System.Reflection.MethodBase.GetCurrentMethod().Name, GlobalVariables.Linenumber, f)

        End Try


        Return ret
    End Function
    Public Shared Function ResizeImage(ByVal InputImage As Image) As Image
        Return New Bitmap(InputImage, New Size(371, 107))
    End Function
    Private Sub DataGridView1_CellFormatting(ByVal sender As Object, ByVal e As DataGridViewCellFormattingEventArgs) Handles DataGridView1.CellFormatting
        If Loaded Then
            If e.ColumnIndex > 0 Then

                If Not (IsDBNull(e.Value)) And Not (e.Value Is Nothing) Then
                    Try
                        Dim tempWord As String = e.Value.ToString.Replace("*", "")
                        Dim words As String() = tempWord.Split(New Char() {" "c})
                        If e.Value.ToString.Contains(Chr(2)) Then
                            e.CellStyle.BackColor = Color.Pink
                        Else
                            e.CellStyle.BackColor = Color.PaleGreen
                        End If


                    Catch ex As Exception
                        Dim st As New StackTrace(True)
                        st = New StackTrace(ex, True)
                        GlobalVariables.Linenumber = st.GetFrame(0).GetFileLineNumber().ToString()
                        Dim f As New StackFrame
                        ErrorHandler(ex, System.Reflection.MethodBase.GetCurrentMethod().Name, GlobalVariables.Linenumber, f)
                    End Try
                End If

            End If
        End If
    End Sub
    Private Sub FormatTheAnimalCol(ByVal sender As Object, ByVal e As DataGridViewCellFormattingEventArgs)
        Dim Total As Integer = 0
        Dim Tests As Integer = 0

        For j = 0 To GlobalVariables.ds.Tables("TestResults").Rows.Count - 1
            If GlobalVariables.ds.Tables("TestResults").Rows(j)("TestName") = e.Value.ToString Then
                Total = Total + Convert.ToInt32(GlobalVariables.ds.Tables("TestResults").Rows(j)("TestEPGTotal")) + Convert.ToInt32(GlobalVariables.ds.Tables("TestResults").Rows(j)("TestOPGTotal"))
                Tests += 1
            End If
        Next
        e.CellStyle.BackColor = Color.White
        If GlobalVariables.ColourEnable Then
            If (Total / Tests) > GlobalVariables.ColourLevel(10) Then
                e.CellStyle.BackColor = Color.FromArgb(255, 50, 50)
            ElseIf (Total / Tests) > GlobalVariables.ColourLevel(9) Then
                e.CellStyle.BackColor = Color.FromArgb(255, 75, 75)
            ElseIf (Total / Tests) > GlobalVariables.ColourLevel(8) Then
                e.CellStyle.BackColor = Color.FromArgb(255, 100, 100)
            ElseIf (Total / Tests) > GlobalVariables.ColourLevel(7) Then
                e.CellStyle.BackColor = Color.FromArgb(255, 125, 125)
            ElseIf (Total / Tests) > GlobalVariables.ColourLevel(6) Then
                e.CellStyle.BackColor = Color.FromArgb(255, 150, 150)
            ElseIf (Total / Tests) > GlobalVariables.ColourLevel(5) Then
                e.CellStyle.BackColor = Color.FromArgb(255, 175, 175)
            ElseIf (Total / Tests) > GlobalVariables.ColourLevel(4) Then
                e.CellStyle.BackColor = Color.FromArgb(255, 200, 200)
            ElseIf (Total / Tests) > GlobalVariables.ColourLevel(3) Then
                e.CellStyle.BackColor = Color.FromArgb(255, 220, 220)
            ElseIf (Total / Tests) > GlobalVariables.ColourLevel(2) Then
                e.CellStyle.BackColor = Color.FromArgb(255, 240, 240)
            ElseIf (Total / Tests) > GlobalVariables.ColourLevel(1) Then
                e.CellStyle.BackColor = Color.FromArgb(255, 253, 253)
            End If
        End If

    End Sub
    Private Sub DataGridView1_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellDoubleClick
        If Loaded Then
            'Actionlog("DataGridView1.CellDoubleClick")
            Try
                Dim res As DialogResult
                GlobalVariables.AlpacaName = DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells("Name").Value.ToString


                If DataGridView1.CurrentCell.ColumnIndex <> 0 And DataGridView1.CurrentCell.RowIndex <> -1 And e.RowIndex <> -1 Then
                    Dim result As String = ""
                    GlobalVariables.Clickgroup = ""
                    GlobalVariables.Clickgroup = sender.columns(e.ColumnIndex).headercell.value
                    If DataGridView1.Rows(e.RowIndex).Cells(e.ColumnIndex).Value.ToString = "" Then
                        Actionlog("Double clicked in the data")
                        If GlobalVariables.Stoll Then
                            Form3.ShowDialog()
                        Else
                            FormTest.ShowDialog()
                        End If
                        Exit Sub
                    End If

                    result = DataGridView1.Rows(e.RowIndex).Cells(e.ColumnIndex).Value


                    MyMsgBox1.Label1.Text = GlobalVariables.AlpacaName & " has been tested, is this new test data " & vbNewLine & "to be entered or do you want to see the last test reaults."
                    MyMsgBox1.ShowDialog()
                    res = MyMsgBox1.DialogResult

                    If res = Windows.Forms.DialogResult.Yes Then
                        Actionlog("Clicked Yes new data")
                        If GlobalVariables.Stoll Then
                            Form3.ShowDialog()
                        Else
                            FormTest.ShowDialog()
                        End If
                    End If

                    If res = Windows.Forms.DialogResult.No Then
                        GlobalVariables.AlpacaTestDate = sender.columns(e.ColumnIndex).headercell.value
                        Actionlog("Clicked No load history data")
                        FindIfStoll()
                    End If
                Else
                    If e.RowIndex <> -1 Then
                        Actionlog("Clicked on the animal name")
                        FormAlpaca.ShowDialog()
                    End If
                End If
            Catch ex As Exception

                If ex.Message = "Object reference not set to an instance of an object." Then
                    If GlobalVariables.Stoll Then
                        Form3.ShowDialog()
                    Else
                        FormTest.ShowDialog()
                    End If
                    Exit Sub
                End If
                Dim st As New StackTrace(ex, True)
                GlobalVariables.Linenumber = st.GetFrame(0).GetFileLineNumber().ToString()
                Dim f As New StackFrame
                ErrorHandler(ex, System.Reflection.MethodBase.GetCurrentMethod().Name, GlobalVariables.Linenumber, f)
            End Try
        End If

    End Sub
    Private Sub FindIfStoll()
        For j = 0 To GlobalVariables.ds.Tables("TestResults").Rows.Count - 1
            If GlobalVariables.ds.Tables("TestResults").Rows(j)("TestNumberName") = GlobalVariables.AlpacaTestDate Then
                If GlobalVariables.ds.Tables("TestResults").Rows(j)("TestName") = GlobalVariables.AlpacaName Then
                    Try
                        Dim Stoll As String = GlobalVariables.ds.Tables("TestResults").Rows(j)("Name")
                        If Stoll = "Stoll" Then
                            FormTestHistoryStoll.ShowDialog()
                            Exit For
                        Else
                            FormTestHistory.ShowDialog()
                            Exit For
                        End If

                    Catch
                        FormTestHistory.ShowDialog()
                        Exit For
                    End Try
                End If
                End If
        Next
    End Sub
    Private Sub CheckBox1_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        FillDataView()
    End Sub
    Private Sub ExitToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Me.Close()
    End Sub
    Private Sub AboutToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        FormAbout.ShowDialog()
    End Sub
    Private Sub NewAnimalToolStripMenuItem_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs)
        FormNewAnimal.ShowDialog()

    End Sub
    Private Sub CreateNewTestGroupToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        NewTestPeriod()
    End Sub
    Private Sub EPGOPGPassRateToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        FormPass.ShowDialog()

    End Sub
    Private Sub ShowOnlyOnFarm_Click(sender As Object, e As EventArgs) Handles ShowOnlyOnFarm.Click

        OnFarm()
        DataGridViewRefresh()
        Actionlog("On farm tool bar pressed - OnFarm = " & GlobalVariables.OnFarm)
    End Sub
    Public Sub OnFarm()
        If GlobalVariables.OnFarm < 0 Or GlobalVariables.OnFarm > 1 Then
            GlobalVariables.OnFarm = 1

        End If
        If GlobalVariables.OnFarm = 1 Then
            GlobalVariables.OnFarm = 0
            SetXMLData()
            If GlobalVariables.OnFarm = 1 Then
                ShowOnlyOnFarm.Image = My.Resources.OnFarmIco
            Else
                ShowOnlyOnFarm.Image = My.Resources.OffFarmIco
            End If
            Exit Sub
        End If
        If GlobalVariables.OnFarm = 0 Then
            GlobalVariables.OnFarm = 1
            SetXMLData()
        End If
        If GlobalVariables.OnFarm = 1 Then
            ShowOnlyOnFarm.Image = My.Resources.OnFarmIco
        Else
            ShowOnlyOnFarm.Image = My.Resources.OffFarmIco
        End If
    End Sub
    Private Sub NewTestPeriod()
        Dim result As DialogResult

        ConnectedDB = New DataBaseFunctions()    'Open the database connection
        GlobalVariables.ds = ConnectedDB.PopulateDataSet()       'Put the datainto a dataset
        If GlobalVariables.ds.Tables("TestGroup").Rows.Count Then
            Dim str As String = GlobalVariables.ds.Tables("TestGroup").Rows(0)("GroupName")
            result = MessageBox.Show("This will close the 'worm count test period' called -- " & str & vbNewLine & vbNewLine & "Then let you create a new 'worm count test period'." & vbNewLine & "Do you wish to continue?", "Warning", MessageBoxButtons.YesNo)

            If (result = DialogResult.Yes) Then
                FormNewTestGroup.ShowDialog()
                DataGridViewRefresh()
            End If
        Else
            FormNewTestGroup.ShowDialog()
            DataGridViewRefresh()
        End If
    End Sub
    Private Sub DataGridView1_CellMouseEnter(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellMouseEnter
        If Loaded Then
            Try
                Dim test As Integer = 0
                Dim str As String
                Dim count As String
                customToolTip.ShowAlways = True
                customToolTip.InitialDelay = 1
                customToolTip.OwnerDraw = True
                customToolTip.UseFading = True
                customToolTip.ReshowDelay = 500
                customToolTip.AutoPopDelay = 10000
                customToolTip.UseFading = True
                customToolTip.ToolTipTitle = "Test summery"
                customToolTip.ToolTipIcon = ToolTipIcon.Info
                customToolTip.IsBalloon = True


                'Check to see if the hovered cell contains ***'s
                If DataGridView1.Rows.Count > 0 And e.RowIndex >= 0 And e.ColumnIndex > 0 Then
                    If Not DataGridView1.Rows(e.RowIndex).Cells(e.ColumnIndex).Value Is Nothing Then
                        str = "Animal name:    " & DataGridView1.Rows(e.RowIndex).Cells("Name").Value.ToString & vbNewLine
                        str = str & "     Test period:   " & DataGridView1.Columns(e.ColumnIndex).Name & vbNewLine & vbNewLine
                        str = str & "_________________________________________________________________________________________________" & vbNewLine
                        'Run through all of the test results  
                        ConnectedDB = New DataBaseFunctions()     'Open the database connection
                        GlobalVariables.ds = ConnectedDB.PopulateDataSet()        'Put the datainto a dataset
                        Dim result() As DataRow = GlobalVariables.ds.Tables("TestResults").Select("TestName='" & DataGridView1.Rows(e.RowIndex).Cells("Name").Value.ToString & "' and TestNumberName='" & DataGridView1.Columns(e.ColumnIndex).Name & "'")
                        For j = result.Count - 1 To 0 Step -1
                            test = test + 1
                            count = (Convert.ToInt32(result(j).ItemArray(6)) + Convert.ToInt32(result(j).ItemArray(7))).ToString
                            While Len(count) < 4
                                count = "-" & count
                            End While

                            str = str & "Test " & test.ToString & ":   " & count & "  " & EggType(result, j) & vbNewLine
                        Next
                        If DataGridView1.Rows(e.RowIndex).Cells(e.ColumnIndex).Value.ToString <> "" And DataGridView1.Rows(e.RowIndex).Cells(e.ColumnIndex).Value.ToString <> "0" Then

                            customToolTip.SetToolTip(DataGridView1, str & vbNewLine & vbNewLine & vbNewLine & vbNewLine & "NOTE: '*' means the test period has more than one test.")

                        End If
                    End If
                End If
            Catch ex As Exception
                Dim st As New StackTrace(True)
                st = New StackTrace(ex, True)
                GlobalVariables.Linenumber = st.GetFrame(0).GetFileLineNumber().ToString()
                Dim f As New StackFrame
                ErrorHandler(ex, System.Reflection.MethodBase.GetCurrentMethod().Name, GlobalVariables.Linenumber, f)
            End Try
        End If
    End Sub
    Private Sub DataGridView1_CellMouseLeave(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellMouseLeave
        customToolTip.SetToolTip(DataGridView1, "")
    End Sub
    Public Function EggType(List() As DataRow, RowNumber As Integer)
        Dim strEggs As String = ""
        Dim Eggs As Int16 = 0
        Dim str As String = ""
        If Convert.ToInt16(List(RowNumber).ItemArray(8)) > 0 Then
            Eggs = List(RowNumber).ItemArray(8).ToString.Substring(0, 2)
            Eggs = Eggs + List(RowNumber).ItemArray(8).ToString.Substring(2, 2)
            strEggs = AddEggString(str, strEggs, Eggs, " Trichostrongyles   ")
        End If

        If Convert.ToInt16(List(RowNumber).ItemArray(9)) > 0 Then
            Eggs = List(RowNumber).ItemArray(9).ToString.Substring(0, 2)
            Eggs = Eggs + List(RowNumber).ItemArray(9).ToString.Substring(2, 2)
            strEggs = AddEggString(str, strEggs, Eggs, " Trichurius   ")
        End If

        If Convert.ToInt16(List(RowNumber).ItemArray(10)) > 0 Then
            Eggs = List(RowNumber).ItemArray(10).ToString.Substring(0, 2)
            Eggs = Eggs + List(RowNumber).ItemArray(10).ToString.Substring(2, 2)
            strEggs = AddEggString(str, strEggs, Eggs, " Nematordirus   ")
        End If

        If Convert.ToInt16(List(RowNumber).ItemArray(11)) > 0 Then
            Eggs = List(RowNumber).ItemArray(11).ToString.Substring(0, 2)
            Eggs = Eggs + List(RowNumber).ItemArray(11).ToString.Substring(2, 2)
            strEggs = AddEggString(str, strEggs, Eggs, " Capillarid   ")
        End If

        If Convert.ToInt16(List(RowNumber).ItemArray(12)) > 0 Then
            Eggs = List(RowNumber).ItemArray(12).ToString.Substring(0, 2)
            Eggs = Eggs + List(RowNumber).ItemArray(12).ToString.Substring(2, 2)
            strEggs = AddEggString(str, strEggs, Eggs, " Moniezid   ")
        End If

        If Convert.ToInt16(List(RowNumber).ItemArray(13)) > 0 Then
            Eggs = List(RowNumber).ItemArray(13).ToString.Substring(0, 2)
            Eggs = Eggs + List(RowNumber).ItemArray(13).ToString.Substring(2, 2)
            strEggs = AddEggString(str, strEggs, Eggs, " Unidentifed   ")
        End If

        If Convert.ToInt16(List(RowNumber).ItemArray(14)) > 0 Then
            Eggs = List(RowNumber).ItemArray(14).ToString.Substring(0, 2)
            Eggs = Eggs + List(RowNumber).ItemArray(14).ToString.Substring(2, 2)
            strEggs = AddEggString(str, strEggs, Eggs, " E-mac   ")
        End If

        If Convert.ToInt16(List(RowNumber).ItemArray(15)) > 0 Then
            Eggs = List(RowNumber).ItemArray(15).ToString.Substring(0, 2)
            Eggs = Eggs + List(RowNumber).ItemArray(15).ToString.Substring(2, 2)
            strEggs = AddEggString(str, strEggs, Eggs, " E-ivitaesis   ")
        End If

        If Convert.ToInt16(List(RowNumber).ItemArray(16)) > 0 Then
            Eggs = List(RowNumber).ItemArray(16).ToString.Substring(0, 2)
            Eggs = Eggs + List(RowNumber).ItemArray(16).ToString.Substring(2, 2)
            strEggs = AddEggString(str, strEggs, Eggs, " E-alpacae   ")
        End If

        If Convert.ToInt16(List(RowNumber).ItemArray(17)) > 0 Then
            Eggs = List(RowNumber).ItemArray(17).ToString.Substring(0, 2)
            Eggs = Eggs + List(RowNumber).ItemArray(17).ToString.Substring(2, 2)
            strEggs = AddEggString(str, strEggs, Eggs, " E-lamae   ")
        End If

        If Convert.ToInt16(List(RowNumber).ItemArray(18)) > 0 Then
            Eggs = List(RowNumber).ItemArray(18).ToString.Substring(0, 2)
            Eggs = Eggs + List(RowNumber).ItemArray(18).ToString.Substring(2, 2)
            strEggs = AddEggString(str, strEggs, Eggs, " E-punoensis   ")
        End If

        If Convert.ToInt16(List(RowNumber).ItemArray(19)) > 0 Then
            Eggs = List(RowNumber).ItemArray(19).ToString.Substring(0, 2)
            Eggs = Eggs + List(RowNumber).ItemArray(19).ToString.Substring(2, 2)
            strEggs = AddEggString(str, strEggs, Eggs, " Unidentifed   ")
        End If
        EggType = strEggs
    End Function
    Private Function AddEggString(str As String, strEggs As String, Eggs As Integer, name As String)
        Dim strLength As Integer = (Len(str) + Len(strEggs) + Len(name))
        strEggs = strEggs & Eggs.ToString & name

        AddEggString = strEggs
    End Function
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        FormGraphGroup.Show()
    End Sub
    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs)
        loadFormWithAnimals()
    End Sub
    Private Sub Button2_Click_1(sender As System.Object, e As System.EventArgs)
        FormGroupTestReport.ShowDialog()
    End Sub
    Private Sub ToolStripButton1_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton1.Click
        Actionlog("Farm report tool bar pressed")
        FormGroupTestReport.ShowDialog()
    End Sub
    Private Sub ToolStripButton2_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton2.Click
        Actionlog("Add animal tool bar pressed")
        FormNewAnimal.ShowDialog()

        DataGridViewRefresh()
        loadFormWithAnimals()

    End Sub
    Private Sub ToolStripButton3_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton3.Click
        Actionlog("New test period tool bar pressed")
        NewTestPeriod()
    End Sub
    Public Sub DataGridViewRefresh()
        Dim DGVHScroll = DataGridView1.Controls.OfType(Of HScrollBar).SingleOrDefault
        Try



            Gridloaded += 1
            'If Data Then has been lost Or corupted In the DataGridView rebuild it
            If GlobalVariables.ds.Tables("GridView").Columns.Count <= 3 And Loaded Then
                Threading.Thread.Sleep(1000)
                DataGridView1.Columns.Clear()
                DataGridView1.DataSource = Nothing
                GlobalVariables.ds.Dispose()
                GlobalVariables.Linenumber = 1040
                Dim f As New StackFrame
                Systemlog("Class: " & "FormMain", "   Module: " & "DataGridViewRefresh", "    Line number: " & GlobalVariables.Linenumber, "Database lost all its data again???")
                loadFormWithAnimals()
                'MsgBox("opps something went wrong try --- press the refresh button above", MsgBoxStyle.Critical, "Opps")
            End If

            'Threading.Thread.Sleep(1000)
            Dim dv = GlobalVariables.ds.Tables("GridView")
            dv.DefaultView.RowFilter = "OnFarm = " & GlobalVariables.OnFarm
            DataGridView1.DataSource = dv
            DataGridView1.Sort(DataGridView1.Columns(1), ComponentModel.ListSortDirection.Ascending)
            DataGridView1.HorizontalScrollingOffset = 5000
            DataGridView1.Columns("Name").Width = 190
            DataGridView1.Columns("OnFarm").Visible = False
            DataGridView1.Columns("ID").Visible = False
            ToolStripStatusLabel4.Text = "Animal count  -  " & DataGridView1.Rows.Count.ToString & " | "
            DataGridView1.Sort(DataGridView1.Columns(0), System.ComponentModel.ListSortDirection.Ascending)  'Sort the animals into alphbetical order
            DataGridView1.DefaultCellStyle.Font = New Font("Tahoma", 10)
            DataGridView1.ColumnHeadersDefaultCellStyle.Font = New Font("Tahoma", 10)
            Loaded = True


            'Check that there are no missing coulumns
            For Each Row As DataRow In GlobalVariables.ds.Tables("TestGroup").Rows
                If Not GlobalVariables.ds.Tables("GridView").Columns.Contains(Row.Item(1)) Then
                    'Create the new test coloumn
                    GlobalVariables.ds.Tables("GridView").Columns.Add(New DataColumn(Row.Item(1)))
                    Dim col As New DataGridViewTextBoxColumn
                    col.DataPropertyName = Row.Item(1)
                    col.HeaderText = Row.Item(1)
                    col.Name = Row.Item(1)
                    col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                End If
            Next

            DataGridView1.HorizontalScrollingOffset = 0
            Exit Sub
        Catch ex As Exception
            Dim st As New StackTrace(True)
            st = New StackTrace(ex, True)
            GlobalVariables.Linenumber = st.GetFrame(0).GetFileLineNumber().ToString()
            Dim f As New StackFrame
            ErrorHandler(ex, System.Reflection.MethodBase.GetCurrentMethod().Name, GlobalVariables.Linenumber, f)
        End Try
    End Sub
    Private Sub ToolStripButton5_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton5.Click
        FormPass.ShowDialog()
        ToolStripButton5.Text = GlobalVariables.pass
        Actionlog("Pass rate tool bar pressed - Pass = " & GlobalVariables.pass)
    End Sub
    Public Sub SortTestDb(Token As String)
        'Sort the datatable
        ConnectedDB = New DataBaseFunctions()               'Open the database connection
        GlobalVariables.ds = ConnectedDB.PopulateDataSet()       'Put the datainto a dataset
        Dim hold1 As DataRow = GlobalVariables.ds.Tables("TestResults").NewRow
        Dim comp As New System.Collections.Comparer(System.Globalization.CultureInfo.CurrentCulture)
        Dim current As Integer = 0
        Dim CurrentPlus1 As Integer = 0
        Dim SameAnimal As Integer = 0
        'If the Token is a interger


        For cntOutter As Int32 = 0 To GlobalVariables.ds.Tables("TestResults").Rows.Count - 1
            For cntInner As Int32 = 0 To GlobalVariables.ds.Tables("TestResults").Rows.Count - 2
                If comp.Compare(GlobalVariables.ds.Tables("TestResults").Rows(cntInner).Item(Token), GlobalVariables.ds.Tables("TestResults").Rows(cntInner + 1).Item(Token)) = 1 Then
                    hold1.BeginEdit()
                    hold1.ItemArray = GlobalVariables.ds.Tables("TestResults").Rows(cntInner + 1).ItemArray
                    hold1.EndEdit()
                    GlobalVariables.ds.Tables("TestResults").Rows.RemoveAt(cntInner + 1)
                    GlobalVariables.ds.Tables("TestResults").Rows.InsertAt(hold1, cntInner)
                    hold1 = GlobalVariables.ds.Tables("TestResults").NewRow
                End If
            Next cntInner
        Next cntOutter

    End Sub
    Private Sub ToolStripButton6_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton6.Click
        Actionlog("Settings tool bar pressed")
        FormSettings.ShowDialog()
    End Sub
    Public Sub Email(Subject As String, body As String, Attachment As String)
        GlobalVariables.Body = body
        GlobalVariables.Attachment = Attachment
        GlobalVariables.Subject = Subject
        FormEmail.ShowDialog()
    End Sub
    Private Sub ToolStripButton7_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton7.Click
        Actionlog("About tool bar pressed")
        FormAbout.ShowDialog()
    End Sub
    Private Sub ToolStripButton8_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton8.Click
        Actionlog("Close tool bar pressed")
        Me.Close()
    End Sub
    Private Sub ToolStripButton9_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton9.Click
        Actionlog("Backup tool bar pressed")
        'FormBackup.Show()
        Dim dt As String = ""

        If (Not System.IO.Directory.Exists("C: \WormCountDatabaseBackup")) Then
            System.IO.Directory.CreateDirectory("C:\WormCountDatabaseBackup")
        End If
        If (Not System.IO.Directory.Exists("C:\WormCountDatabaseBackup\WCD")) Then
            System.IO.Directory.CreateDirectory("C:\WormCountDatabaseBackup\WCD")
        End If



        Dim dtNow As DateTime = Now
        dt = "C:\WormCountDatabaseBackup\" & "BackupWormDb" & String.Format("{0:00}", dtNow.Day) & String.Format("{0:00}", dtNow.Month) & String.Format("{0:00}", dtNow.Year) & String.Format("{0:00}", dtNow.Hour) & String.Format("{0:00}", dtNow.Minute) & String.Format("{0:00}", dtNow.Second)
        Try
            My.Computer.FileSystem.CopyFile("C:\WormCountDATA/Header.png", "C:\WormCountDatabaseBackup\WCD\Header.png", True)
            My.Computer.FileSystem.CopyFile("C:\WormCountDATA/Settings.xml", "C:\WormCountDatabaseBackup\WCD\Settings.xml", True)
            My.Computer.FileSystem.CopyFile("C:\WormCountDATA/systemlog.txt", "C:\WormCountDatabaseBackup\WCD\systemlog.txt", True)
            My.Computer.FileSystem.CopyFile("C:\WormCountDATA/WormCount.mdb", "C:\WormCountDatabaseBackup\WCD\WormCount.mdb", True)

            ZipFile.CreateFromDirectory("C:\WormCountDatabaseBackup\WCD\", dt & ".zip")
            Email("Worm count databas - Backup", "Here is you backup, save it in case it is needed", dt & ".zip")
        Catch ex As Exception
            MsgBox("backup failed", vbAbort & vbCritical & vbSystem, "Error")
        End Try
    End Sub
    Private Sub ToolStripButton10_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton10.Click
        Actionlog("Remove animal tool bar pressed")
        FormAnimalDelete.ShowDialog()

    End Sub
    Private Sub Timer1_Tick(sender As System.Object, e As System.EventArgs) Handles Timer1.Tick

        If Loaded Then
            Dim testNeeded As Boolean = False
            Try
                ToolStripStatusLabel2.Text = "Version number " & GlobalVariables.Version & "    | "


            Catch ex As Exception
            End Try

            If Process.GetProcessesByName("AWC").Length > 0 Then
                ToolStripStatusLabel5.Text = "Service AWC Running"
                ToolStripStatusLabel5.BackColor = Color.LightGreen
            Else
                ToolStripStatusLabel5.Text = "Service AWC Stopped"
                ToolStripStatusLabel5.BackColor = Color.Red
            End If

            'Test to see if there is data in the grd view if not refresh
            If GlobalVariables.ds.Tables("GridView").Columns.Count >= 3 Then
                If GlobalVariables.ds.Tables("GridView").Rows.Count >= 2 Then
                    Try
                        Dim test As String = DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(2).Value.ToString
                    Catch
                        DataGridViewRefresh()
                    End Try
                End If
            End If

        End If

    End Sub
    Private Function Isreminder() As Boolean

        If GlobalVariables.ReTestAddedTime > 0 And GlobalVariables.ReminderState Then
            Isreminder = True
        Else
            Isreminder = False
        End If
        Return Isreminder
    End Function
    Public Function CallNote(ByVal title As String, ByVal message As String, ByVal image As Image) As Boolean
        Dim SecondForm As New FormNote
        SecondForm.Show()
        SecondForm.Label4.Text = title
        SecondForm.Label2.Text = message
        SecondForm.PictureBox1.Image = image
        Return True
    End Function
    Private Function GetAnimalStats(ByVal name As String) As String
        Dim tests As String = ""
        Dim NumberOfTest As Integer = 0
        Dim NumberOfFails As Integer = 0
        Dim HighesTestResult As Integer = 0
        Dim TestHighestDate As DateTime


        Try
            ConnectedDB = New DataBaseFunctions()    'Open the database connection
            GlobalVariables.ds = ConnectedDB.PopulateDataSet()       'Put the datainto a dataset
            For Each row As DataRow In GlobalVariables.ds.Tables("TestResults").Rows

                'Find if this is the animals stats required
                If row.Field(Of String)("TestName") = name Then
                    'See if this is the worst test result
                    If (row.Field(Of String)("TestEPGTotal") + row.Field(Of String)("TestOPGTotal")) > HighesTestResult Then
                        HighesTestResult = (row.Field(Of String)("TestEPGTotal") + row.Field(Of String)("TestOPGTotal"))
                        TestHighestDate = row.Field(Of DateTime)("TestDate")
                    End If

                    If (row.Field(Of String)("TestEPGTotal") + row.Field(Of String)("TestOPGTotal")) >= GlobalVariables.pass Then
                        NumberOfFails = NumberOfFails + 1
                        NumberOfTest = NumberOfTest + 1
                    Else
                        NumberOfTest = NumberOfTest + 1
                    End If

                    tests = "Has had " & NumberOfTest & " tests and passed " & (NumberOfTest - NumberOfFails) & " failed " & NumberOfFails & " times. The worst test result was " & HighesTestResult & " on the " & TestHighestDate
                End If
            Next
        Catch ex As Exception
            Dim st As New StackTrace(True)
            st = New StackTrace(ex, True)
            GlobalVariables.Linenumber = st.GetFrame(0).GetFileLineNumber().ToString()
            Dim f As New StackFrame
            ErrorHandler(ex, System.Reflection.MethodBase.GetCurrentMethod().Name, GlobalVariables.Linenumber, f)
        End Try

        Return tests
    End Function
    Public Function IsAnimalOnFarm(ByVal name As String) As Boolean
        Dim ans As Boolean = False
        ConnectedDB = New DataBaseFunctions()    'Open the database connection
        GlobalVariables.ds = ConnectedDB.PopulateDataSet()       'Put the datainto a dataset
        For Each row As DataRow In GlobalVariables.ds.Tables("Alpaca").Rows
            Try

                If row.Field(Of String)("Name") = name Then
                    If row.Field(Of Boolean)("OnFarm") = True Then
                        ans = True
                    Else
                        ans = False
                    End If
                End If
            Catch ex As Exception
                Dim st As New StackTrace(True)
                st = New StackTrace(ex, True)
                GlobalVariables.Linenumber = st.GetFrame(0).GetFileLineNumber().ToString()
                Dim f As New StackFrame
                ErrorHandler(ex, System.Reflection.MethodBase.GetCurrentMethod().Name, GlobalVariables.Linenumber, f)
            End Try
        Next
        Return ans
    End Function
    Public Sub ftpErrorlog()
        Try



            Dim t As New clsComputerInfo
            Systemlog("AvailablePhysicalMemory", "           ", "   ", My.Computer.Info.AvailablePhysicalMemory)
            Systemlog("AvailableVirtualMemory", "         ", "  ", My.Computer.Info.AvailableVirtualMemory)
            Systemlog("DisplayName", "          ", "          ", My.Computer.Info.InstalledUICulture.DisplayName)
            Systemlog("OSFullName", "          ", "          ", My.Computer.Info.OSFullName)
            Systemlog("OSPlatform", "          ", "          ", My.Computer.Info.OSPlatform)
            Systemlog("OSVersion", "          ", "          ", My.Computer.Info.OSVersion)
            Systemlog("TotalPhysicalMemory", "        ", "    ", My.Computer.Info.TotalPhysicalMemory)
            Systemlog("TotalVirtualMemory", "        ", "   ", My.Computer.Info.TotalVirtualMemory)
            Systemlog("GetMACAddress", "          ", "          ", t.GetMACAddress())
            Systemlog("GetMotherBoardID", "          ", "   ", t.GetMotherBoardID())
            Systemlog("GetProcessorId", "          ", "          ", t.GetProcessorId())
            Systemlog("GetVolumeSerial", "          ", "          ", t.GetVolumeSerial())



            FormBackup.BackupFile()
            Dim myFile As String = ""
            Try
                myFile = Directory.GetFiles("C:\WormCountDatabaseBackup\").OrderByDescending(Function(f) New DirectoryInfo(f).LastWriteTime).First()
                myFile = Replace(myFile, "C:\WormCountDatabaseBackup\", "")
            Catch
            End Try


            FtpFolderCreate("ftp://www.mullacottalpacas.com/wwwroot/aig/WDB/" & GlobalVariables.FarmName.Replace(" ", ""), GlobalVariables.User, GlobalVariables.Password)

            'Send error log
            Dim Sendrequest1 As System.Net.FtpWebRequest = DirectCast(System.Net.WebRequest.Create("ftp://www.mullacottalpacas.com/wwwroot/aig/WDB/" & GlobalVariables.FarmName.Replace(" ", "") & "_" & myFile), System.Net.FtpWebRequest)
            Sendrequest1.EnableSsl = False
            Sendrequest1.UsePassive = False
            Sendrequest1.Credentials = New System.Net.NetworkCredential(GlobalVariables.User, GlobalVariables.Password)
            Sendrequest1.Method = System.Net.WebRequestMethods.Ftp.UploadFile
            Sendrequest1.KeepAlive = False

            Dim dtNow As DateTime = Now
            Dim Sendfile1() As Byte = System.IO.File.ReadAllBytes("C:\WormCountDatabaseBackup\" & myFile)
            Dim strz1 As System.IO.Stream = Sendrequest1.GetRequestStream()
            strz1.Write(Sendfile1, 0, Sendfile1.Length)
            'strz1.Close()
            strz1.Dispose()




        Catch ex As Exception
            Dim st As New StackTrace(True)
            st = New StackTrace(ex, True)
            GlobalVariables.Linenumber = st.GetFrame(0).GetFileLineNumber().ToString()
            Dim f As New StackFrame
            ErrorHandler(ex, System.Reflection.MethodBase.GetCurrentMethod().Name, GlobalVariables.Linenumber, f)
        End Try
    End Sub
    Private Function FtpFolderCreate(folder_name As String, username As String, password As String) As Boolean



        Dim request As Net.FtpWebRequest = CType(FtpWebRequest.Create(folder_name), FtpWebRequest)
        request.Credentials = New NetworkCredential(username, password)
        request.Method = WebRequestMethods.Ftp.MakeDirectory

        Try
            Using response As FtpWebResponse = DirectCast(request.GetResponse(), FtpWebResponse)
                ' Folder created
            End Using
        Catch ex As WebException
            Dim response As FtpWebResponse = DirectCast(ex.Response, FtpWebResponse)
            ' an error occurred
            If response.StatusCode = FtpStatusCode.ActionNotTakenFileUnavailable Then
                Return False
            End If
        End Try
        Return True
    End Function
    Private Sub ToolStripButton11_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton11.Click
        Actionlog("Refresh tool bar pressed")
        DataGridViewRefresh()
    End Sub
    Private Sub CheckReminders()
        Dim Bodystr As String = ""

        ConnectedDB = New DataBaseFunctions()       'Open the database connection
        GlobalVariables.ds = ConnectedDB.PopulateDataSet()          'Put the data into a dataset
        Dim str As String = ""
        If GlobalVariables.ds.Tables("Reminder").Rows.Count > 0 Then
            For j = 0 To GlobalVariables.ds.Tables("Reminder").Rows.Count - 1
                If Now >= GlobalVariables.ds.Tables("Reminder").Rows(j)("ReminderDate") Then
                    If GlobalVariables.ds.Tables("Reminder").Rows(j)("ReminderEmail") = True And GlobalVariables.ds.Tables("Reminder").Rows(j)("ReminderDone") = False Then
                        Bodystr = GlobalVariables.ds.Tables("Reminder").Rows(j)("ReminderNote")
                        str = GlobalVariables.ds.Tables("Reminder").Rows(j)("ReminderNote")
                        ConnectedDB.UpdateDatabase("UPDATE Reminder SET ReminderDone = '" & True & "' WHERE ReminderNote = '" & str & "'")
                        Email("Worm count database - Reminder", Bodystr, "")
                        Exit For
                    End If
                End If
            Next
        End If



    End Sub
    Private Sub ServiceCheck()
        Module1.Main()
    End Sub
    Private Sub ToolStripStatusLabel5_Click(sender As Object, e As EventArgs) Handles ToolStripStatusLabel5.Click
        Try
            Dim psi As New ProcessStartInfo()

            psi.Verb = "runas" ' aka run as administrator
            psi.FileName = GlobalVariables.DbDriveLocation & "\Service\UnInstall.bat"

            psi.Verb = "runas" ' aka run as administrator
            psi.FileName = GlobalVariables.DbDriveLocation & "\Service\Install.bat"

            Process.Start(psi)
        Catch ex As Exception
            Dim st As New StackTrace(True)
            st = New StackTrace(ex, True)
            GlobalVariables.Linenumber = st.GetFrame(0).GetFileLineNumber().ToString()
            Dim f As New StackFrame
            ErrorHandler(ex, System.Reflection.MethodBase.GetCurrentMethod().Name, GlobalVariables.Linenumber, f)
        End Try
    End Sub
    Private Sub ToolStripButton13_Click(sender As Object, e As EventArgs) Handles ToolStripButton13.Click
        Actionlog("Reminder tool bar pressed")
        FormReminder.ShowDialog()
    End Sub
    Private Sub BackgroundWorker1_DoWork(sender As Object, e As DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        SendErrorReport()
    End Sub
    Private Sub GetUserAndPassword()

        Dim xmlUrl As String = "http://www.mullacottalpacas.com/aig/private.xml"

        Systemlog("System: " & "Main", "   Module: " & "GetUserAndPassword", "    Line number:  " & "1354", "Global request")
        Try
            Using Reader As XmlReader = XmlReader.Create(xmlUrl)
                Systemlog("System: " & "Main", "   Module: " & "GetUserAndPassword", "    Line number:  " & "1356", "Global request link opened")
                While Reader.Read()
                    ' Check for start elements.
                    If Reader.IsStartElement() Then
                        If Reader.Name = "user" Then
                            ' Text data.
                            If Reader.Read() Then
                                GlobalVariables.User = Reader.Value.Trim()
                                Systemlog("System: " & "Main", "   Module: " & "GetUserAndPassword", "    Line number:  " & "1361", "Global user loaded")
                            End If
                        End If
                    End If

                    If Reader.IsStartElement() Then
                        If Reader.Name = "pass" Then
                            ' Text data.
                            If Reader.Read() Then
                                GlobalVariables.Password = Reader.Value.Trim()
                                Systemlog("System: " & "Main", "   Module: " & "GetUserAndPassword", "    Line number:  " & "1372", "Global pass loaded")

                            End If
                        End If
                    End If


                End While
            End Using

        Catch ex As Exception
            Dim st As New StackTrace(True)
            st = New StackTrace(ex, True)
            GlobalVariables.Linenumber = st.GetFrame(0).GetFileLineNumber().ToString()
            Dim f As New StackFrame
            ErrorHandler(ex, System.Reflection.MethodBase.GetCurrentMethod().Name, GlobalVariables.Linenumber, f)
        End Try
    End Sub
    Private Sub Button1_Click_1(sender As Object, e As EventArgs)
        CreateTable()
    End Sub
    Private Sub ToolStripButton4_Click(sender As Object, e As EventArgs) Handles ShowStoll.Click
        Stoll()
    End Sub
    Public Sub Stoll()

        If GlobalVariables.Stoll = True Then
            GlobalVariables.Stoll = False
            SetXMLData()
            If GlobalVariables.Stoll = True Then
                ShowStoll.Image = My.Resources.Stoll2
            Else
                ShowStoll.Image = My.Resources.Stoll1
            End If
            Exit Sub
        End If
        If GlobalVariables.Stoll = False Then
            GlobalVariables.Stoll = True
            SetXMLData()
        End If
        If GlobalVariables.Stoll = True Then
            ShowStoll.Image = My.Resources.Stoll2
        Else
            ShowStoll.Image = My.Resources.Stoll1
        End If
    End Sub

    Private Sub Button1_Click_2(sender As Object, e As EventArgs)
        For Each row As DataRow In GlobalVariables.ds.Tables("TestResults").Rows

            Dim sql As String = "Update TestResults Set EPGTotal = " & CInt(row.ItemArray(6)) + CInt(row.ItemArray(7)) & " Where TestID = " & row.ItemArray(0) & ""

            ConnectedDB.UpdateDatabase(sql)
        Next
    End Sub
End Class


Public Class GlobalVariables
    Public Shared Password As String
    Public Shared Stoll As Boolean
    Public Shared User As String
    Public Shared EmailSending As Boolean
    Public Shared AlpacaName As String
    Public Shared AlpacaTestDate As String
    Public Shared Linenumber As String
    Public Shared pass As Integer
    Public Shared OnFarm As Integer
    Public Shared Clickgroup As String
    Public Shared BackupNumber As Integer
    Public Shared ReTestTime As Integer
    Public Shared ReTestAddedTime As Integer
    Public Shared Messages As Boolean = False
    Public Shared StartMax As Boolean = False
    Public Shared ReminderState As Boolean = False
    Public Shared SelfClose As Boolean = False
    Public Shared TestGroupName As String
    Public Shared Email As String
    Public Shared EmailUsername As String
    Public Shared FarmName As String
    Public Shared EmailPassword As String
    Public Shared EmailServer As String
    Public Shared DbDriveLocation As String
    Public Shared ChartAnimal As String
    Public Shared StartAtBoot As Boolean = True
    Public Shared Note As String
    Public Shared bar As Double
    Public Shared LatestUpdate As String
    Public Shared Version As String
    Public Shared PacaManpath As String
    Public Shared Body As String
    Public Shared Attachment As String
    Public Shared Subject As String
    Public Shared PacaManInstalled As Boolean = True
    Public Shared ColourEnable As Boolean = True
    Public Shared BiasEnable As Boolean = True
    Public Shared ColourLevel(10) As Integer
    Public Shared EggType(10) As Eggs
    Public Shared ds As DataSet
End Class

Public Structure Eggs
    Public Name As String
    Public Bias As Double
End Structure
