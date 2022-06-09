

Imports System.Xml
Imports System.Reflection
Imports System.Net

Public Class FormAbout




    Private Sub Label8_MouseHover(sender As Object, e As EventArgs)
        Me.Cursor = Cursors.Hand '
    End Sub
    Private Sub Label8_MouseLeave(sender As Object, e As EventArgs)
        Me.Cursor = Cursors.Arrow
    End Sub

    Private Sub Label7_MouseHover(sender As Object, e As EventArgs) Handles Label7.MouseHover
        Me.Cursor = Cursors.Hand '
    End Sub
    Private Sub Label7_MouseLeave(sender As Object, e As EventArgs) Handles Label7.MouseLeave
        Me.Cursor = Cursors.Arrow
    End Sub



    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        Me.Close()
    End Sub

    Private Sub FormAbout_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        CenterToParent()
        Label9.Text = ""
    End Sub



    Private Sub Label7_Click(sender As System.Object, e As System.EventArgs) Handles Label7.Click
        Dim webAddress As String = Label7.Text
        Process.Start(webAddress)
    End Sub



    Private Sub FormAbout_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        Me.Icon = My.Resources.Form
        Label1.Text = Process.GetCurrentProcess.ProcessName.ToString
        Try
            Label2.Text = "Version number " & System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.ToString
        Catch
            Label2.Text = ""
        End Try
        Label3.Text = "Simon Brookes"
        Label4.Text = "Lower Mullacott Farm"
        Label5.Text = "Ilfracombe"
        Label6.Text = "EX34 8NA"



        Dim t As New clsComputerInfo
        TextBox1.Text = ""
        TextBox1.Text = TextBox1.Text & "AvailablePhysicalMemory ---" & My.Computer.Info.AvailablePhysicalMemory & vbNewLine
        TextBox1.Text = TextBox1.Text & "AvailableVirtualMemory ----" & My.Computer.Info.AvailableVirtualMemory & vbNewLine
        TextBox1.Text = TextBox1.Text & "DisplayName ---------------" & My.Computer.Info.InstalledUICulture.DisplayName & vbNewLine
        TextBox1.Text = TextBox1.Text & "OSFullName ----------------" & My.Computer.Info.OSFullName & vbNewLine
        TextBox1.Text = TextBox1.Text & "OSPlatform ----------------" & My.Computer.Info.OSPlatform & vbNewLine
        TextBox1.Text = TextBox1.Text & "OSVersion -----------------" & My.Computer.Info.OSVersion & vbNewLine
        TextBox1.Text = TextBox1.Text & "TotalPhysicalMemory -------" & My.Computer.Info.TotalPhysicalMemory & vbNewLine
        TextBox1.Text = TextBox1.Text & "TotalVirtualMemory --------" & My.Computer.Info.TotalVirtualMemory & vbNewLine
        TextBox1.Text = TextBox1.Text & "GetMACAddress -------------" & t.GetMACAddress() & vbNewLine
        TextBox1.Text = TextBox1.Text & "GetMotherBoardID ----------" & t.GetMotherBoardID() & vbNewLine
        TextBox1.Text = TextBox1.Text & "GetProcessorId ------------" & t.GetProcessorId() & vbNewLine
        TextBox1.Text = TextBox1.Text & "GetVolumeSerial -----------" & t.GetVolumeSerial() & vbNewLine
        TextBox1.Text = TextBox1.Text & vbNewLine & "_____________________________________________________"
        TextBox1.Text = TextBox1.Text & vbNewLine & "This program will send error logs and crash reports"
        TextBox1.Text = TextBox1.Text & vbNewLine & "to aid in improving the product. No personal information"
        TextBox1.Text = TextBox1.Text & vbNewLine & "is recorded."
        TextBox1.Text = TextBox1.Text & vbNewLine & "_____________________________________________________"
        TextBox1.Text = TextBox1.Text & vbNewLine & "References used in this program: "
        TextBox1.Text = TextBox1.Text & vbNewLine & "Sue Thomas of Lyme Alpacas"
        TextBox1.Text = TextBox1.Text & vbNewLine & "Veterinary Parasitology Reference manual"
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs)
        Label9.Text = "Checking for updates at - Mullacottalpacas.com"
        Refresh()
        GetUpdateUrlAndVersion()
    End Sub

    Public Sub GetUpdateUrlAndVersion()
        Try
            Dim downloadUrl As String = ""
            Dim newVersion As Version = Nothing
            Dim applicationVersion As Version = Nothing
            Dim xmlUrl As String = "http://www.mullacottalpacas.com/aig/update.xml"

            applicationVersion = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version
            Label9.Text = "Checking for xml"
            Refresh()

            Using Reader As XmlReader = XmlReader.Create(xmlUrl)
                While Reader.Read()
                    ' Check for start elements.
                    If Reader.IsStartElement() Then
                        If Reader.Name = "version" Then
                            ' Text data.
                            If Reader.Read() Then
                                newVersion = New Version(Reader.Value.Trim())
                                Label9.Text = "New version " & newVersion.ToString
                                Refresh()
                            End If
                        End If
                    End If

                    If Reader.IsStartElement() Then
                        If Reader.Name = "url" Then
                            ' Text data.
                            If Reader.Read() Then
                                downloadUrl = Reader.Value.Trim()
                            End If
                        End If
                    End If

                End While

            End Using

            If applicationVersion < newVersion Then
                Label9.Text = "Update found"
                Refresh()
                Dim web_client As WebClient = New WebClient
                Dim rtn As Integer = 0

                ' Download the file.
                Label9.Text = "Renaming remote file"
                Refresh()

                If System.IO.File.Exists("C:\temp\WormCountDatabase.txt") = True Then
                    System.IO.File.Delete("C:\temp\WormCountDatabase.txt")
                End If
                Label9.Text = "Creating zip"
                Refresh()

                If System.IO.File.Exists("C:\temp\WormCountDatabase.zip") = True Then
                    System.IO.File.Delete("C:\temp\WormCountDatabase.zip")
                End If

                Label9.Text = "Downloading zip"
                Refresh()
                web_client.DownloadFile(downloadUrl, "C:\temp\WormCountDatabase.txt")
                My.Computer.FileSystem.RenameFile("C:\temp\WormCountDatabase.txt", "WormCountDatabase.zip")
                Label9.Text = "Cleaning up after download"
                If System.IO.File.Exists("C:\temp\WormCountDatabase.txt") = True Then
                    System.IO.File.Delete("C:\temp\WormCountDatabase.txt")
                End If

                Label9.Text = "Unzip file"
                Refresh()
                'Now unzip it
                Dim startPath As String = "C:\temp"
                Dim zipPath As String = "C:\temp\WormCountDatabase.zip"
                Dim extractPath As String = "c:\"

                Label9.Text = "Clean extract location"
                Refresh()


                Label9.Text = "Extract all"
                Refresh()
                If System.IO.Directory.Exists("C:\WormCountDatabase") = True Then
                    System.IO.Directory.Delete("C:\WormCountDatabase", True)
                End If
                'ZipFile.ExtractToDirectory(zipPath, extractPath)

                Dim res As List(Of String) = System.IO.Directory.GetFiles("C:\WormCountDatabase", "*", System.IO.SearchOption.AllDirectories).ToList
                For i As Integer = 0 To res.Count - 1
                    Label9.Text = res.Item(i)
                    Refresh()
                    Threading.Thread.Sleep(200)
                Next i

                Label9.Text = "Install updates"
                Refresh()
                Threading.Thread.Sleep(1000)

                Label9.Text = "Application restart required Please wait"
                Refresh()
                Threading.Thread.Sleep(2000)
                System.Diagnostics.Process.Start("C:\WormCountDatabase\setup.exe")
                Application.Exit()
            Else
                Label9.Text = "You are up to date, no new version ready to download."
            End If

        Catch ex As Exception
            Dim st As New StackTrace(True)
            st = New StackTrace(ex, True)
            GlobalVariables.Linenumber = st.GetFrame(0).GetFileLineNumber().ToString()
            Dim f As New StackFrame
            FormMain.ErrorHandler(ex, System.Reflection.MethodBase.GetCurrentMethod().Name, GlobalVariables.Linenumber, f)
        End Try

    End Sub


End Class