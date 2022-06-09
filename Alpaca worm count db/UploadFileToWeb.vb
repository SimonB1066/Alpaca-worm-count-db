Imports System.Xml
Imports System.Net
Imports System.IO

Public Class UploadFileToWeb
    Dim User As String = ""
    Dim Pass As String = ""

    Public Sub Upload(url As String, customer As String)
        GetUserAndPassword()



        Try
            BackupFile()

            Dim myFile As String = ""
            Try
                myFile = Directory.GetFiles("C:\WormCountDatabaseBackup\").OrderByDescending(Function(f) New DirectoryInfo(f).LastWriteTime).First()
                myFile = Replace(myFile, "C:\WormCountDatabaseBackup\", "")
            Catch
            End Try



            'Send error log
            Dim Sendrequest1 As System.Net.FtpWebRequest = DirectCast(System.Net.WebRequest.Create("ftp://www.mullacottalpacas.com/wwwroot/aig/WDB/" & GlobalVariables.FarmName.Replace(" ", "") & "_Backup.zip"), System.Net.FtpWebRequest)
            Sendrequest1.EnableSsl = False
            Sendrequest1.UsePassive = False
            Sendrequest1.Credentials = New System.Net.NetworkCredential(User, Pass)
            Sendrequest1.Method = System.Net.WebRequestMethods.Ftp.UploadFile
            Sendrequest1.KeepAlive = False

            Dim dtNow As DateTime = Now
            Dim Sendfile1() As Byte = System.IO.File.ReadAllBytes("C:\WormCountDatabaseBackup\" & myFile)
            Dim strz1 As System.IO.Stream = Sendrequest1.GetRequestStream()
            strz1.Write(Sendfile1, 0, Sendfile1.Length)
            'strz1.Close()
            strz1.Dispose()

            System.IO.File.Delete(Directory.GetFiles("C:\WormCountDatabaseBackup\").OrderByDescending(Function(f) New DirectoryInfo(f).LastWriteTime).First())
        Catch ex As Exception
        End Try

    End Sub

    Private Sub BackupFile()
        Dim dt As String = ""

        If (Not System.IO.Directory.Exists("C:\WormCountDatabaseBackup")) Then
            System.IO.Directory.CreateDirectory("C:\WormCountDatabaseBackup")
        End If


        Dim dtNow As DateTime = Now
        dt = "C:\WormCountDatabaseBackup\" & "BackupWormDb" & String.Format("{0:00}", dtNow.Day) & String.Format("{0:00}", dtNow.Month) & String.Format("{0:00}", dtNow.Year) & String.Format("{0:00}", dtNow.Hour) & String.Format("{0:00}", dtNow.Minute) & String.Format("{0:00}", dtNow.Second)
        Try
            ZipFile.CreateFromDirectory("C:\WormCountDATA", dt & ".zip")
        Catch ex As Exception
        End Try
    End Sub


    Private Sub GetUserAndPassword()

        Dim xmlUrl As String = "http://www.mullacottalpacas.com/aig/private.xml"

        Using Reader As XmlReader = XmlReader.Create(xmlUrl)
            While Reader.Read()
                ' Check for start elements.
                If Reader.IsStartElement() Then
                    If Reader.Name = "user" Then
                        ' Text data.
                        If Reader.Read() Then
                            User = Reader.Value.Trim()

                        End If
                    End If
                End If

                If Reader.IsStartElement() Then
                    If Reader.Name = "pass" Then
                        ' Text data.
                        If Reader.Read() Then
                            Pass = Reader.Value.Trim()

                        End If
                    End If
                End If


            End While
        End Using
    End Sub



End Class
