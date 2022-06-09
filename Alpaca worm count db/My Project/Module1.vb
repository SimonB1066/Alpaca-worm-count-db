Imports System.ServiceProcess



Module Module1



    Public ServiceName As String = "AWC"
    Public ServiceLocation As String = GlobalVariables.DbDriveLocation & "Service"
    Public installutil As String = "C:\Windows\Microsoft.NET\Framework\v4.0.30319\installutil.exe"
    Public cmd As String = "C:\Windows\system32\cmd.exe"

    Sub Main()
        ' StopService()
        ' DeleteService()
        'UpdateService()
        InstallService()
        ' StartService()
    End Sub

    Sub StopService()
        Dim Service As ServiceController = New ServiceController
        Service.ServiceName = ServiceName
        Try
            If Service.Status = ServiceControllerStatus.Running AndAlso Service.CanStop Then
                Service.Stop()
                Service.WaitForStatus(ServiceControllerStatus.Stopped)
            End If
        Catch ex As Exception
            Dim st As New StackTrace(True)
            st = New StackTrace(ex, True)
            GlobalVariables.Linenumber = st.GetFrame(0).GetFileLineNumber().ToString()
            Dim f As New StackFrame
            FormMain.ErrorHandler(ex, System.Reflection.MethodBase.GetCurrentMethod().Name, GlobalVariables.Linenumber, f)
        End Try

    End Sub

    Sub StartService()
        Try
            Dim srvController = New ServiceController("AWC")
        Catch
            Exit Sub
        End Try

        Threading.Thread.Sleep(10000)
        Dim Service As ServiceController = New ServiceController
        Service.ServiceName = ServiceName
        If ((Service.Status.Equals(ServiceControllerStatus.Stopped)) Or
           (Service.Status.Equals(ServiceControllerStatus.StopPending))) Then
            Service.Start()
            Service.WaitForStatus(ServiceControllerStatus.Running)
        End If
    End Sub

    Sub UpdateService()
        Try
            'My.Computer.FileSystem.CopyDirectory(UpdateLocation, ServiceLocation, True)
        Catch ex As Exception
            Dim st As New StackTrace(True)
            st = New StackTrace(ex, True)
            GlobalVariables.Linenumber = st.GetFrame(0).GetFileLineNumber().ToString()
            Dim f As New StackFrame
            FormMain.ErrorHandler(ex, System.Reflection.MethodBase.GetCurrentMethod().Name, GlobalVariables.Linenumber, f)
        End Try

    End Sub

    Sub DeleteService()
        Try
            Dim processinstance As New Process
            With processinstance
                .StartInfo.FileName = "installutil"
                .StartInfo.Arguments = "-u " & ServiceName
                .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
            End With
            processinstance.Start()
            Do Until processinstance.HasExited
                Threading.Thread.Sleep(5)
            Loop
            'Process.Start(cmd, installutil + " -u " + ServiceName)
        Catch ex As Exception
            Dim st As New StackTrace(True)
            st = New StackTrace(ex, True)
            GlobalVariables.Linenumber = st.GetFrame(0).GetFileLineNumber().ToString()
            Dim f As New StackFrame
            FormMain.ErrorHandler(ex, System.Reflection.MethodBase.GetCurrentMethod().Name, GlobalVariables.Linenumber, f)
        End Try
    End Sub


    Sub InstallService()
        Dim srvController As ServiceController = Nothing
        Try
            srvController = New ServiceController("AWC")
            If srvController.DisplayName = "AWC" Then
                FormMain.Actionlog("Service AWC running")
                Exit Sub
            End If
        Catch

        End Try


        Try
            FormMain.Actionlog("Service AWC not running, try to load and start")
            Dim psi As New ProcessStartInfo()
            Dim str1 As String = installutil & " "
            Dim str2 As String = ServiceLocation & "\" & ServiceName & ".exe"
            psi.Verb = "runas" ' aka run as administrator
            psi.FileName = GlobalVariables.DbDriveLocation & "Service\Install.bat"

            Dim sb As New System.Text.StringBuilder
            sb.AppendLine("@echo off")
            sb.AppendLine("""" & GlobalVariables.DbDriveLocation & "Service\installutil.exe" & """" & " """ & GlobalVariables.DbDriveLocation & "Service\" & "AWC.exe""")
            sb.AppendLine("sc description AWC ""Mullacott Alpacas worm count database reminder app""")
            sb.AppendLine("sc start AWC")
            IO.File.WriteAllText(GlobalVariables.DbDriveLocation & "Service\Install.bat", sb.ToString())


            Process.Start(psi)


            'Process.Start(cmd, installutil + " " + ServiceLocation + "\" & ServiceName & ".exe")
        Catch ex As Exception
            Dim st As New StackTrace(True)
            st = New StackTrace(ex, True)
            GlobalVariables.Linenumber = st.GetFrame(0).GetFileLineNumber().ToString()
            Dim f As New StackFrame
            FormMain.ErrorHandler(ex, System.Reflection.MethodBase.GetCurrentMethod().Name, GlobalVariables.Linenumber, f)
        End Try
    End Sub

End Module
