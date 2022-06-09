Imports System.ComponentModel
Imports System.IO

Public Class UpdateMsgBox
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Try
            Dim webAddress As String = "http://www.mullacottalpacas.com/aig/wdb.html"
            Process.Start(webAddress)
            Threading.Thread.Sleep(500)
            Refresh()

        Catch ex As Exception
            Dim st As New StackTrace(True)
            st = New StackTrace(ex, True)
            GlobalVariables.Linenumber = st.GetFrame(0).GetFileLineNumber().ToString()
            Dim f As New StackFrame
            FormMain.ErrorHandler(ex, System.Reflection.MethodBase.GetCurrentMethod().Name, GlobalVariables.Linenumber, f)
        End Try
        Close()

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Close()
    End Sub

    Private Sub UpdateMsgBox_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CenterToScreen()


        GlobalVariables.bar = 0

        Label1.Text = "New version found " & Replace(Replace(GlobalVariables.LatestUpdate, ".txt", ""), "_", ".")
    End Sub



End Class