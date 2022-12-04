

Public Class DataBaseFunctions
    Private DBFile As String
    Public OpenCon As OleDb.OleDbConnection
    Dim da As OleDb.OleDbDataAdapter
    Dim ds As New DataSet


    Public Sub New()
        DBFile = "WormCount.mdb"

    End Sub

    Public Function PopulateDataSet() As DataSet
        Try
            Dim dbProvider As String = "PROVIDER=Microsoft.Jet.OLEDB.4.0;"
            Dim dbSource As String = "Data Source = " & GlobalVariables.DbDriveLocation & DBFile
            Using OpenCon = New OleDb.OleDbConnection(dbProvider & dbSource)

                OpenCon.Open()

                Using da = New OleDb.OleDbDataAdapter("SELECT * FROM " & "Alpaca", OpenCon)
                    Dim builder As New OleDb.OleDbCommandBuilder
                    builder = New OleDb.OleDbCommandBuilder(da)
                    da.Fill(ds, "Alpaca")
                End Using

                Using da = New OleDb.OleDbDataAdapter("SELECT * FROM " & "TestResults Query", OpenCon)
                    Dim builder As New OleDb.OleDbCommandBuilder
                    builder = New OleDb.OleDbCommandBuilder(da)
                    da.Fill(ds, "TestResults Query")
                End Using

                Using da = New OleDb.OleDbDataAdapter("SELECT * FROM " & "TestGroup ORDER BY ID DESC", OpenCon)
                    Dim tbuilder As New OleDb.OleDbCommandBuilder
                    tbuilder = New OleDb.OleDbCommandBuilder(da)
                    da.Fill(ds, "TestGroup")
                End Using

                Using da = New OleDb.OleDbDataAdapter("SELECT * FROM " & "TestResults ORDER BY TestDate DESC", OpenCon)
                    Dim rbuilder As New OleDb.OleDbCommandBuilder
                    rbuilder = New OleDb.OleDbCommandBuilder(da)
                    da.Fill(ds, "TestResults")
                End Using


                Using da = New OleDb.OleDbDataAdapter("SELECT * FROM " & "GridView", OpenCon)
                    Dim gbuilder As New OleDb.OleDbCommandBuilder
                    gbuilder = New OleDb.OleDbCommandBuilder(da)
                    da.Fill(ds, "GridView")
                End Using

                Using da = New OleDb.OleDbDataAdapter("SELECT * FROM " & "Reminder", OpenCon)
                    Dim ebuilder As New OleDb.OleDbCommandBuilder
                    ebuilder = New OleDb.OleDbCommandBuilder(da)
                    da.Fill(ds, "Reminder")
                End Using

                Using da = New OleDb.OleDbDataAdapter("SELECT * FROM " & "TestGroupQuery", OpenCon)
                    Dim ebuilder As New OleDb.OleDbCommandBuilder
                    ebuilder = New OleDb.OleDbCommandBuilder(da)
                    da.Fill(ds, "TestGroupQuery")
                End Using

            End Using


        Catch ex As Exception
            Dim st As New StackTrace(True)
            st = New StackTrace(ex, True)
            GlobalVariables.Linenumber = st.GetFrame(0).GetFileLineNumber().ToString()
            Dim f As New StackFrame
            FormMain.ErrorHandler(ex, System.Reflection.MethodBase.GetCurrentMethod().Name, GlobalVariables.LineNumber, f)
        End Try

        Return ds
    End Function

    Public Function UpdateDatabase(sql As String)
        Dim connection As OleDb.OleDbConnection
        Dim oledbAdapter As New OleDb.OleDbDataAdapter
        Dim dbProvider As String = "PROVIDER=Microsoft.Jet.OLEDB.4.0;"
        Dim dbSource As String = "Data Source = " & GlobalVariables.DbDriveLocation & "WormCount.mdb"
        Dim rnt As Integer
        Dim StartTime As Date


        Try
            connection = New OleDb.OleDbConnection((dbProvider & dbSource))
            connection.Open()
            oledbAdapter.UpdateCommand = connection.CreateCommand
            sql = Replace(sql, "True", "1")
            sql = Replace(sql, "False", "0")
            oledbAdapter.UpdateCommand.CommandText = sql
            StartTime = Now
            rnt = oledbAdapter.UpdateCommand.ExecuteNonQuery()
            oledbAdapter.Dispose()
            connection.Dispose()


            UpdateDatabase = False
            Threading.Thread.Sleep(0)
            If rnt > -1 Then
                UpdateDatabase = True
            End If


        Catch ex As Exception
            Dim st As New StackTrace(True)
            st = New StackTrace(ex, True)
            GlobalVariables.Linenumber = st.GetFrame(0).GetFileLineNumber().ToString()
            Dim f As New StackFrame
            FormMain.ErrorHandler(ex, System.Reflection.MethodBase.GetCurrentMethod().Name, GlobalVariables.Linenumber, f)
            UpdateDatabase = False
        End Try


    End Function

    Public Function Close()

        da = Nothing
        ds = Nothing
        OpenCon = Nothing
        Close = True
    End Function


    Public Function Addrow(sql As String)
        Dim connection As OleDb.OleDbConnection
        Dim oledbAdapter As New OleDb.OleDbDataAdapter
        Dim dbProvider As String = "PROVIDER=Microsoft.Jet.OLEDB.4.0;"
        Dim dbSource As String = "Data Source = " & GlobalVariables.DbDriveLocation & "WormCount.mdb"
        Dim rnt As Integer
        Dim StartTime As Date


        Try
            connection = New OleDb.OleDbConnection((dbProvider & dbSource))
            connection.Open()
            oledbAdapter.UpdateCommand = connection.CreateCommand
            sql = Replace(sql, "True", "1")
            sql = Replace(sql, "False", "0")
            oledbAdapter.UpdateCommand.CommandText = sql
            StartTime = Now
            rnt = oledbAdapter.UpdateCommand.ExecuteNonQuery()
            oledbAdapter.Dispose()
            connection.Dispose()


            Addrow = False
            Threading.Thread.Sleep(0)
            If rnt > -1 Then
                Addrow = True
            End If


        Catch ex As Exception
            Dim st As New StackTrace(True)
            st = New StackTrace(ex, True)
            GlobalVariables.Linenumber = st.GetFrame(0).GetFileLineNumber().ToString()
            Dim f As New StackFrame
            FormMain.ErrorHandler(ex, System.Reflection.MethodBase.GetCurrentMethod().Name, GlobalVariables.Linenumber, f)
            Addrow = False
        End Try


    End Function


End Class
