Imports System.Data.OleDb
Imports System.Data.SqlClient
Imports System.Windows.Forms.DataVisualization.Charting

Public Class FormGraphAnimal

   


    Private Sub FormGraphAnimal_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        'Set all the chart settings
        Me.Icon = My.Resources.Form
        Chart1.ChartAreas.Add("ChartArea2")
        Chart1.ChartAreas(0).AxisX.Interval = 1
        Chart1.Titles.Add("Egg count for  " & GlobalVariables.ChartAnimal)
        Chart1.ChartAreas("ChartArea2").AxisX.Title = "Test date"
        Chart1.ChartAreas("ChartArea2").AxisY.Title = "EPG and OPG"
        Chart1.Series.Add("Animal")
        Chart1.Series("Animal").ChartType = DataVisualization.Charting.SeriesChartType.Line

        ChartByAnimal()
    End Sub
    Private Sub FormGraphAnimal_FormClosed(sender As Object, e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed

    End Sub

    Private Sub ChartByAnimal()
        Dim dbProvider As String
        Dim dbSource As String
        Dim Conn As OleDbConnection = New OleDbConnection
        dbProvider = "PROVIDER=Microsoft.Jet.OLEDB.4.0;"
        dbSource = "Data Source = C:/WormCountDatabase/WormCount.mdb"
        Conn.ConnectionString = dbProvider & dbSource
        Conn.Open()
        Dim sql As String = "SELECT [TestName],[TestEPGTotal],[TestOPGTotal],[TestDate] FROM [TestResults] WHERE [TestName] = '" & GlobalVariables.ChartAnimal & "' Order By ABS(TestDate) ASC"
        Dim cmd As OleDbCommand = New OleDbCommand(sql, Conn)
        Dim dr As OleDbDataReader = cmd.ExecuteReader


        While dr.Read

            Chart1.Series("Animal").Points.AddXY(dr("TestDate").ToString, (System.Convert.ToInt32(dr("TestEPGTotal")) + System.Convert.ToInt32(dr("TestOPGTotal")).ToString).ToString)
            Chart1.Series("Animal").Sort(PointSortOrder.Ascending, "Y")

        End While




        dr.Close()
        cmd.Dispose()
    End Sub




End Class