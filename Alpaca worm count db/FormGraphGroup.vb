Imports System.Data.OleDb
Imports System.Data.SqlClient
Imports System.Windows.Forms.DataVisualization.Charting



Public Class FormGraphGroup

    Dim sqlstr As String
    Dim Checkedstr As String


    Private Sub FormGraphGroup_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.Icon = My.Resources.Form
        'Set all the chart settings
        Chart1.ChartAreas.Add("ChartArea1")
        Chart1.ChartAreas(0).AxisX.Interval = 1
        Chart1.Titles.Add("Egg count for test group " & GlobalVariables.TestGroupName)
        Chart1.ChartAreas("ChartArea1").AxisX.Title = "Animal name"
        Chart1.ChartAreas("ChartArea1").AxisY.Title = "EPG and OPG"


        ' CheckedListBox1.Items.Add("All")
        FormMain.ConnectedDB = New DataBaseFunctions()              'Open the database connection
        GlobalVariables.ds = FormMain.ConnectedDB.PopulateDataSet()        'Put the datainto a dataset
        For Each GroupRow As DataRow In GlobalVariables.ds.Tables("TestGroup").Rows
            CheckedListBox1.Items.Add(GroupRow.Item(1))
            Chart1.Series.Add(GroupRow.Item(1))
            Chart1.Legends.Add(GroupRow.Item(1))
            Chart1.Series(GroupRow.Item(1)).Enabled = False
        Next

        'Now set the last entry in the checkboxlist on (Ticked)
        Dim last As Integer = CheckedListBox1.Items.Count
        CheckedListBox1.SetItemChecked(last - 1, True)
        CheckedListBox1.SelectedItem = GlobalVariables.ds.Tables("TestGroup")(GlobalVariables.ds.Tables("TestGroup").Rows.Count - 1)("GroupName")
        Chart1.Series(CheckedListBox1.SelectedItem).Enabled = True
        ChartByTestGroup()


    End Sub

    Private Sub ChartByTestGroup()
        Dim dbProvider As String
        Dim dbSource As String



        Dim Conn As OleDbConnection = New OleDbConnection
        dbProvider = "PROVIDER=Microsoft.Jet.OLEDB.4.0;"
        dbSource = "Data Source = C:/WormCountDatabase/WormCount.mdb"
        Conn.ConnectionString = dbProvider & dbSource
        Conn.Open()

        Dim cmd As OleDbCommand = New OleDbCommand("SELECT [TestName],[TestEPGTotal],[TestOPGTotal],[TestNumberName] FROM [TestResults] Order By ABS(TestEPGTotal) ASC", Conn)
        Dim dr As OleDbDataReader = cmd.ExecuteReader


        While dr.Read

            Chart1.Series(dr("TestNumberName")).Points.AddXY(dr("TestName").ToString, (System.Convert.ToInt32(dr("TestEPGTotal")) + System.Convert.ToInt32(dr("TestOPGTotal")).ToString).ToString)
            Chart1.Series(dr("TestNumberName")).Sort(PointSortOrder.Ascending, "Y")

            Dim grpDT As New DataTable
            FormMain.ConnectedDB = New DataBaseFunctions()              'Open the database connection
            GlobalVariables.ds = FormMain.ConnectedDB.PopulateDataSet()        'Put the datainto a dataset
            grpDT = GlobalVariables.ds.Tables("TestResults")
            For Each dp As DataPoint In Chart1.Series(dr("TestNumberName")).Points
                If dp.YValues(0) > 25 Then
                    For i As Integer = 0 To grpDT.Rows.Count - 1
                        If dp.YValues(0) <= GlobalVariables.pass Then
                            dp.Color = Me.Chart1.Series(dr("TestNumberName")).Color
                        Else
                            dp.Color = Color.Red
                        End If
                    Next
                End If
            Next
        End While


       

        dr.Close()
        cmd.Dispose()
    End Sub

    Private Sub CheckedListBox1_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles CheckedListBox1.SelectedIndexChanged

        If CheckedListBox1.GetItemCheckState(CheckedListBox1.SelectedIndex) Then
            Chart1.Series(CheckedListBox1.SelectedItem).Enabled = True
        Else
            Chart1.Series(CheckedListBox1.SelectedItem).Enabled = False
        End If

        Chart1.ChartAreas(0).RecalculateAxesScale()


    End Sub




    Private Sub chart1_GetToolTipText(sender As Object, e As ToolTipEventArgs) Handles Chart1.GetToolTipText
        ' Check selected chart element and set tooltip text for it
        Select Case e.HitTestResult.ChartElementType
            Case ChartElementType.DataPoint
                Dim dataPoint = e.HitTestResult.Series.Points(e.HitTestResult.PointIndex)
                e.Text = dataPoint.AxisLabel & " test result was --- " & dataPoint.YValues(0).ToString & " EPG"
                GlobalVariables.ChartAnimal = dataPoint.AxisLabel
                Exit Select
        End Select
    End Sub

    Private Sub Chart1_DoubleClick(sender As System.Object, e As System.EventArgs) Handles Chart1.DoubleClick

        'FormGraphAnimal.ShowDialog()

    End Sub
End Class