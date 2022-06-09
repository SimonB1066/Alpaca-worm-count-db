Imports System.Globalization

Public Class FormNewTestGroup


    Private Sub ToolStrip1_MouseHover(sender As Object, e As EventArgs) Handles ToolStrip1.MouseHover
        Me.Cursor = Cursors.Hand '
    End Sub
    Private Sub ToolStrip1_MouseLeave(sender As Object, e As EventArgs) Handles ToolStrip1.MouseLeave
        Me.Cursor = Cursors.Arrow
    End Sub


    Private Sub FormNewTestGroup_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        CenterToScreen()
        Me.Icon = My.Resources.Form
        Dim iMonth As Integer = Month(DateTime.Now)

        TextBox1.Text = MonthName(iMonth) + " " + Year(DateTime.Now).ToString
        Dim dv As DataView
        FormMain.ConnectedDB = New DataBaseFunctions()              'Open the database connection
        GlobalVariables.ds = FormMain.ConnectedDB.PopulateDataSet()        'Put the datainto a dataset
        dv = New DataView(GlobalVariables.ds.Tables("TestGroup"), "", "", DataViewRowState.CurrentRows)
        DataGridView1.DataSource = dv
        DataGridView1.Refresh()


    End Sub

    Private Sub ToolStripButton2_Click(sender As Object, e As EventArgs) Handles ToolStripButton2.Click
        Close()
    End Sub

    Private Sub ToolStripButton1_Click(sender As Object, e As EventArgs) Handles ToolStripButton1.Click
        'Update the alpaca database
        FormMain.ConnectedDB = New DataBaseFunctions()              'Open the database connection
        GlobalVariables.ds = FormMain.ConnectedDB.PopulateDataSet()        'Put the datainto a dataset

        Dim drow As DataRow = GlobalVariables.ds.Tables("TestGroup").NewRow

        drow.Item(1) = TextBox1.Text
        drow.Item(2) = CDate(Now.ToShortDateString)

        Dim foundRow() As DataRow
        foundRow = GlobalVariables.ds.Tables("TestGroup").Select("GroupName='" & TextBox1.Text & "'")

        If foundRow.Length < 1 Then
            Dim dtUK As Date = DateTime.ParseExact(Now.ToShortDateString, "dd/MM/yyyy", CultureInfo.InvariantCulture)


            Dim sql As String = "INSERT INTO TestGroup (GroupName,GroupCreated) VALUES ('" & TextBox1.Text & "',#" & dtUK & "#)"
            FormMain.ConnectedDB.UpdateDatabase(sql)

            sql = "INSERT INTO TestResults (TestDate,TestName,TestFaecalGrade,TestNumber,TestNumberName,TestEPGTotal,TestOPGTotal,TestTrichostrongyles,TestTrichurius,TestNematordirus,TestCapillarid,TestMoniezid,TestEPGUnidentifed,TestEmac,TestEivitaesis,TestEalpacae,TestElamae,TestEpunoensis,TestOPGUnidentifed,TestTreat) VALUES (#" & dtUK & "#,'xxxSystemxxx','2','0','" & TextBox1.Text & "','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0')"
            FormMain.ConnectedDB.UpdateDatabase(sql)

            Dim ID As String = GlobalVariables.ds.Tables("TestGroup").Rows.IndexOf(drow)
            GlobalVariables.TestGroupName = TextBox1.Text
            Me.Close()



            Dim col As New DataGridViewTextBoxColumn
            col.HeaderText = GlobalVariables.TestGroupName
            col.Name = GlobalVariables.TestGroupName
            col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            FormMain.DataGridView1.Columns.Add(col)

            FormMain.DataGridViewRefresh()

        Else
            MessageBox.Show("The entered group name is already in use.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If
    End Sub
End Class