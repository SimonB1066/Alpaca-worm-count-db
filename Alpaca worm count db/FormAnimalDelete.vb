Public Class FormAnimalDelete

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs)
        Me.Close()

    End Sub

    Private Sub ToolStrip1_MouseHover(sender As Object, e As EventArgs) Handles ToolStrip1.MouseHover
        Me.Cursor = Cursors.Hand '
    End Sub
    Private Sub ToolStrip1_MouseLeave(sender As Object, e As EventArgs) Handles ToolStrip1.MouseLeave
        Me.Cursor = Cursors.Arrow
    End Sub

    Private Sub FormAnimalDelete_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        CenterToScreen()
        Me.Icon = My.Resources.Form
        ComboBox1.Items.Clear()
        FormMain.ConnectedDB = New DataBaseFunctions()              'Open the database connection
        GlobalVariables.ds = FormMain.ConnectedDB.PopulateDataSet()        'Put the datainto a dataset
        For Each row As DataRow In GlobalVariables.ds.Tables("Alpaca").Rows
            Try
                ComboBox1.Items.Add(row.Field(Of String)("Name"))
            Catch ex As Exception
                Dim st As New StackTrace(True)
                st = New StackTrace(ex, True)
                GlobalVariables.Linenumber = st.GetFrame(0).GetFileLineNumber().ToString()
                Dim f As New StackFrame
                FormMain.ErrorHandler(ex, System.Reflection.MethodBase.GetCurrentMethod().Name, GlobalVariables.Linenumber, f)
            End Try
        Next

        ComboBox1.Sorted = True
        ComboBox1.SelectedIndex = -1
        ComboBox1.ResetText()
    End Sub

    Private Sub ToolStripButton1_Click(sender As Object, e As EventArgs) Handles ToolStripButton1.Click
        Dim AnimalToDelete As String = ""
        AnimalToDelete = ComboBox1.SelectedItem
        Try
            FormMain.ConnectedDB.UpdateDatabase("DELETE FROM Alpaca WHERE Name='" & ComboBox1.SelectedItem & "'")

        Catch ex As Exception
            Dim st As New StackTrace(True)
            st = New StackTrace(ex, True)
            GlobalVariables.Linenumber = st.GetFrame(0).GetFileLineNumber().ToString()
            Dim f As New StackFrame
            FormMain.ErrorHandler(ex, System.Reflection.MethodBase.GetCurrentMethod().Name, GlobalVariables.Linenumber, f)
        End Try
        FormMain.DataGridViewRefresh()
        Me.Close()
    End Sub

    Private Sub ToolStripButton2_Click(sender As Object, e As EventArgs) Handles ToolStripButton2.Click
        Close()
    End Sub
End Class