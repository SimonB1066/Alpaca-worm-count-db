Public Class FormNewAnimal




    Private Sub FormNewTestGroup_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        CenterToScreen()
        Me.Icon = My.Resources.Form
        TextBox1.Text = ""

    End Sub

    Private Sub ToolStrip1_MouseHover(sender As Object, e As EventArgs) Handles ToolStrip1.MouseHover
        Me.Cursor = Cursors.Hand '
    End Sub
    Private Sub ToolStrip1_MouseLeave(sender As Object, e As EventArgs) Handles ToolStrip1.MouseLeave
        Me.Cursor = Cursors.Arrow
    End Sub

    Private Sub ToolStripButton1_Click(sender As Object, e As EventArgs) Handles ToolStripButton1.Click
        Try
            If TextBox1.Text <> "" Then
                For Each row As DataRow In GlobalVariables.ds.Tables("Alpaca").Rows
                    Try
                        If TextBox1.Text.ToUpper = row.Field(Of String)("Name").ToUpper Then
                            MessageBox.Show("You already have an animal called " & TextBox1.Text & " in the database." & vbNewLine & "Enter a unique name.", "Warning")
                            Exit Sub
                        End If
                    Catch ex As Exception
                        Dim st As New StackTrace(True)
                        st = New StackTrace(ex, True)
                        GlobalVariables.Linenumber = st.GetFrame(0).GetFileLineNumber().ToString()
                        Dim f As New StackFrame
                        FormMain.ErrorHandler(ex, System.Reflection.MethodBase.GetCurrentMethod().Name, GlobalVariables.Linenumber, f)
                    End Try
                Next
                'Update the alpaca database

                FormMain.ConnectedDB.UpdateDatabase("INSERT INTO Alpaca (Name,OnFarm) VALUES ('" & TextBox1.Text & "',TRUE)")


            Else
                MessageBox.Show("Please enter a valid name", "Error")
            End If
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
        Me.Close()
    End Sub
End Class