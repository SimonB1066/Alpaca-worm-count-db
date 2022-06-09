Public Class FormTestHistoryStoll
    Dim page As Integer = 0
    Dim foundRow As DataRow
    Dim list As New List(Of Integer)

    Private Sub ToolStrip1_MouseHover(sender As Object, e As EventArgs) Handles ToolStrip1.MouseHover
        Me.Cursor = Cursors.Hand '
    End Sub
    Private Sub ToolStrip1_MouseLeave(sender As Object, e As EventArgs) Handles ToolStrip1.MouseLeave
        Me.Cursor = Cursors.Arrow
    End Sub

    Private Sub FormTest_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        CenterToScreen()
        Me.Icon = My.Resources.Form

        Dim NumberOfTests As Integer
        If IsDBNull(GlobalVariables.AlpacaTestDate) Or GlobalVariables.AlpacaTestDate = "" Then
            MessageBox.Show("No test date found", "Warning", MessageBoxButtons.OK)
            Me.Close()
            Exit Sub
        End If
        Try


            list.Clear()
            page = 0

            For j = 0 To GlobalVariables.ds.Tables("TestResults").Rows.Count - 1
                If GlobalVariables.ds.Tables("TestResults").Rows(j)("TestNumberName") = GlobalVariables.AlpacaTestDate Then
                    If GlobalVariables.ds.Tables("TestResults").Rows(j)("TestName") = GlobalVariables.AlpacaName Then
                        list.Add(j)
                        NumberOfTests = NumberOfTests + 1
                    End If
                End If
            Next
            If NumberOfTests > 1 Then
                Label12.Text = "There has been multiple test (" & NumberOfTests.ToString & ") in this test period for this animal"
                ToolStripButton1.Enabled = True
                ToolStripButton3.Enabled = False
            Else
                Label12.Text = ""
                ToolStripButton1.Enabled = False
                ToolStripButton3.Enabled = False
            End If


            GetHistory()


        Catch

            Dim result1 As DialogResult
            If GlobalVariables.TestGroupName = GlobalVariables.AlpacaTestDate Then
                result1 = MessageBox.Show("No test date found, do you want to add a test result?", "Warning", MessageBoxButtons.YesNo)
            Else
                'result1 = MessageBox.Show("No historical data found?", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Me.Close()
                Exit Sub
            End If

            If result1 = Windows.Forms.DialogResult.Yes Then
                FormTest.ShowDialog()
            End If
            Me.Close()
            Exit Sub
        End Try
    End Sub

    Public Sub GetHistory()
        Try

            FormMain.ConnectedDB = New DataBaseFunctions()              'Open the database connection
            GlobalVariables.ds = FormMain.ConnectedDB.PopulateDataSet()        'Put the datainto a dataset
            foundRow = GlobalVariables.ds.Tables("TestResults").Rows(list(page)) '.Select("TestNumberName='" & GlobalVariables.AlpacaTestDate & "' AND TestName='" & GlobalVariables.AlpacaName & "'")(0)
            TextBox3.Text = foundRow.Item(1)
            TextBox1.Text = foundRow.Item(2)
            ToolStripLabel1.Text = "Test number " & (list.Count - page).ToString

            Select Case foundRow.Item(3)
                Case "1"
                    TextBox2.Text = "Grade 1 Separate hard pellets"
                Case "2"
                    TextBox2.Text = "Grade 2 Pelleted and shaped"
                Case "3"
                    TextBox2.Text = "Grade 3 Shaped but with soft pellets"
                Case "4"
                    TextBox2.Text = "Grade 4 Shaped but no structure"
                Case "5"
                    TextBox2.Text = "Grade 5 No shape or structure"
                Case "6"
                    TextBox2.Text = "Grade 6 Watery no solids"

            End Select

            Label67.Text = foundRow.Item(6)
            Label66.Text = foundRow.Item(7)
            Chamber1_1.Text = System.Convert.ToInt32(foundRow.Item(8).ToString.Substring(0, 2))
            Chamber1_2.Text = System.Convert.ToInt32(foundRow.Item(9).ToString.Substring(0, 2))
            Chamber1_3.Text = System.Convert.ToInt32(foundRow.Item(10).ToString.Substring(0, 2))
            Chamber1_4.Text = System.Convert.ToInt32(foundRow.Item(11).ToString.Substring(0, 2))
            Chamber1_5.Text = System.Convert.ToInt32(foundRow.Item(12).ToString.Substring(0, 2))
            Chamber1_6.Text = System.Convert.ToInt32(foundRow.Item(13).ToString.Substring(0, 2))
            Chamber1_7.Text = System.Convert.ToInt32(foundRow.Item(14).ToString.Substring(0, 2))
            Chamber1_8.Text = System.Convert.ToInt32(foundRow.Item(15).ToString.Substring(0, 2))
            Chamber1_9.Text = System.Convert.ToInt32(foundRow.Item(16).ToString.Substring(0, 2))
            Chamber1_10.Text = System.Convert.ToInt32(foundRow.Item(17).ToString.Substring(0, 2))
            Chamber1_11.Text = System.Convert.ToInt32(foundRow.Item(18).ToString.Substring(0, 2))
            Chamber1_12.Text = System.Convert.ToInt32(foundRow.Item(19).ToString.Substring(0, 2))


            EPG1.Text = System.Convert.ToInt32(Chamber1_1.Text)
            EPG2.Text = System.Convert.ToInt32(Chamber1_2.Text)
            EPG3.Text = System.Convert.ToInt32(Chamber1_3.Text)
            EPG4.Text = System.Convert.ToInt32(Chamber1_4.Text)
            EPG5.Text = System.Convert.ToInt32(Chamber1_5.Text)
            EPG6.Text = System.Convert.ToInt32(Chamber1_6.Text)

            OPG1.Text = System.Convert.ToInt32(Chamber1_7.Text)
            OPG2.Text = System.Convert.ToInt32(Chamber1_8.Text)
            OPG3.Text = System.Convert.ToInt32(Chamber1_9.Text)
            OPG4.Text = System.Convert.ToInt32(Chamber1_10.Text)
            OPG5.Text = System.Convert.ToInt32(Chamber1_11.Text)
            OPG6.Text = System.Convert.ToInt32(Chamber1_12.Text)

            Me.Refresh()
            Me.Update()

        Catch ex As Exception
            Dim st As New StackTrace(True)
            st = New StackTrace(ex, True)
            GlobalVariables.Linenumber = st.GetFrame(0).GetFileLineNumber().ToString()
            Dim f As New StackFrame
            FormMain.ErrorHandler(ex, System.Reflection.MethodBase.GetCurrentMethod().Name, GlobalVariables.Linenumber, f)
        End Try

    End Sub
    Private Sub Form1_Shown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shown
        Try
            If GlobalVariables.SelfClose Then
                Using bmp = New Bitmap(Me.Width, Me.Height)
                    Me.DrawToBitmap(bmp, New Rectangle(0, 0, bmp.Width, bmp.Height))
                    bmp.Save(GlobalVariables.DbDriveLocation & "screenshot.png")
                    'bmp.Dispose()
                End Using
                GlobalVariables.SelfClose = False
                Me.Close()
            End If
        Catch

        End Try

    End Sub







    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles Button2.Click

        Try

            Dim foundRow As DataRow
            Dim tempIndex As Integer
            FormMain.ConnectedDB = New DataBaseFunctions()              'Open the database connection
            GlobalVariables.ds = FormMain.ConnectedDB.PopulateDataSet()        'Put the datainto a dataset
            foundRow = GlobalVariables.ds.Tables("TestResults").Select("TestNumberName='" & GlobalVariables.AlpacaTestDate & "' AND TestName='" & GlobalVariables.AlpacaName & "'")(0)
            tempIndex = GlobalVariables.ds.Tables("TestResults").Rows.IndexOf(foundRow)
            GlobalVariables.ds.Tables("TestResults").Rows(tempIndex).BeginEdit()
            GlobalVariables.ds.Tables("TestResults").Rows(tempIndex)(20) = "Treated"
            GlobalVariables.ds.Tables("TestResults").Rows(tempIndex).EndEdit()
            Me.Refresh()
            Me.Update()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try


    End Sub




    Private Sub ToolStripButton2_Click(sender As Object, e As EventArgs) Handles ToolStripButton2.Click
        Me.Close()
    End Sub

    Private Sub ToolStripButton1_Click(sender As Object, e As EventArgs) Handles ToolStripButton1.Click
        If page < list.Count - 1 Then
            page = page + 1
            ToolStripButton1.Enabled = True
            ToolStripButton3.Enabled = True
        End If
        If page = list.Count - 1 Then
            ToolStripButton1.Enabled = False
        End If
        GetHistory()
    End Sub

    Private Sub ToolStripButton3_Click(sender As Object, e As EventArgs) Handles ToolStripButton3.Click
        If page >= 0 Then
            page = page - 1
            ToolStripButton1.Enabled = True
            ToolStripButton3.Enabled = True
        End If
        If page = 0 Then
            ToolStripButton3.Enabled = False
        End If
        GetHistory()
    End Sub

    Private Sub ToolStripButton4_Click(sender As Object, e As EventArgs) Handles ToolStripButton4.Click

        Dim result As MsgBoxResult = MsgBoxResult.Cancel
        result = MsgBox("Are you sure you want to delete this record?", vbYesNo & MsgBoxStyle.Information & vbSystem, "Warning")
        If result = MsgBoxResult.Yes Then
            FormMain.ConnectedDB = New DataBaseFunctions()              'Open the database connection
            GlobalVariables.ds = FormMain.ConnectedDB.PopulateDataSet()        'Put the datainto a dataset
            foundRow = GlobalVariables.ds.Tables("TestResults").Rows(list(page))
            FormMain.ConnectedDB.UpdateDatabase("DELETE FROM TestResults WHERE TestID=" & CLng(Replace(foundRow.Item(0), " ", "")))
            Close()
            FormMain.loadFormWithAnimals()
        End If
    End Sub
End Class