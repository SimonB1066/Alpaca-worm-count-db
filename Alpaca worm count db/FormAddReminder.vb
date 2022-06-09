Imports System.Globalization

Public Class FormAddReminder
    Dim Reminder As String

    Private Sub ToolStrip1_MouseHover(sender As Object, e As EventArgs) Handles ToolStrip1.MouseHover
        Me.Cursor = Cursors.Hand '
    End Sub
    Private Sub ToolStrip1_MouseLeave(sender As Object, e As EventArgs) Handles ToolStrip1.MouseLeave
        Me.Cursor = Cursors.Arrow
    End Sub


    Private Sub FormAddReminder_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CenterToScreen()
        Me.Icon = My.Resources.Form
        Dim j As Integer

        TextBox1.Visible = True
        Label4.Visible = True
        RadioButton6.Checked = True

        Try
            ComboBox1.Items.Clear()

            'Fill the combo box with animals names
            FormMain.ConnectedDB = New DataBaseFunctions()
            GlobalVariables.ds = FormMain.ConnectedDB.PopulateDataSet()
            If GlobalVariables.ds.Tables("Alpaca").Rows.Count > 0 Then
                For j = 0 To GlobalVariables.ds.Tables("Alpaca").Rows.Count - 1
                    If GlobalVariables.ds.Tables("Alpaca").Rows(j)("OnFarm") = True Then
                        ComboBox1.Items.Add(GlobalVariables.ds.Tables("Alpaca").Rows(j)("Name"))
                    End If

                Next
            End If
            ComboBox1.Sorted = True
            ComboBox1.SelectedItem = ComboBox1.Items(0).ToString

        Catch ex As Exception

            If ex.Message <> "Application is not installed." Then
                Dim st As New StackTrace(True)
                st = New StackTrace(ex, True)
                GlobalVariables.Linenumber = st.GetFrame(0).GetFileLineNumber().ToString()
                Dim f As New StackFrame
                FormMain.ErrorHandler(ex, System.Reflection.MethodBase.GetCurrentMethod().Name, GlobalVariables.Linenumber, f)
            End If
        End Try
    End Sub



    Private Sub RadioButton1_Click(sender As Object, e As EventArgs) Handles RadioButton1.Click
        TextBox1.Visible = False
        Label4.Visible = False
        Reminder = BuildString()
    End Sub

    Private Sub RadioButton2_Click(sender As Object, e As EventArgs) Handles RadioButton2.Click
        TextBox1.Visible = False
        Label4.Visible = False
        Reminder = BuildString()
    End Sub

    Private Sub RadioButton3_Click(sender As Object, e As EventArgs) Handles RadioButton3.Click
        TextBox1.Visible = False
        Label4.Visible = False
        Reminder = BuildString()
    End Sub

    Private Sub RadioButton4_Click(sender As Object, e As EventArgs) Handles RadioButton4.Click
        TextBox1.Visible = False
        Label4.Visible = False
        Reminder = BuildString()
    End Sub

    Private Sub RadioButton5_Click(sender As Object, e As EventArgs) Handles RadioButton5.Click
        TextBox1.Visible = False
        Label4.Visible = False
        Reminder = BuildString()
    End Sub

    Private Sub RadioButton6_Click(sender As Object, e As EventArgs) Handles RadioButton6.Click
        If RadioButton6.Checked Then
            TextBox1.Visible = True
            Label4.Visible = True
        End If
        Reminder = BuildString()
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        Reminder = BuildString()
    End Sub

    Private Function BuildString()
        Dim str As String = ""
        Dim str1 As String = ""
        Dim TempDate As Date
        Dim TestCount As Integer = 0

        If RadioButton1.Checked Then
            'Group test
            Try
                FormMain.ConnectedDB = New DataBaseFunctions()
                GlobalVariables.ds = FormMain.ConnectedDB.PopulateDataSet()
                If GlobalVariables.ds.Tables("TestResults").Rows.Count > 1 Then
                    For j = 0 To GlobalVariables.ds.Tables("TestResults").Rows.Count - 1
                        If Convert.ToDateTime(GlobalVariables.ds.Tables("TestResults").Rows(j)("TestDate")) > TempDate Then
                            TempDate = Convert.ToDateTime(GlobalVariables.ds.Tables("TestResults").Rows(j)("TestDate"))
                            TestCount = TestCount + 1
                            str = GlobalVariables.ds.Tables("TestResults").Rows(TestCount)("TestDate")
                        End If

                    Next
                End If


            Catch ex As Exception
                Dim st As New StackTrace(True)
                st = New StackTrace(ex, True)
                GlobalVariables.Linenumber = st.GetFrame(0).GetFileLineNumber().ToString()
                Dim f As New StackFrame
                FormMain.ErrorHandler(ex, System.Reflection.MethodBase.GetCurrentMethod().Name, GlobalVariables.Linenumber, f)
            End Try
            str = "The last test was done on the " & str & ", and a new whole herd feacal test is required."
        End If


        If RadioButton2.Checked Then
            'Follow up test
            Try
                FormMain.ConnectedDB = New DataBaseFunctions()
                GlobalVariables.ds = FormMain.ConnectedDB.PopulateDataSet()


                If GlobalVariables.ds.Tables("TestResults").Rows.Count > 0 Then
                    Dim result() As DataRow = GlobalVariables.ds.Tables("TestResults").Select("TestName='" & ComboBox1.Text & "'")
                    For j = 0 To result.Length - 1
                        If result.Length > 0 Then
                            If Convert.ToDateTime(result(j).ItemArray(1)) > TempDate Then
                                TempDate = Convert.ToDateTime(GlobalVariables.ds.Tables("TestResults").Rows(j)("TestDate"))
                                str1 = EggType(result, j)
                                If str1 = "" Then
                                    str1 = "Clear"
                                End If
                            End If
                        End If
                    Next
                End If


            Catch ex As Exception
                Dim st As New StackTrace(True)
                st = New StackTrace(ex, True)
                GlobalVariables.Linenumber = st.GetFrame(0).GetFileLineNumber().ToString()
                Dim f As New StackFrame
                FormMain.ErrorHandler(ex, System.Reflection.MethodBase.GetCurrentMethod().Name, GlobalVariables.Linenumber, f)
            End Try


            str = ComboBox1.Text & " -" & "Was last tested on " & TempDate.ToShortDateString & " and treated, it had a result of. " & vbNewLine & str1 & vbNewLine & "This is a reminder that a retest is required."
        End If


        If RadioButton3.Checked Then
            'Follow up test
            Try
                FormMain.ConnectedDB = New DataBaseFunctions()
                GlobalVariables.ds = FormMain.ConnectedDB.PopulateDataSet()


                If GlobalVariables.ds.Tables("TestResults").Rows.Count > 0 Then
                    Dim result() As DataRow = GlobalVariables.ds.Tables("TestResults").Select("TestName='" & ComboBox1.Text & "'")
                    For j = 0 To result.Length - 1
                        If result.Length > 0 Then
                            If Convert.ToDateTime(result(j).ItemArray(1)) > TempDate Then
                                TempDate = Convert.ToDateTime(GlobalVariables.ds.Tables("TestResults").Rows(j)("TestDate"))
                                str1 = EggType(result, j)
                                If str1 = "" Then
                                    str1 = "Clear"
                                End If
                            End If
                        End If
                    Next
                End If


            Catch ex As Exception
                Dim st As New StackTrace(True)
                st = New StackTrace(ex, True)
                GlobalVariables.Linenumber = st.GetFrame(0).GetFileLineNumber().ToString()
                Dim f As New StackFrame
                FormMain.ErrorHandler(ex, System.Reflection.MethodBase.GetCurrentMethod().Name, GlobalVariables.Linenumber, f)
            End Try


            str = ComboBox1.Text & " -" & "Was last tested on " & TempDate.ToShortDateString & " and had a result of. " & vbNewLine & str1 & vbNewLine & "This is a reminder that a retest is required."
        End If

        If RadioButton4.Checked Then
            str = ComboBox1.Text & " -" & RadioButton4.Text
        End If

        If RadioButton5.Checked Then
            str = RadioButton5.Text
        End If

        If RadioButton6.Checked Then
            str = TextBox1.Text
        End If

        RichTextBox1.Text = "To:     " & vbTab & GlobalVariables.Email & vbNewLine &
                            "From:   " & vbTab & "AWD.Com" & vbNewLine &
                            "Subject:" & vbTab & "Reminder from Alpaca Worm Database" & vbNewLine & vbNewLine & str
        BuildString = str
    End Function

    Public Function EggType(List() As DataRow, RowNumber As Integer)
        Dim strEggs As String = ""
        Dim Eggs As Int16 = 0
        Dim str As String = ""
        If Convert.ToInt16(List(RowNumber).ItemArray(8)) > 0 Then
            Eggs = List(RowNumber).ItemArray(8).ToString.Substring(0, 2)
            Eggs = Eggs + List(RowNumber).ItemArray(8).ToString.Substring(2, 2)
            strEggs = AddEggString(str, strEggs, Eggs, " Trichostrongyles   ")
        End If

        If Convert.ToInt16(List(RowNumber).ItemArray(9)) > 0 Then
            Eggs = List(RowNumber).ItemArray(9).ToString.Substring(0, 2)
            Eggs = Eggs + List(RowNumber).ItemArray(9).ToString.Substring(2, 2)
            strEggs = AddEggString(str, strEggs, Eggs, " Trichurius   ")
        End If

        If Convert.ToInt16(List(RowNumber).ItemArray(10)) > 0 Then
            Eggs = List(RowNumber).ItemArray(10).ToString.Substring(0, 2)
            Eggs = Eggs + List(RowNumber).ItemArray(10).ToString.Substring(2, 2)
            strEggs = AddEggString(str, strEggs, Eggs, " Nematordirus   ")
        End If

        If Convert.ToInt16(List(RowNumber).ItemArray(11)) > 0 Then
            Eggs = List(RowNumber).ItemArray(11).ToString.Substring(0, 2)
            Eggs = Eggs + List(RowNumber).ItemArray(11).ToString.Substring(2, 2)
            strEggs = AddEggString(str, strEggs, Eggs, " Capillarid   ")
        End If

        If Convert.ToInt16(List(RowNumber).ItemArray(12)) > 0 Then
            Eggs = List(RowNumber).ItemArray(12).ToString.Substring(0, 2)
            Eggs = Eggs + List(RowNumber).ItemArray(12).ToString.Substring(2, 2)
            strEggs = AddEggString(str, strEggs, Eggs, " Moniezid   ")
        End If

        If Convert.ToInt16(List(RowNumber).ItemArray(13)) > 0 Then
            Eggs = List(RowNumber).ItemArray(13).ToString.Substring(0, 2)
            Eggs = Eggs + List(RowNumber).ItemArray(13).ToString.Substring(2, 2)
            strEggs = AddEggString(str, strEggs, Eggs, " Unidentifed   ")
        End If

        If Convert.ToInt16(List(RowNumber).ItemArray(14)) > 0 Then
            Eggs = List(RowNumber).ItemArray(14).ToString.Substring(0, 2)
            Eggs = Eggs + List(RowNumber).ItemArray(14).ToString.Substring(2, 2)
            strEggs = AddEggString(str, strEggs, Eggs, " E-mac   ")
        End If

        If Convert.ToInt16(List(RowNumber).ItemArray(15)) > 0 Then
            Eggs = List(RowNumber).ItemArray(15).ToString.Substring(0, 2)
            Eggs = Eggs + List(RowNumber).ItemArray(15).ToString.Substring(2, 2)
            strEggs = AddEggString(str, strEggs, Eggs, " E-ivitaesis   ")
        End If

        If Convert.ToInt16(List(RowNumber).ItemArray(16)) > 0 Then
            Eggs = List(RowNumber).ItemArray(16).ToString.Substring(0, 2)
            Eggs = Eggs + List(RowNumber).ItemArray(16).ToString.Substring(2, 2)
            strEggs = AddEggString(str, strEggs, Eggs, " E-alpacae   ")
        End If

        If Convert.ToInt16(List(RowNumber).ItemArray(17)) > 0 Then
            Eggs = List(RowNumber).ItemArray(17).ToString.Substring(0, 2)
            Eggs = Eggs + List(RowNumber).ItemArray(17).ToString.Substring(2, 2)
            strEggs = AddEggString(str, strEggs, Eggs, " E-lamae   ")
        End If

        If Convert.ToInt16(List(RowNumber).ItemArray(18)) > 0 Then
            Eggs = List(RowNumber).ItemArray(18).ToString.Substring(0, 2)
            Eggs = Eggs + List(RowNumber).ItemArray(18).ToString.Substring(2, 2)
            strEggs = AddEggString(str, strEggs, Eggs, " E-punoensis   ")
        End If

        If Convert.ToInt16(List(RowNumber).ItemArray(19)) > 0 Then
            Eggs = List(RowNumber).ItemArray(19).ToString.Substring(0, 2)
            Eggs = Eggs + List(RowNumber).ItemArray(19).ToString.Substring(2, 2)
            strEggs = AddEggString(str, strEggs, Eggs, " Unidentifed   ")
        End If
        EggType = strEggs
    End Function

    Private Function AddEggString(str As String, strEggs As String, Eggs As Integer, name As String)
        Dim strLength As Integer = (Len(str) + Len(strEggs) + Len(name))
        strEggs = strEggs & Eggs.ToString & name

        AddEggString = strEggs
    End Function



    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        RichTextBox1.Text = "To:     " & vbTab & GlobalVariables.Email & vbNewLine &
                            "From:   " & vbTab & "AWD.Com" & vbNewLine &
                            "Subject:" & vbTab & "Reminder from Alpaca Worm Database" & vbNewLine & vbNewLine & TextBox1.Text
        Reminder = TextBox1.Text
    End Sub

    Private Sub ToolStripButton4_Click(sender As Object, e As EventArgs) Handles ToolStripButton4.Click
        Dim dtUK As Date = DateTime.ParseExact(DateTimePicker1.Value.ToShortDateString, "dd/MM/yyyy", CultureInfo.InvariantCulture)
        FormMain.ConnectedDB.UpdateDatabase("INSERT INTO Reminder (ReminderDate,ReminderNote,ReminderNotify,ReminderEmail,ReminderDone) VALUES (#" & dtUK & "#,'" & Reminder & "',TRUE,TRUE,FALSE)")
        Close()
    End Sub

    Private Sub ToolStripButton2_Click(sender As Object, e As EventArgs) Handles ToolStripButton2.Click
        Close()
    End Sub
End Class