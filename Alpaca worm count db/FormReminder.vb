Imports System.ComponentModel



Public Class FormReminder

    Dim strMonth As String = ""
    Dim iMonth As Integer = 1
    Dim strYear As String = ""
    Dim customToolTip As New ToolTip

    Private Sub ToolStrip1_MouseHover(sender As Object, e As EventArgs) Handles ToolStrip1.MouseHover
        Me.Cursor = Cursors.Hand '
    End Sub
    Private Sub ToolStrip1_MouseLeave(sender As Object, e As EventArgs) Handles ToolStrip1.MouseLeave
        Me.Cursor = Cursors.Arrow
    End Sub


    Private Sub FormReminder_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Try


            CenterToScreen()
            Me.Icon = My.Resources.Form

            customToolTip.ShowAlways = True
            customToolTip.InitialDelay = 1
            customToolTip.OwnerDraw = True
            customToolTip.UseFading = True
            customToolTip.ReshowDelay = 500
            customToolTip.AutoPopDelay = 10000
            customToolTip.UseFading = True
            customToolTip.ToolTipTitle = "Reminder Event"
            customToolTip.ToolTipIcon = ToolTipIcon.Info
            customToolTip.IsBalloon = True
            customToolTip.BackColor = Color.Beige


            FillDateGridView()
            Dim theDate As New Date
            theDate = Now
            strYear = theDate.Year.ToString
            iMonth = theDate.Month
            strMonth = Now.ToString("MMMM")

            FillTheCalander()


        Catch ex As Exception
            Dim st As New StackTrace(True)
            st = New StackTrace(ex, True)
            GlobalVariables.Linenumber = st.GetFrame(0).GetFileLineNumber().ToString()
            Dim f As New StackFrame
            FormMain.ErrorHandler(ex, System.Reflection.MethodBase.GetCurrentMethod().Name, GlobalVariables.Linenumber, f)
        End Try
    End Sub





    Private Sub FormReminder_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        DataGridView1.DataSource = Nothing
        GlobalVariables.ds.Dispose()
    End Sub

    Private Sub FillTheCalander()

        Button5.Text = strMonth & " " & strYear
        Dim ndate As DateTime
        Dim firstDay = New DateTime(strYear, iMonth, 1)
        Dim dayOfFirstDay As Integer = firstDay.DayOfWeek
        Dim days As Integer = System.DateTime.DaysInMonth(strYear, iMonth)
        If dayOfFirstDay = 0 Then
            dayOfFirstDay = 7

        End If

        Dim TextBox As TextBox
        Dim name As String = ""



        For I = 1 To 42
            name = "TextBox" & I.ToString
            TextBox = Me.Controls.Item(name)
            TextBox.Text = CalanderDay(I, dayOfFirstDay, days)
            Try
                ndate = CDate(TextBox.Text & "/" & iMonth.ToString & "/" & strYear)
                If Now.Date = ndate.Date And TextBox.BackColor = Color.LightGray Then
                    TextBox.BackColor = Color.LimeGreen
                End If

                Dim number As Integer = 0
                For Each row As DataGridViewRow In DataGridView1.Rows
                    Dim now = Date.Now

                    If row.Cells("ReminderDate").Value = ndate And TextBox.BackColor <> Color.DarkGray Then
                        number = number + 1
                        TextBox.Text = TextBox.Text & vbNewLine & vbNewLine & "REMINDER " & number.ToString & "- " & row.Cells("ReminderNote").Value

                    End If
                Next
            Catch
            End Try
        Next



    End Sub

    Private Function CalanderDay(Position As Integer, FirstDay As Integer, LastDay As Integer)
        CalanderDay = (Position - FirstDay) + 1
        If CalanderDay <= 0 Then
            Dim PrevMonth As Integer = 1
            PrevMonth = iMonth - 1
            If PrevMonth <= 0 Then
                PrevMonth = 12
            End If
            If PrevMonth > 12 Then
                PrevMonth = 1
            End If
            Dim days As Integer = System.DateTime.DaysInMonth(strYear, PrevMonth) + 1
            CalanderDay = days - (FirstDay - Position)
            Dim TextBox As TextBox
            Dim name As String = ""
            name = "TextBox" & Position.ToString
            TextBox = Me.Controls.Item(name)
            TextBox.BackColor = Color.DarkGray
            Exit Function
        Else
            Dim TextBox As TextBox
            Dim name As String = ""
            name = "TextBox" & Position.ToString
            TextBox = Me.Controls.Item(name)
            TextBox.BackColor = Color.LightGray

        End If

        If CalanderDay > LastDay Then
            CalanderDay = CalanderDay - LastDay
            Dim TextBox As TextBox
            Dim name As String = ""
            name = "TextBox" & Position.ToString
            TextBox = Me.Controls.Item(name)
            TextBox.BackColor = Color.DarkGray
        Else
            Dim TextBox As TextBox
            Dim name As String = ""
            name = "TextBox" & Position.ToString
            TextBox = Me.Controls.Item(name)
            TextBox.BackColor = Color.LightGray
        End If

    End Function

    Private Sub FillDateGridView()
        Try
            FormMain.ConnectedDB = New DataBaseFunctions()
            GlobalVariables.ds = FormMain.ConnectedDB.PopulateDataSet()
            Dim dv = GlobalVariables.ds.Tables("Reminder")
            DataGridView1.DataSource = dv
            DataGridView1.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
            DataGridView1.Sort(DataGridView1.Columns(1), ComponentModel.ListSortDirection.Ascending)
            DataGridView1.Columns(1).HeaderText = "Date"
            DataGridView1.Columns(2).HeaderText = "Memo"
            DataGridView1.Columns(5).HeaderText = "Complete"
            DataGridView1.Columns("ReminderID").Visible = False
            DataGridView1.Columns("ReminderNotify").Visible = False
            DataGridView1.Columns("ReminderEmail").Visible = False
            DataGridView1.Columns("ReminderDone").Visible = True
            DataGridView1.Columns("ReminderDate").Width = 100
            DataGridView1.Columns("ReminderNote").Width = 400
            DataGridView1.RowTemplate.Height = 100
            DataGridView1.Sort(DataGridView1.Columns("ReminderDate"), System.ComponentModel.ListSortDirection.Ascending)
            DataGridView1.DefaultCellStyle.Font = New Font("Tahoma", 10)
            DataGridView1.DefaultCellStyle.WrapMode = DataGridViewTriState.True
            DataGridView1.ColumnHeadersDefaultCellStyle.Font = New Font("Tahoma", 10)


        Catch ex As Exception
            Dim st As New StackTrace(True)
            st = New StackTrace(ex, True)
            GlobalVariables.Linenumber = st.GetFrame(0).GetFileLineNumber().ToString()
            Dim f As New StackFrame
            FormMain.ErrorHandler(ex, System.Reflection.MethodBase.GetCurrentMethod().Name, GlobalVariables.Linenumber, f)
        End Try
    End Sub

    Private Sub DataGridView1_DataBindingComplete(sender As Object, e As DataGridViewBindingCompleteEventArgs) Handles DataGridView1.DataBindingComplete
        For Each row As DataGridViewRow In DataGridView1.Rows
            Dim now = Date.Now
            Dim ndate = DateTime.Parse(row.Cells("ReminderDate").Value.ToString)

            If row.Cells("ReminderDone").Value = "TRUE" Then
                row.DefaultCellStyle.BackColor = Color.LightGreen
            Else
                row.DefaultCellStyle.BackColor = Color.LightBlue
            End If
            'If row.Cells("ReminderDone").Value = "TRUE" And Not CheckBox1.Checked Then
            'DataGridView1.Rows.Remove(DataGridView1.Rows(CInt(row.Index)))
            'End If
        Next
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs)
        FillDateGridView()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

        iMonth = iMonth - 1
        If iMonth <= 0 Then
            iMonth = 12
            strYear = (CInt(strYear) - 1).ToString
        End If
        strMonth = MonthName(iMonth)

        FillTheCalander()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        iMonth = iMonth + 1
        If iMonth > 12 Then
            iMonth = 1
            strYear = (CInt(strYear) + 1).ToString
        End If
        strMonth = MonthName(iMonth)

        FillTheCalander()
    End Sub

    Private Function AddDate(st As String)
        AddDate = st.Substring(0, 2) & "/" & iMonth.ToString & "/" & strYear & st.Substring(2, Len(st) - 2)
    End Function

    Private Sub TextBox1_MouseEnter(sender As Object, e As EventArgs) Handles TextBox1.MouseEnter
        Try
            Dim TextBox As TextBox

            TextBox = Me.Controls.Item(sender.name)
            If Len(TextBox.Text) > 3 Then
                customToolTip.SetToolTip(TextBox, AddDate(TextBox.Text))
            End If
        Catch ex As Exception
            Dim st As New StackTrace(True)
            st = New StackTrace(ex, True)
            GlobalVariables.Linenumber = st.GetFrame(0).GetFileLineNumber().ToString()
            Dim f As New StackFrame
            FormMain.ErrorHandler(ex, System.Reflection.MethodBase.GetCurrentMethod().Name, GlobalVariables.Linenumber, f)
        End Try
    End Sub

    Private Sub TextBox2_MouseEnter(sender As Object, e As EventArgs) Handles TextBox2.MouseEnter
        Try
            Dim TextBox As TextBox
            TextBox = Me.Controls.Item(sender.name)
            If Len(TextBox.Text) > 3 Then
                customToolTip.SetToolTip(TextBox, AddDate(TextBox.Text))
            End If
        Catch ex As Exception
            Dim st As New StackTrace(True)
            st = New StackTrace(ex, True)
            GlobalVariables.Linenumber = st.GetFrame(0).GetFileLineNumber().ToString()
            Dim f As New StackFrame
            FormMain.ErrorHandler(ex, System.Reflection.MethodBase.GetCurrentMethod().Name, GlobalVariables.Linenumber, f)
        End Try
    End Sub

    Private Sub TextBox3_MouseEnter(sender As Object, e As EventArgs) Handles TextBox3.MouseEnter
        Try
            Dim TextBox As TextBox
            TextBox = Me.Controls.Item(sender.name)
            If Len(TextBox.Text) > 3 Then
                customToolTip.SetToolTip(TextBox, AddDate(TextBox.Text))
            End If
        Catch ex As Exception
            Dim st As New StackTrace(True)
            st = New StackTrace(ex, True)
            GlobalVariables.Linenumber = st.GetFrame(0).GetFileLineNumber().ToString()
            Dim f As New StackFrame
            FormMain.ErrorHandler(ex, System.Reflection.MethodBase.GetCurrentMethod().Name, GlobalVariables.Linenumber, f)
        End Try
    End Sub

    Private Sub TextBox4_MouseEnter(sender As Object, e As EventArgs) Handles TextBox4.MouseEnter
        Try
            Dim TextBox As TextBox
            TextBox = Me.Controls.Item(sender.name)
            If Len(TextBox.Text) > 3 Then
                customToolTip.SetToolTip(TextBox, AddDate(TextBox.Text))
            End If
        Catch ex As Exception
            Dim st As New StackTrace(True)
            st = New StackTrace(ex, True)
            GlobalVariables.Linenumber = st.GetFrame(0).GetFileLineNumber().ToString()
            Dim f As New StackFrame
            FormMain.ErrorHandler(ex, System.Reflection.MethodBase.GetCurrentMethod().Name, GlobalVariables.Linenumber, f)
        End Try
    End Sub

    Private Sub TextBox5_MouseEnter(sender As Object, e As EventArgs) Handles TextBox5.MouseEnter
        Try
            Dim TextBox As TextBox
            TextBox = Me.Controls.Item(sender.name)
            If Len(TextBox.Text) > 3 Then
                customToolTip.SetToolTip(TextBox, AddDate(TextBox.Text))
            End If
        Catch ex As Exception
            Dim st As New StackTrace(True)
            st = New StackTrace(ex, True)
            GlobalVariables.Linenumber = st.GetFrame(0).GetFileLineNumber().ToString()
            Dim f As New StackFrame
            FormMain.ErrorHandler(ex, System.Reflection.MethodBase.GetCurrentMethod().Name, GlobalVariables.Linenumber, f)
        End Try
    End Sub

    Private Sub TextBox6_MouseEnter(sender As Object, e As EventArgs) Handles TextBox6.MouseEnter
        Try
            Dim TextBox As TextBox
            TextBox = Me.Controls.Item(sender.name)
            If Len(TextBox.Text) > 3 Then
                customToolTip.SetToolTip(TextBox, AddDate(TextBox.Text))
            End If
        Catch ex As Exception
            Dim st As New StackTrace(True)
            st = New StackTrace(ex, True)
            GlobalVariables.Linenumber = st.GetFrame(0).GetFileLineNumber().ToString()
            Dim f As New StackFrame
            FormMain.ErrorHandler(ex, System.Reflection.MethodBase.GetCurrentMethod().Name, GlobalVariables.Linenumber, f)
        End Try
    End Sub

    Private Sub TextBox7_MouseEnter(sender As Object, e As EventArgs) Handles TextBox7.MouseEnter
        Try
            Dim TextBox As TextBox
            TextBox = Me.Controls.Item(sender.name)
            If Len(TextBox.Text) > 3 Then
                customToolTip.SetToolTip(TextBox, AddDate(TextBox.Text))
            End If
        Catch ex As Exception
            Dim st As New StackTrace(True)
            st = New StackTrace(ex, True)
            GlobalVariables.Linenumber = st.GetFrame(0).GetFileLineNumber().ToString()
            Dim f As New StackFrame
            FormMain.ErrorHandler(ex, System.Reflection.MethodBase.GetCurrentMethod().Name, GlobalVariables.Linenumber, f)
        End Try
    End Sub

    Private Sub TextBox8_MouseEnter(sender As Object, e As EventArgs) Handles TextBox8.MouseEnter
        Try
            Dim TextBox As TextBox
            TextBox = Me.Controls.Item(sender.name)
            If Len(TextBox.Text) > 3 Then
                customToolTip.SetToolTip(TextBox, AddDate(TextBox.Text))
            End If
        Catch ex As Exception
            Dim st As New StackTrace(True)
            st = New StackTrace(ex, True)
            GlobalVariables.Linenumber = st.GetFrame(0).GetFileLineNumber().ToString()
            Dim f As New StackFrame
            FormMain.ErrorHandler(ex, System.Reflection.MethodBase.GetCurrentMethod().Name, GlobalVariables.Linenumber, f)
        End Try
    End Sub

    Private Sub TextBox9_MouseEnter(sender As Object, e As EventArgs) Handles TextBox9.MouseEnter
        Try
            Dim TextBox As TextBox
            TextBox = Me.Controls.Item(sender.name)
            If Len(TextBox.Text) > 3 Then
                customToolTip.SetToolTip(TextBox, AddDate(TextBox.Text))
            End If
        Catch ex As Exception
            Dim st As New StackTrace(True)
            st = New StackTrace(ex, True)
            GlobalVariables.Linenumber = st.GetFrame(0).GetFileLineNumber().ToString()
            Dim f As New StackFrame
            FormMain.ErrorHandler(ex, System.Reflection.MethodBase.GetCurrentMethod().Name, GlobalVariables.Linenumber, f)
        End Try
    End Sub

    Private Sub TextBox10_MouseEnter(sender As Object, e As EventArgs) Handles TextBox10.MouseEnter
        Try
            Dim TextBox As TextBox
            TextBox = Me.Controls.Item(sender.name)
            If Len(TextBox.Text) > 3 Then
                customToolTip.SetToolTip(TextBox, AddDate(TextBox.Text))
            End If
        Catch ex As Exception
            Dim st As New StackTrace(True)
            st = New StackTrace(ex, True)
            GlobalVariables.Linenumber = st.GetFrame(0).GetFileLineNumber().ToString()
            Dim f As New StackFrame
            FormMain.ErrorHandler(ex, System.Reflection.MethodBase.GetCurrentMethod().Name, GlobalVariables.Linenumber, f)
        End Try
    End Sub

    Private Sub TextBox11_MouseEnter(sender As Object, e As EventArgs) Handles TextBox11.MouseEnter
        Try
            Dim TextBox As TextBox
            TextBox = Me.Controls.Item(sender.name)
            If Len(TextBox.Text) > 3 Then
                customToolTip.SetToolTip(TextBox, AddDate(TextBox.Text))
            End If
        Catch ex As Exception
            Dim st As New StackTrace(True)
            st = New StackTrace(ex, True)
            GlobalVariables.Linenumber = st.GetFrame(0).GetFileLineNumber().ToString()
            Dim f As New StackFrame
            FormMain.ErrorHandler(ex, System.Reflection.MethodBase.GetCurrentMethod().Name, GlobalVariables.Linenumber, f)
        End Try
    End Sub

    Private Sub TextBox12_MouseEnter(sender As Object, e As EventArgs) Handles TextBox12.MouseEnter
        Try
            Dim TextBox As TextBox
            TextBox = Me.Controls.Item(sender.name)
            If Len(TextBox.Text) > 3 Then
                customToolTip.SetToolTip(TextBox, AddDate(TextBox.Text))
            End If
        Catch ex As Exception
            Dim st As New StackTrace(True)
            st = New StackTrace(ex, True)
            GlobalVariables.Linenumber = st.GetFrame(0).GetFileLineNumber().ToString()
            Dim f As New StackFrame
            FormMain.ErrorHandler(ex, System.Reflection.MethodBase.GetCurrentMethod().Name, GlobalVariables.Linenumber, f)
        End Try
    End Sub

    Private Sub TextBox13_MouseEnter(sender As Object, e As EventArgs) Handles TextBox13.MouseEnter
        Try
            Dim TextBox As TextBox
            TextBox = Me.Controls.Item(sender.name)
            If Len(TextBox.Text) > 3 Then
                customToolTip.SetToolTip(TextBox, AddDate(TextBox.Text))
            End If
        Catch ex As Exception
            Dim st As New StackTrace(True)
            st = New StackTrace(ex, True)
            GlobalVariables.Linenumber = st.GetFrame(0).GetFileLineNumber().ToString()
            Dim f As New StackFrame
            FormMain.ErrorHandler(ex, System.Reflection.MethodBase.GetCurrentMethod().Name, GlobalVariables.Linenumber, f)
        End Try
    End Sub

    Private Sub TextBox14_MouseEnter(sender As Object, e As EventArgs) Handles TextBox14.MouseEnter
        Try
            Dim TextBox As TextBox
            TextBox = Me.Controls.Item(sender.name)
            If Len(TextBox.Text) > 3 Then
                customToolTip.SetToolTip(TextBox, AddDate(TextBox.Text))
            End If
        Catch ex As Exception
            Dim st As New StackTrace(True)
            st = New StackTrace(ex, True)
            GlobalVariables.Linenumber = st.GetFrame(0).GetFileLineNumber().ToString()
            Dim f As New StackFrame
            FormMain.ErrorHandler(ex, System.Reflection.MethodBase.GetCurrentMethod().Name, GlobalVariables.Linenumber, f)
        End Try
    End Sub

    Private Sub TextBox15_MouseEnter(sender As Object, e As EventArgs) Handles TextBox15.MouseEnter
        Try
            Dim TextBox As TextBox
            TextBox = Me.Controls.Item(sender.name)
            If Len(TextBox.Text) > 3 Then
                customToolTip.SetToolTip(TextBox, AddDate(TextBox.Text))
            End If
        Catch ex As Exception
            Dim st As New StackTrace(True)
            st = New StackTrace(ex, True)
            GlobalVariables.Linenumber = st.GetFrame(0).GetFileLineNumber().ToString()
            Dim f As New StackFrame
            FormMain.ErrorHandler(ex, System.Reflection.MethodBase.GetCurrentMethod().Name, GlobalVariables.Linenumber, f)
        End Try
    End Sub

    Private Sub TextBox16_MouseEnter(sender As Object, e As EventArgs) Handles TextBox16.MouseEnter
        Try
            Dim TextBox As TextBox
            TextBox = Me.Controls.Item(sender.name)
            If Len(TextBox.Text) > 3 Then
                customToolTip.SetToolTip(TextBox, AddDate(TextBox.Text))
            End If
        Catch ex As Exception
            Dim st As New StackTrace(True)
            st = New StackTrace(ex, True)
            GlobalVariables.Linenumber = st.GetFrame(0).GetFileLineNumber().ToString()
            Dim f As New StackFrame
            FormMain.ErrorHandler(ex, System.Reflection.MethodBase.GetCurrentMethod().Name, GlobalVariables.Linenumber, f)
        End Try
    End Sub

    Private Sub TextBox17_MouseEnter(sender As Object, e As EventArgs) Handles TextBox17.MouseEnter
        Try
            Dim TextBox As TextBox
            TextBox = Me.Controls.Item(sender.name)
            If Len(TextBox.Text) > 3 Then
                customToolTip.SetToolTip(TextBox, AddDate(TextBox.Text))
            End If
        Catch ex As Exception
            Dim st As New StackTrace(True)
            st = New StackTrace(ex, True)
            GlobalVariables.Linenumber = st.GetFrame(0).GetFileLineNumber().ToString()
            Dim f As New StackFrame
            FormMain.ErrorHandler(ex, System.Reflection.MethodBase.GetCurrentMethod().Name, GlobalVariables.Linenumber, f)
        End Try
    End Sub

    Private Sub TextBox18_MouseEnter(sender As Object, e As EventArgs) Handles TextBox18.MouseEnter
        Try
            Dim TextBox As TextBox
            TextBox = Me.Controls.Item(sender.name)
            If Len(TextBox.Text) > 3 Then
                customToolTip.SetToolTip(TextBox, AddDate(TextBox.Text))
            End If
        Catch ex As Exception
            Dim st As New StackTrace(True)
            st = New StackTrace(ex, True)
            GlobalVariables.Linenumber = st.GetFrame(0).GetFileLineNumber().ToString()
            Dim f As New StackFrame
            FormMain.ErrorHandler(ex, System.Reflection.MethodBase.GetCurrentMethod().Name, GlobalVariables.Linenumber, f)
        End Try
    End Sub

    Private Sub TextBox19_MouseEnter(sender As Object, e As EventArgs) Handles TextBox19.MouseEnter
        Try
            Dim TextBox As TextBox
            TextBox = Me.Controls.Item(sender.name)
            If Len(TextBox.Text) > 3 Then
                customToolTip.SetToolTip(TextBox, AddDate(TextBox.Text))
            End If
        Catch ex As Exception
            Dim st As New StackTrace(True)
            st = New StackTrace(ex, True)
            GlobalVariables.Linenumber = st.GetFrame(0).GetFileLineNumber().ToString()
            Dim f As New StackFrame
            FormMain.ErrorHandler(ex, System.Reflection.MethodBase.GetCurrentMethod().Name, GlobalVariables.Linenumber, f)
        End Try
    End Sub

    Private Sub TextBox20_MouseEnter(sender As Object, e As EventArgs) Handles TextBox20.MouseEnter
        Try
            Dim TextBox As TextBox
            TextBox = Me.Controls.Item(sender.name)
            If Len(TextBox.Text) > 3 Then
                customToolTip.SetToolTip(TextBox, AddDate(TextBox.Text))
            End If
        Catch ex As Exception
            Dim st As New StackTrace(True)
            st = New StackTrace(ex, True)
            GlobalVariables.Linenumber = st.GetFrame(0).GetFileLineNumber().ToString()
            Dim f As New StackFrame
            FormMain.ErrorHandler(ex, System.Reflection.MethodBase.GetCurrentMethod().Name, GlobalVariables.Linenumber, f)
        End Try
    End Sub


    Private Sub TextBox21_MouseEnter(sender As Object, e As EventArgs) Handles TextBox21.MouseEnter
        Try
            Dim TextBox As TextBox
            TextBox = Me.Controls.Item(sender.name)
            If Len(TextBox.Text) > 3 Then
                customToolTip.SetToolTip(TextBox, AddDate(TextBox.Text))
            End If
        Catch ex As Exception
            Dim st As New StackTrace(True)
            st = New StackTrace(ex, True)
            GlobalVariables.Linenumber = st.GetFrame(0).GetFileLineNumber().ToString()
            Dim f As New StackFrame
            FormMain.ErrorHandler(ex, System.Reflection.MethodBase.GetCurrentMethod().Name, GlobalVariables.Linenumber, f)
        End Try
    End Sub

    Private Sub TextBox22_MouseEnter(sender As Object, e As EventArgs) Handles TextBox22.MouseEnter
        Try
            Dim TextBox As TextBox
            TextBox = Me.Controls.Item(sender.name)
            If Len(TextBox.Text) > 3 Then
                customToolTip.SetToolTip(TextBox, AddDate(TextBox.Text))
            End If
        Catch ex As Exception
            Dim st As New StackTrace(True)
            st = New StackTrace(ex, True)
            GlobalVariables.Linenumber = st.GetFrame(0).GetFileLineNumber().ToString()
            Dim f As New StackFrame
            FormMain.ErrorHandler(ex, System.Reflection.MethodBase.GetCurrentMethod().Name, GlobalVariables.Linenumber, f)
        End Try
    End Sub

    Private Sub TextBox23_MouseEnter(sender As Object, e As EventArgs) Handles TextBox23.MouseEnter
        Try
            Dim TextBox As TextBox
            TextBox = Me.Controls.Item(sender.name)
            If Len(TextBox.Text) > 3 Then
                customToolTip.SetToolTip(TextBox, AddDate(TextBox.Text))
            End If
        Catch ex As Exception
            Dim st As New StackTrace(True)
            st = New StackTrace(ex, True)
            GlobalVariables.Linenumber = st.GetFrame(0).GetFileLineNumber().ToString()
            Dim f As New StackFrame
            FormMain.ErrorHandler(ex, System.Reflection.MethodBase.GetCurrentMethod().Name, GlobalVariables.Linenumber, f)
        End Try
    End Sub

    Private Sub TextBox24_MouseEnter(sender As Object, e As EventArgs) Handles TextBox24.MouseEnter
        Try
            Dim TextBox As TextBox
            TextBox = Me.Controls.Item(sender.name)
            If Len(TextBox.Text) > 3 Then
                customToolTip.SetToolTip(TextBox, AddDate(TextBox.Text))
            End If
        Catch ex As Exception
            Dim st As New StackTrace(True)
            st = New StackTrace(ex, True)
            GlobalVariables.Linenumber = st.GetFrame(0).GetFileLineNumber().ToString()
            Dim f As New StackFrame
            FormMain.ErrorHandler(ex, System.Reflection.MethodBase.GetCurrentMethod().Name, GlobalVariables.Linenumber, f)
        End Try
    End Sub

    Private Sub TextBox25_MouseEnter(sender As Object, e As EventArgs) Handles TextBox25.MouseEnter
        Try
            Dim TextBox As TextBox
            TextBox = Me.Controls.Item(sender.name)
            If Len(TextBox.Text) > 3 Then
                customToolTip.SetToolTip(TextBox, AddDate(TextBox.Text))
            End If
        Catch ex As Exception
            Dim st As New StackTrace(True)
            st = New StackTrace(ex, True)
            GlobalVariables.Linenumber = st.GetFrame(0).GetFileLineNumber().ToString()
            Dim f As New StackFrame
            FormMain.ErrorHandler(ex, System.Reflection.MethodBase.GetCurrentMethod().Name, GlobalVariables.Linenumber, f)
        End Try
    End Sub

    Private Sub TextBox26_MouseEnter(sender As Object, e As EventArgs) Handles TextBox26.MouseEnter
        Try
            Dim TextBox As TextBox
            TextBox = Me.Controls.Item(sender.name)
            If Len(TextBox.Text) > 3 Then
                customToolTip.SetToolTip(TextBox, AddDate(TextBox.Text))
            End If
        Catch ex As Exception
            Dim st As New StackTrace(True)
            st = New StackTrace(ex, True)
            GlobalVariables.Linenumber = st.GetFrame(0).GetFileLineNumber().ToString()
            Dim f As New StackFrame
            FormMain.ErrorHandler(ex, System.Reflection.MethodBase.GetCurrentMethod().Name, GlobalVariables.Linenumber, f)
        End Try
    End Sub

    Private Sub TextBox27_MouseEnter(sender As Object, e As EventArgs) Handles TextBox27.MouseEnter
        Try
            Dim TextBox As TextBox
            TextBox = Me.Controls.Item(sender.name)
            If Len(TextBox.Text) > 3 Then
                customToolTip.SetToolTip(TextBox, AddDate(TextBox.Text))
            End If
        Catch ex As Exception
            Dim st As New StackTrace(True)
            st = New StackTrace(ex, True)
            GlobalVariables.Linenumber = st.GetFrame(0).GetFileLineNumber().ToString()
            Dim f As New StackFrame
            FormMain.ErrorHandler(ex, System.Reflection.MethodBase.GetCurrentMethod().Name, GlobalVariables.Linenumber, f)
        End Try
    End Sub

    Private Sub TextBox28_MouseEnter(sender As Object, e As EventArgs) Handles TextBox28.MouseEnter
        Try
            Dim TextBox As TextBox
            TextBox = Me.Controls.Item(sender.name)
            If Len(TextBox.Text) > 3 Then
                customToolTip.SetToolTip(TextBox, AddDate(TextBox.Text))
            End If
        Catch ex As Exception
            Dim st As New StackTrace(True)
            st = New StackTrace(ex, True)
            GlobalVariables.Linenumber = st.GetFrame(0).GetFileLineNumber().ToString()
            Dim f As New StackFrame
            FormMain.ErrorHandler(ex, System.Reflection.MethodBase.GetCurrentMethod().Name, GlobalVariables.Linenumber, f)
        End Try
    End Sub

    Private Sub TextBox29_MouseEnter(sender As Object, e As EventArgs) Handles TextBox29.MouseEnter
        Try
            Dim TextBox As TextBox
            TextBox = Me.Controls.Item(sender.name)
            If Len(TextBox.Text) > 3 Then
                customToolTip.SetToolTip(TextBox, AddDate(TextBox.Text))
            End If
        Catch ex As Exception
            Dim st As New StackTrace(True)
            st = New StackTrace(ex, True)
            GlobalVariables.Linenumber = st.GetFrame(0).GetFileLineNumber().ToString()
            Dim f As New StackFrame
            FormMain.ErrorHandler(ex, System.Reflection.MethodBase.GetCurrentMethod().Name, GlobalVariables.Linenumber, f)
        End Try
    End Sub

    Private Sub TextBox30_MouseEnter(sender As Object, e As EventArgs) Handles TextBox30.MouseEnter
        Try
            Dim TextBox As TextBox
            TextBox = Me.Controls.Item(sender.name)
            If Len(TextBox.Text) > 3 Then
                customToolTip.SetToolTip(TextBox, AddDate(TextBox.Text))
            End If
        Catch ex As Exception
            Dim st As New StackTrace(True)
            st = New StackTrace(ex, True)
            GlobalVariables.Linenumber = st.GetFrame(0).GetFileLineNumber().ToString()
            Dim f As New StackFrame
            FormMain.ErrorHandler(ex, System.Reflection.MethodBase.GetCurrentMethod().Name, GlobalVariables.Linenumber, f)
        End Try
    End Sub

    Private Sub TextBox31_MouseEnter(sender As Object, e As EventArgs) Handles TextBox31.MouseEnter
        Try
            Dim TextBox As TextBox
            TextBox = Me.Controls.Item(sender.name)
            If Len(TextBox.Text) > 3 Then
                customToolTip.SetToolTip(TextBox, AddDate(TextBox.Text))
            End If
        Catch ex As Exception
            Dim st As New StackTrace(True)
            st = New StackTrace(ex, True)
            GlobalVariables.Linenumber = st.GetFrame(0).GetFileLineNumber().ToString()
            Dim f As New StackFrame
            FormMain.ErrorHandler(ex, System.Reflection.MethodBase.GetCurrentMethod().Name, GlobalVariables.Linenumber, f)
        End Try
    End Sub

    Private Sub TextBox32_MouseEnter(sender As Object, e As EventArgs) Handles TextBox32.MouseEnter
        Try
            Dim TextBox As TextBox
            TextBox = Me.Controls.Item(sender.name)
            If Len(TextBox.Text) > 3 Then
                customToolTip.SetToolTip(TextBox, AddDate(TextBox.Text))
            End If
        Catch ex As Exception
            Dim st As New StackTrace(True)
            st = New StackTrace(ex, True)
            GlobalVariables.Linenumber = st.GetFrame(0).GetFileLineNumber().ToString()
            Dim f As New StackFrame
            FormMain.ErrorHandler(ex, System.Reflection.MethodBase.GetCurrentMethod().Name, GlobalVariables.Linenumber, f)
        End Try
    End Sub

    Private Sub TextBox33_MouseEnter(sender As Object, e As EventArgs) Handles TextBox33.MouseEnter
        Try
            Dim TextBox As TextBox
            TextBox = Me.Controls.Item(sender.name)
            If Len(TextBox.Text) > 3 Then
                customToolTip.SetToolTip(TextBox, AddDate(TextBox.Text))
            End If
        Catch ex As Exception
            Dim st As New StackTrace(True)
            st = New StackTrace(ex, True)
            GlobalVariables.Linenumber = st.GetFrame(0).GetFileLineNumber().ToString()
            Dim f As New StackFrame
            FormMain.ErrorHandler(ex, System.Reflection.MethodBase.GetCurrentMethod().Name, GlobalVariables.Linenumber, f)
        End Try
    End Sub

    Private Sub TextBox34_MouseEnter(sender As Object, e As EventArgs) Handles TextBox34.MouseEnter
        Try
            Dim TextBox As TextBox
            TextBox = Me.Controls.Item(sender.name)
            If Len(TextBox.Text) > 3 Then
                customToolTip.SetToolTip(TextBox, AddDate(TextBox.Text))
            End If
        Catch ex As Exception
            Dim st As New StackTrace(True)
            st = New StackTrace(ex, True)
            GlobalVariables.Linenumber = st.GetFrame(0).GetFileLineNumber().ToString()
            Dim f As New StackFrame
            FormMain.ErrorHandler(ex, System.Reflection.MethodBase.GetCurrentMethod().Name, GlobalVariables.Linenumber, f)
        End Try
    End Sub

    Private Sub TextBox35_MouseEnter(sender As Object, e As EventArgs) Handles TextBox35.MouseEnter
        Try
            Dim TextBox As TextBox
            TextBox = Me.Controls.Item(sender.name)
            If Len(TextBox.Text) > 3 Then
                customToolTip.SetToolTip(TextBox, AddDate(TextBox.Text))
            End If
        Catch ex As Exception
            Dim st As New StackTrace(True)
            st = New StackTrace(ex, True)
            GlobalVariables.Linenumber = st.GetFrame(0).GetFileLineNumber().ToString()
            Dim f As New StackFrame
            FormMain.ErrorHandler(ex, System.Reflection.MethodBase.GetCurrentMethod().Name, GlobalVariables.Linenumber, f)
        End Try
    End Sub

    Private Sub TextBox36_MouseEnter(sender As Object, e As EventArgs) Handles TextBox36.MouseEnter
        Try
            Dim TextBox As TextBox
            TextBox = Me.Controls.Item(sender.name)
            If Len(TextBox.Text) > 3 Then
                customToolTip.SetToolTip(TextBox, AddDate(TextBox.Text))
            End If
        Catch ex As Exception
            Dim st As New StackTrace(True)
            st = New StackTrace(ex, True)
            GlobalVariables.Linenumber = st.GetFrame(0).GetFileLineNumber().ToString()
            Dim f As New StackFrame
            FormMain.ErrorHandler(ex, System.Reflection.MethodBase.GetCurrentMethod().Name, GlobalVariables.Linenumber, f)
        End Try
    End Sub

    Private Sub TextBox37_MouseEnter(sender As Object, e As EventArgs) Handles TextBox37.MouseEnter
        Try
            Dim TextBox As TextBox
            TextBox = Me.Controls.Item(sender.name)
            If Len(TextBox.Text) > 3 Then
                customToolTip.SetToolTip(TextBox, AddDate(TextBox.Text))
            End If
        Catch ex As Exception
            Dim st As New StackTrace(True)
            st = New StackTrace(ex, True)
            GlobalVariables.Linenumber = st.GetFrame(0).GetFileLineNumber().ToString()
            Dim f As New StackFrame
            FormMain.ErrorHandler(ex, System.Reflection.MethodBase.GetCurrentMethod().Name, GlobalVariables.Linenumber, f)
        End Try
    End Sub

    Private Sub TextBox38_MouseEnter(sender As Object, e As EventArgs) Handles TextBox38.MouseEnter
        Try
            Dim TextBox As TextBox
            TextBox = Me.Controls.Item(sender.name)
            If Len(TextBox.Text) > 3 Then
                customToolTip.SetToolTip(TextBox, AddDate(TextBox.Text))
            End If
        Catch ex As Exception
            Dim st As New StackTrace(True)
            st = New StackTrace(ex, True)
            GlobalVariables.Linenumber = st.GetFrame(0).GetFileLineNumber().ToString()
            Dim f As New StackFrame
            FormMain.ErrorHandler(ex, System.Reflection.MethodBase.GetCurrentMethod().Name, GlobalVariables.Linenumber, f)
        End Try
    End Sub

    Private Sub TextBox39_MouseEnter(sender As Object, e As EventArgs) Handles TextBox39.MouseEnter
        Try
            Dim TextBox As TextBox
            TextBox = Me.Controls.Item(sender.name)
            If Len(TextBox.Text) > 3 Then
                customToolTip.SetToolTip(TextBox, AddDate(TextBox.Text))
            End If
        Catch ex As Exception
            Dim st As New StackTrace(True)
            st = New StackTrace(ex, True)
            GlobalVariables.Linenumber = st.GetFrame(0).GetFileLineNumber().ToString()
            Dim f As New StackFrame
            FormMain.ErrorHandler(ex, System.Reflection.MethodBase.GetCurrentMethod().Name, GlobalVariables.Linenumber, f)
        End Try
    End Sub

    Private Sub TextBox40_MouseEnter(sender As Object, e As EventArgs) Handles TextBox40.MouseEnter
        Try
            Dim TextBox As TextBox
            TextBox = Me.Controls.Item(sender.name)
            If Len(TextBox.Text) > 3 Then
                customToolTip.SetToolTip(TextBox, AddDate(TextBox.Text))
            End If
        Catch ex As Exception
            Dim st As New StackTrace(True)
            st = New StackTrace(ex, True)
            GlobalVariables.Linenumber = st.GetFrame(0).GetFileLineNumber().ToString()
            Dim f As New StackFrame
            FormMain.ErrorHandler(ex, System.Reflection.MethodBase.GetCurrentMethod().Name, GlobalVariables.Linenumber, f)
        End Try

    End Sub

    Private Sub TextBox41_MouseEnter(sender As Object, e As EventArgs) Handles TextBox41.MouseEnter
        Try
            Dim TextBox As TextBox
            TextBox = Me.Controls.Item(sender.name)
            If Len(TextBox.Text) > 3 Then
                customToolTip.SetToolTip(TextBox, AddDate(TextBox.Text))
            End If
        Catch ex As Exception
            Dim st As New StackTrace(True)
            st = New StackTrace(ex, True)
            GlobalVariables.Linenumber = st.GetFrame(0).GetFileLineNumber().ToString()
            Dim f As New StackFrame
            FormMain.ErrorHandler(ex, System.Reflection.MethodBase.GetCurrentMethod().Name, GlobalVariables.Linenumber, f)
        End Try
    End Sub

    Private Sub TextBox42_MouseEnter(sender As Object, e As EventArgs) Handles TextBox42.MouseEnter
        Try
            Dim TextBox As TextBox
            TextBox = Me.Controls.Item(sender.name)
            If Len(TextBox.Text) > 3 Then
                customToolTip.SetToolTip(TextBox, AddDate(TextBox.Text))
            End If
        Catch ex As Exception
            Dim st As New StackTrace(True)
            st = New StackTrace(ex, True)
            GlobalVariables.Linenumber = st.GetFrame(0).GetFileLineNumber().ToString()
            Dim f As New StackFrame
            FormMain.ErrorHandler(ex, System.Reflection.MethodBase.GetCurrentMethod().Name, GlobalVariables.Linenumber, f)
        End Try
    End Sub

    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click
        Dim theDate As New Date
        theDate = Now
        strYear = theDate.Year.ToString
        iMonth = theDate.Month
        strMonth = Now.ToString("MMMM")
        FillTheCalander()
    End Sub



    Private Sub ToolStripButton2_Click(sender As Object, e As EventArgs) Handles ToolStripButton2.Click
        Close()
    End Sub

    Private Sub ToolStripButton1_Click(sender As Object, e As EventArgs) Handles ToolStripButton1.Click
        FormAddReminder.ShowDialog()
        FillDateGridView()
        FillTheCalander()
    End Sub

    Private Sub ToolStripButton3_Click(sender As Object, e As EventArgs) Handles ToolStripButton3.Click
        Dim theDate As New Date
        theDate = Now
        strYear = theDate.Year.ToString
        iMonth = theDate.Month
        strMonth = Now.ToString("MMMM")
        FillTheCalander()
    End Sub


End Class