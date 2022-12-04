Imports System.Globalization


Public Class Form3



    Private Sub Form3_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            CenterToScreen()
            Me.Icon = My.Resources.Form
            ClearForm()
            TextBox1.Text = GlobalVariables.AlpacaName
            FormMain.ConnectedDB = New DataBaseFunctions()              'Open the database connection
            GlobalVariables.ds = FormMain.ConnectedDB.PopulateDataSet()        'Put the datainto a dataset
            For Each GroupRow As DataRow In GlobalVariables.ds.Tables("TestGroup").Rows
                ComboBox2.Items.Add(GroupRow.Item(1))
            Next
            ComboBox2.Text = ""
            ComboBox2.Text = GlobalVariables.Clickgroup.ToString
        Catch
        End Try
    End Sub


    Private Sub ToolStrip1_MouseHover(sender As Object, e As EventArgs) Handles ToolStrip1.MouseHover
        Me.Cursor = Cursors.Hand '
    End Sub
    Private Sub ToolStrip1_MouseLeave(sender As Object, e As EventArgs) Handles ToolStrip1.MouseLeave
        Me.Cursor = Cursors.Arrow
    End Sub





    Private Sub UpdateForm()
        'Label69.Text = NumericUpDown1.Value + NumericUpDown2.Value + NumericUpDown3.Value + NumericUpDown4.Value + NumericUpDown5.Value + NumericUpDown6.Value + NumericUpDown7.Value + NumericUpDown8.Value + NumericUpDown9.Value + NumericUpDown10.Value + NumericUpDown11.Value + NumericUpDown12.Value
        'Label5.Text = NumericUpDown13.Value + NumericUpDown14.Value + NumericUpDown15.Value + NumericUpDown16.Value + NumericUpDown17.Value + NumericUpDown18.Value + NumericUpDown19.Value + NumericUpDown20.Value + NumericUpDown21.Value + NumericUpDown22.Value + NumericUpDown23.Value + NumericUpDown24.Value
        EPG1.Text = (NumericUpDown1.Value)
        EPG2.Text = (NumericUpDown2.Value)
        EPG3.Text = (NumericUpDown3.Value)
        EPG4.Text = (NumericUpDown4.Value)
        EPG5.Text = (NumericUpDown5.Value)
        EPG6.Text = (NumericUpDown6.Value)
        Label67.Text = (Convert.ToInt32(EPG1.Text) + Convert.ToInt32(EPG2.Text) + Convert.ToInt32(EPG3.Text) + Convert.ToInt32(EPG4.Text) + Convert.ToInt32(EPG5.Text) + Convert.ToInt32(EPG6.Text)) * 5

        OPG1.Text = (NumericUpDown7.Value)
        OPG2.Text = (NumericUpDown8.Value)
        OPG3.Text = (NumericUpDown9.Value)
        OPG4.Text = (NumericUpDown10.Value)
        OPG5.Text = (NumericUpDown12.Value)
        OPG6.Text = (NumericUpDown11.Value)
        Label66.Text = (Convert.ToInt32(OPG1.Text) + Convert.ToInt32(OPG2.Text) + Convert.ToInt32(OPG3.Text) + Convert.ToInt32(OPG4.Text) + Convert.ToInt32(OPG5.Text) + Convert.ToInt32(OPG6.Text)) * 5

    End Sub

    Private Sub ClearForm()
        NumericUpDown1.Value = 0
        NumericUpDown2.Value = 0
        NumericUpDown3.Value = 0
        NumericUpDown4.Value = 0
        NumericUpDown5.Value = 0
        NumericUpDown6.Value = 0
        NumericUpDown7.Value = 0
        NumericUpDown8.Value = 0
        NumericUpDown9.Value = 0
        NumericUpDown10.Value = 0
        NumericUpDown11.Value = 0
        NumericUpDown12.Value = 0
        UpdateForm()
    End Sub




    Private Sub UpdateDb()

        'Update the testresult database
        FormMain.ConnectedDB = New DataBaseFunctions()              'Open the database connection
        GlobalVariables.ds = FormMain.ConnectedDB.PopulateDataSet()        'Put the datainto a dataset
        Dim drow As DataRow = GlobalVariables.ds.Tables("TestResults").NewRow

        drow.Item(1) = DateTime.ParseExact(DateTimePicker1.Value.ToShortDateString, "dd/MM/yyyy", CultureInfo.InvariantCulture)
        drow.Item(2) = TextBox1.Text
        drow.Item(3) = (ComboBox1.SelectedIndex + 1).ToString
        drow.Item(4) = "0"
        drow.Item(5) = GlobalVariables.TestGroupName
        drow.Item(6) = Label67.Text
        drow.Item(7) = Label66.Text
        drow.Item(8) = NumericUpDown1.Value.ToString("00.##") & "00"
        drow.Item(9) = NumericUpDown2.Value.ToString("00.##") & "00"
        drow.Item(10) = NumericUpDown3.Value.ToString("00.##") & "00"
        drow.Item(11) = NumericUpDown4.Value.ToString("00.##") & "00"
        drow.Item(12) = NumericUpDown5.Value.ToString("00.##") & "00"
        drow.Item(13) = NumericUpDown6.Value.ToString("00.##") & "00"
        drow.Item(14) = NumericUpDown7.Value.ToString("00.##") & "00"
        drow.Item(15) = NumericUpDown8.Value.ToString("00.##") & "00"
        drow.Item(16) = NumericUpDown9.Value.ToString("00.##") & "00"
        drow.Item(17) = NumericUpDown10.Value.ToString("00.##") & "00"
        drow.Item(18) = NumericUpDown11.Value.ToString("00.##") & "00"
        drow.Item(19) = NumericUpDown12.Value.ToString("00.##") & "00"
        drow.Item(20) = "No Treatment"
        drow.Item(21) = "Stoll"
        drow.Item(24) = CInt(Label67.Text) + CInt(Label66.Text)

        Dim dtUK As Date = DateTime.ParseExact(DateTimePicker1.Value.ToShortDateString, "dd/MM/yyyy", CultureInfo.InvariantCulture)
        Dim sql As String = "INSERT INTO TestResults (TestDate,TestName,TestFaecalGrade,TestNumber,TestNumberName,TestEPGTotal,TestOPGTotal,TestTrichostrongyles,TestTrichurius,TestNematordirus,TestCapillarid,TestMoniezid,TestEPGUnidentifed,TestEmac,TestEivitaesis,TestEalpacae,TestElamae,TestEpunoensis,TestOPGUnidentifed,TestTreat,Name) VALUES (#" & dtUK & "#,'" & TextBox1.Text & "','" & (ComboBox1.SelectedIndex + 1).ToString & "','" & "0" & "','" & GlobalVariables.TestGroupName & "','" & Label67.Text & "','" & Label66.Text & "','" & drow.Item(8) & "','" & drow.Item(9) & "','" & drow.Item(10) & "','" & drow.Item(11) & "','" & drow.Item(12) & "','" & drow.Item(13) & "','" & drow.Item(14) & "','" & drow.Item(15) & "','" & drow.Item(16) & "','" & drow.Item(17) & "','" & drow.Item(18) & "','" & drow.Item(19) & "','" & drow.Item(20) & "','" & drow.Item(21) & "')"
        FormMain.ConnectedDB.UpdateDatabase(sql)
        FormMain.Actionlog("New test result added")
    End Sub


    Private Sub PictureBox1_MouseHover(ByVal sender As Object, ByVal e As System.EventArgs) Handles PictureBox1.MouseHover

    End Sub

    Private Sub PictureBox1_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles PictureBox1.MouseLeave

    End Sub

    Private Sub PictureBox2_MouseHover(ByVal sender As Object, ByVal e As System.EventArgs) Handles PictureBox2.MouseHover

    End Sub

    Private Sub PictureBox2_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles PictureBox2.MouseLeave

    End Sub

    Private Sub PictureBox3_MouseHover(ByVal sender As Object, ByVal e As System.EventArgs) Handles PictureBox3.MouseHover

    End Sub

    Private Sub PictureBox3_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles PictureBox3.MouseLeave

    End Sub

    Private Sub PictureBox4_MouseHover(ByVal sender As Object, ByVal e As System.EventArgs) Handles PictureBox4.MouseHover

    End Sub

    Private Sub PictureBox4_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles PictureBox4.MouseLeave

    End Sub

    Private Sub PictureBox5_MouseHover(ByVal sender As Object, ByVal e As System.EventArgs) Handles PictureBox5.MouseHover

    End Sub

    Private Sub PictureBox5_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles PictureBox5.MouseLeave

    End Sub

    Private Sub PictureBox6_MouseHover(ByVal sender As Object, ByVal e As System.EventArgs) Handles PictureBox6.MouseHover

    End Sub

    Private Sub PictureBox6_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles PictureBox6.MouseLeave

    End Sub

    Private Sub PictureBox7_MouseHover(ByVal sender As Object, ByVal e As System.EventArgs) Handles PictureBox7.MouseHover

    End Sub

    Private Sub PictureBox7_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles PictureBox7.MouseLeave

    End Sub

    Private Sub PictureBox8_MouseHover(ByVal sender As Object, ByVal e As System.EventArgs) Handles PictureBox8.MouseHover

    End Sub

    Private Sub PictureBox8_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles PictureBox8.MouseLeave

    End Sub

    Private Sub PictureBox9_MouseHover(ByVal sender As Object, ByVal e As System.EventArgs) Handles PictureBox9.MouseHover

    End Sub

    Private Sub PictureBox9_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles PictureBox9.MouseLeave

    End Sub

    Private Sub PictureBox10_MouseHover(ByVal sender As Object, ByVal e As System.EventArgs) Handles PictureBox10.MouseHover

    End Sub

    Private Sub PictureBox10_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles PictureBox10.MouseLeave

    End Sub




    Private Sub ComboBox2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox2.SelectedIndexChanged

    End Sub


    Dim nudWheel As Boolean = False
    Dim lv As Decimal
    Dim fv As Decimal
    Dim direction As Boolean = False 'false = decrement, true = increment
    Dim wheelCT As Integer = 0


    Private Sub NumericUpDown1_MouseWheel(sender As Object, e As MouseEventArgs) Handles NumericUpDown1.MouseWheel
        nudWheel = True
        wheelCT = 0
        lv = NumericUpDown1.Value 'before ValueChanged event
    End Sub



    Private Sub NumericUpDown1_ValueChanged(sender As Object, e As EventArgs) Handles NumericUpDown1.ValueChanged
        Static recur As Boolean = False
        If recur Then Exit Sub
        If nudWheel Then
            If wheelCT = 0 Then
                fv = NumericUpDown1.Value
                If fv > lv Then
                    direction = True
                Else
                    direction = False
                End If
            End If
            wheelCT += 1
            recur = True
            If direction Then
                NumericUpDown1.Value -= NumericUpDown1.Increment
            Else
                NumericUpDown1.Value += NumericUpDown1.Increment
            End If
            recur = False
            If wheelCT = SystemInformation.MouseWheelScrollLines Then
                recur = True
                NumericUpDown1.Value = fv
                recur = False
                nudWheel = False
            End If
        End If
        UpdateForm()
    End Sub


    Private Sub NumericUpDown2_MouseWheel(sender As Object, e As MouseEventArgs) Handles NumericUpDown2.MouseWheel
        nudWheel = True
        wheelCT = 0
        lv = NumericUpDown2.Value 'before ValueChanged event
    End Sub
    Private Sub NumericUpDown2_ValueChanged(sender As Object, e As EventArgs) Handles NumericUpDown2.ValueChanged


        Static recur As Boolean = False
        If recur Then Exit Sub
        If nudWheel Then
            If wheelCT = 0 Then
                fv = NumericUpDown2.Value
                If fv > lv Then
                    direction = True
                Else
                    direction = False
                End If
            End If
            wheelCT += 1
            recur = True
            If direction Then
                NumericUpDown2.Value -= NumericUpDown2.Increment
            Else
                NumericUpDown2.Value += NumericUpDown2.Increment
            End If
            recur = False
            If wheelCT = SystemInformation.MouseWheelScrollLines Then
                recur = True
                NumericUpDown2.Value = fv
                recur = False
                nudWheel = False
            End If
        End If
        UpdateForm()
    End Sub

    Private Sub NumericUpDown3_MouseWheel(sender As Object, e As MouseEventArgs) Handles NumericUpDown3.MouseWheel
        nudWheel = True
        wheelCT = 0
        lv = NumericUpDown3.Value 'before ValueChanged event
    End Sub
    Private Sub NumericUpDown3_ValueChanged(sender As Object, e As EventArgs) Handles NumericUpDown3.ValueChanged


        Static recur As Boolean = False
        If recur Then Exit Sub
        If nudWheel Then
            If wheelCT = 0 Then
                fv = NumericUpDown3.Value
                If fv > lv Then
                    direction = True
                Else
                    direction = False
                End If
            End If
            wheelCT += 1
            recur = True
            If direction Then
                NumericUpDown3.Value -= NumericUpDown3.Increment
            Else
                NumericUpDown3.Value += NumericUpDown3.Increment
            End If
            recur = False
            If wheelCT = SystemInformation.MouseWheelScrollLines Then
                recur = True
                NumericUpDown3.Value = fv
                recur = False
                nudWheel = False
            End If
        End If
        UpdateForm()
    End Sub

    Private Sub NumericUpDown4_MouseWheel(sender As Object, e As MouseEventArgs) Handles NumericUpDown4.MouseWheel
        nudWheel = True
        wheelCT = 0
        lv = NumericUpDown4.Value 'before ValueChanged event
    End Sub
    Private Sub NumericUpDown4_ValueChanged(sender As Object, e As EventArgs) Handles NumericUpDown4.ValueChanged


        Static recur As Boolean = False
        If recur Then Exit Sub
        If nudWheel Then
            If wheelCT = 0 Then
                fv = NumericUpDown4.Value
                If fv > lv Then
                    direction = True
                Else
                    direction = False
                End If
            End If
            wheelCT += 1
            recur = True
            If direction Then
                NumericUpDown4.Value -= NumericUpDown4.Increment
            Else
                NumericUpDown4.Value += NumericUpDown4.Increment
            End If
            recur = False
            If wheelCT = SystemInformation.MouseWheelScrollLines Then
                recur = True
                NumericUpDown4.Value = fv
                recur = False
                nudWheel = False
            End If
        End If
        UpdateForm()
    End Sub

    Private Sub NumericUpDown5_MouseWheel(sender As Object, e As MouseEventArgs) Handles NumericUpDown5.MouseWheel
        nudWheel = True
        wheelCT = 0
        lv = NumericUpDown5.Value 'before ValueChanged event
    End Sub
    Private Sub NumericUpDown5_ValueChanged(sender As Object, e As EventArgs) Handles NumericUpDown5.ValueChanged


        Static recur As Boolean = False
        If recur Then Exit Sub
        If nudWheel Then
            If wheelCT = 0 Then
                fv = NumericUpDown5.Value
                If fv > lv Then
                    direction = True
                Else
                    direction = False
                End If
            End If
            wheelCT += 1
            recur = True
            If direction Then
                NumericUpDown5.Value -= NumericUpDown5.Increment
            Else
                NumericUpDown5.Value += NumericUpDown5.Increment
            End If
            recur = False
            If wheelCT = SystemInformation.MouseWheelScrollLines Then
                recur = True
                NumericUpDown5.Value = fv
                recur = False
                nudWheel = False
            End If
        End If
        UpdateForm()
    End Sub

    Private Sub NumericUpDown6_MouseWheel(sender As Object, e As MouseEventArgs) Handles NumericUpDown6.MouseWheel
        nudWheel = True
        wheelCT = 0
        lv = NumericUpDown6.Value 'before ValueChanged event
    End Sub
    Private Sub NumericUpDown6_ValueChanged(sender As Object, e As EventArgs) Handles NumericUpDown6.ValueChanged


        Static recur As Boolean = False
        If recur Then Exit Sub
        If nudWheel Then
            If wheelCT = 0 Then
                fv = NumericUpDown6.Value
                If fv > lv Then
                    direction = True
                Else
                    direction = False
                End If
            End If
            wheelCT += 1
            recur = True
            If direction Then
                NumericUpDown6.Value -= NumericUpDown6.Increment
            Else
                NumericUpDown6.Value += NumericUpDown6.Increment
            End If
            recur = False
            If wheelCT = SystemInformation.MouseWheelScrollLines Then
                recur = True
                NumericUpDown6.Value = fv
                recur = False
                nudWheel = False
            End If
        End If
        UpdateForm()
    End Sub

    Private Sub NumericUpDown7_MouseWheel(sender As Object, e As MouseEventArgs) Handles NumericUpDown7.MouseWheel
        nudWheel = True
        wheelCT = 0
        lv = NumericUpDown7.Value 'before ValueChanged event
    End Sub
    Private Sub NumericUpDown7_ValueChanged(sender As Object, e As EventArgs) Handles NumericUpDown7.ValueChanged


        Static recur As Boolean = False
        If recur Then Exit Sub
        If nudWheel Then
            If wheelCT = 0 Then
                fv = NumericUpDown7.Value
                If fv > lv Then
                    direction = True
                Else
                    direction = False
                End If
            End If
            wheelCT += 1
            recur = True
            If direction Then
                NumericUpDown7.Value -= NumericUpDown7.Increment
            Else
                NumericUpDown7.Value += NumericUpDown7.Increment
            End If
            recur = False
            If wheelCT = SystemInformation.MouseWheelScrollLines Then
                recur = True
                NumericUpDown7.Value = fv
                recur = False
                nudWheel = False
            End If
        End If
        UpdateForm()
    End Sub

    Private Sub NumericUpDown8_MouseWheel(sender As Object, e As MouseEventArgs) Handles NumericUpDown8.MouseWheel
        nudWheel = True
        wheelCT = 0
        lv = NumericUpDown8.Value 'before ValueChanged event
    End Sub
    Private Sub NumericUpDown8_ValueChanged(sender As Object, e As EventArgs) Handles NumericUpDown8.ValueChanged


        Static recur As Boolean = False
        If recur Then Exit Sub
        If nudWheel Then
            If wheelCT = 0 Then
                fv = NumericUpDown8.Value
                If fv > lv Then
                    direction = True
                Else
                    direction = False
                End If
            End If
            wheelCT += 1
            recur = True
            If direction Then
                NumericUpDown8.Value -= NumericUpDown8.Increment
            Else
                NumericUpDown8.Value += NumericUpDown8.Increment
            End If
            recur = False
            If wheelCT = SystemInformation.MouseWheelScrollLines Then
                recur = True
                NumericUpDown8.Value = fv
                recur = False
                nudWheel = False
            End If
        End If
        UpdateForm()
    End Sub

    Private Sub NumericUpDown9_MouseWheel(sender As Object, e As MouseEventArgs) Handles NumericUpDown9.MouseWheel
        nudWheel = True
        wheelCT = 0
        lv = NumericUpDown9.Value 'before ValueChanged event
    End Sub
    Private Sub NumericUpDown9_ValueChanged(sender As Object, e As EventArgs) Handles NumericUpDown9.ValueChanged


        Static recur As Boolean = False
        If recur Then Exit Sub
        If nudWheel Then
            If wheelCT = 0 Then
                fv = NumericUpDown9.Value
                If fv > lv Then
                    direction = True
                Else
                    direction = False
                End If
            End If
            wheelCT += 1
            recur = True
            If direction Then
                NumericUpDown9.Value -= NumericUpDown9.Increment
            Else
                NumericUpDown9.Value += NumericUpDown9.Increment
            End If
            recur = False
            If wheelCT = SystemInformation.MouseWheelScrollLines Then
                recur = True
                NumericUpDown9.Value = fv
                recur = False
                nudWheel = False
            End If
        End If
        UpdateForm()
    End Sub

    Private Sub NumericUpDown10_MouseWheel(sender As Object, e As MouseEventArgs) Handles NumericUpDown10.MouseWheel
        nudWheel = True
        wheelCT = 0
        lv = NumericUpDown10.Value 'before ValueChanged event
    End Sub
    Private Sub NumericUpDown10_ValueChanged(sender As Object, e As EventArgs) Handles NumericUpDown10.ValueChanged


        Static recur As Boolean = False
        If recur Then Exit Sub
        If nudWheel Then
            If wheelCT = 0 Then
                fv = NumericUpDown10.Value
                If fv > lv Then
                    direction = True
                Else
                    direction = False
                End If
            End If
            wheelCT += 1
            recur = True
            If direction Then
                NumericUpDown10.Value -= NumericUpDown10.Increment
            Else
                NumericUpDown10.Value += NumericUpDown10.Increment
            End If
            recur = False
            If wheelCT = SystemInformation.MouseWheelScrollLines Then
                recur = True
                NumericUpDown10.Value = fv
                recur = False
                nudWheel = False
            End If
        End If
        UpdateForm()
    End Sub

    Private Sub NumericUpDown11_MouseWheel(sender As Object, e As MouseEventArgs) Handles NumericUpDown11.MouseWheel
        nudWheel = True
        wheelCT = 0
        lv = NumericUpDown11.Value 'before ValueChanged event
    End Sub
    Private Sub NumericUpDown11_ValueChanged(sender As Object, e As EventArgs) Handles NumericUpDown11.ValueChanged


        Static recur As Boolean = False
        If recur Then Exit Sub
        If nudWheel Then
            If wheelCT = 0 Then
                fv = NumericUpDown11.Value
                If fv > lv Then
                    direction = True
                Else
                    direction = False
                End If
            End If
            wheelCT += 1
            recur = True
            If direction Then
                NumericUpDown11.Value -= NumericUpDown11.Increment
            Else
                NumericUpDown11.Value += NumericUpDown11.Increment
            End If
            recur = False
            If wheelCT = SystemInformation.MouseWheelScrollLines Then
                recur = True
                NumericUpDown11.Value = fv
                recur = False
                nudWheel = False
            End If
        End If
        UpdateForm()
    End Sub


    Private Sub ToolStripButton2_Click(sender As Object, e As EventArgs) Handles ToolStripButton2.Click
        Me.Close()
    End Sub

    Private Sub ToolStripButton1_Click(sender As Object, e As EventArgs) Handles ToolStripButton1.Click

        Dim result As DialogResult

        result = MessageBox.Show("Confirm data is to be saved into the database?", "Update database", MessageBoxButtons.YesNo)

        If (result = DialogResult.Yes) Then

            If ComboBox1.SelectedIndex = -1 Then
                MessageBox.Show("You must choose a facial type!", "Warning", MessageBoxButtons.OK)
                Exit Sub
            End If
            UpdateDb()
            FormMain.loadFormWithAnimals()
            FormMain.Actionlog("New test result saved to database")
            Me.Close()
        Else
            FormMain.Actionlog("Exit without saving test result")
            Me.Close()
        End If


    End Sub

    Private Sub ToolStripButton3_Click(sender As Object, e As EventArgs)

    End Sub


End Class