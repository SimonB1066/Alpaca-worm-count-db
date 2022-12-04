Imports System.IO
Imports System.Globalization

Public Class FormAlpaca
    Dim enableEdits As Boolean = False
    Dim EnterAninalName As String
    Private Sub ToolStrip1_MouseHover(sender As Object, e As EventArgs) Handles ToolStrip1.MouseHover
        Me.Cursor = Cursors.Hand '
    End Sub
    Private Sub ToolStrip1_MouseLeave(sender As Object, e As EventArgs) Handles ToolStrip1.MouseLeave
        Me.Cursor = Cursors.Arrow
    End Sub
    Private Sub FormAlpaca_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        CenterToScreen()
        Me.Icon = My.Resources.Form
    End Sub
    Private Sub BuildReport()
        RichTextBox1.Clear()
        RichTextBox1.ZoomFactor = 1


        RichTextBox1.Text = ""

        Dim site As String = GlobalVariables.AlpacaName
        Dim SelectedPicture As String = GlobalVariables.DbDriveLocation & "Header.png"
        Dim Picture As Bitmap = New Bitmap(SelectedPicture)
        Clipboard.SetImage(FormMain.ResizeImage(Picture))
        Dim PictureFormat As DataFormats.Format = DataFormats.GetFormat(DataFormats.Bitmap)
        If RichTextBox1.CanPaste(PictureFormat) Then
            RichTextBox1.Paste(PictureFormat)
        End If


        RichTextBox1.SelectionFont = New Font("Tahoma", 20, FontStyle.Bold)
        RichTextBox1.SelectionColor = Color.Navy
        RichTextBox1.AppendText("    " & site & vbNewLine)
        RichTextBox1.SelectionFont = New Font("Courier New", 12, FontStyle.Regular)
        RichTextBox1.SelectionColor = Color.Black
        RichTextBox1.AppendText("Parasite and worm egg report" & vbNewLine & vbNewLine)
        RichTextBox1.SelectionFont = New Font("Courier New", 9.75, FontStyle.Regular)
        RichTextBox1.AppendText("Report date   " & DateTime.Today & vbNewLine)
        RichTextBox1.AppendText("____________________________________________________________" & vbNewLine)
        RichTextBox1.AppendText("Test date         Total EPG" & vbNewLine)
        RichTextBox1.AppendText(vbNewLine)


        Dim NumberOfAnimals As Int16 = GlobalVariables.ds.Tables("Alpaca").Rows.Count - 1
        Dim WorstAnimal As String = ""
        Dim Total As Int16 = 0
        Dim HighestCount As Int16 = 0
        Dim RateNumber As Int16 = 0
        Dim backcolour As Boolean = False
        Dim strEggs As String = ""
        Dim Eggs As Int16 = 0

        For j = 0 To GlobalVariables.ds.Tables("TestResults").Rows.Count - 1
            'Total = Convert.ToInt32(GlobalVariables.ds.Tables("TestResults").Rows(j)("TestEPGTotal")) + Convert.ToInt32(GlobalVariables.ds.Tables("TestResults").Rows(j)("TestOPGTotal"))
            Total = Convert.ToInt32(GlobalVariables.ds.Tables("TestResults").Rows(j)("TestEPGTotal")) + Convert.ToInt32(GlobalVariables.ds.Tables("TestResults").Rows(j)("TestOPGTotal"))
            Eggs = 0
            strEggs = ""



            If GlobalVariables.ds.Tables("TestResults").Rows(j)("TestName") = GlobalVariables.AlpacaName Then

                'Now build the string name animal first the the test name and last the result and type of egg found
                Dim str As String = ""

                str = str & GlobalVariables.ds.Tables("TestResults").Rows(j)("TestDate").ToString.Substring(0, 10)
                str = str.PadRight(20, " ")

                str = str & Total.ToString
                str = str.PadRight(30, " ")


                If Convert.ToInt16(GlobalVariables.ds.Tables("TestResults").Rows(j)("TestTrichostrongyles")) > 0 Then
                    Eggs = System.Convert.ToInt32(GlobalVariables.ds.Tables("TestResults").Rows(j)("TestTrichostrongyles").ToString.Substring(0, 2))
                    Eggs = Eggs + System.Convert.ToInt32(GlobalVariables.ds.Tables("TestResults").Rows(j)("TestTrichostrongyles").ToString.Substring(2, 2))
                    strEggs = AddEggString(str, strEggs, Eggs, "..Trichostrongyles  ")
                End If

                If Convert.ToInt16(GlobalVariables.ds.Tables("TestResults").Rows(j)("TestTrichurius")) > 0 Then
                    Eggs = System.Convert.ToInt32(GlobalVariables.ds.Tables("TestResults").Rows(j)("TestTrichurius").ToString.Substring(0, 2))
                    Eggs = Eggs + System.Convert.ToInt32(GlobalVariables.ds.Tables("TestResults").Rows(j)("TestTrichurius").ToString.Substring(2, 2))
                    strEggs = AddEggString(str, strEggs, Eggs, "..Trichurius  ")
                End If

                If Convert.ToInt16(GlobalVariables.ds.Tables("TestResults").Rows(j)("TestNematordirus")) > 0 Then
                    Eggs = System.Convert.ToInt32(GlobalVariables.ds.Tables("TestResults").Rows(j)("TestNematordirus").ToString.Substring(0, 2))
                    Eggs = Eggs + System.Convert.ToInt32(GlobalVariables.ds.Tables("TestResults").Rows(j)("TestNematordirus").ToString.Substring(2, 2))
                    strEggs = AddEggString(str, strEggs, Eggs, "..Nematordirus  ")
                End If

                If Convert.ToInt16(GlobalVariables.ds.Tables("TestResults").Rows(j)("TestCapillarid")) > 0 Then
                    Eggs = System.Convert.ToInt32(GlobalVariables.ds.Tables("TestResults").Rows(j)("TestCapillarid").ToString.Substring(0, 2))
                    Eggs = Eggs + System.Convert.ToInt32(GlobalVariables.ds.Tables("TestResults").Rows(j)("TestCapillarid").ToString.Substring(2, 2))
                    strEggs = AddEggString(str, strEggs, Eggs, "..Capillarid  ")
                End If

                If Convert.ToInt16(GlobalVariables.ds.Tables("TestResults").Rows(j)("TestMoniezid")) > 0 Then
                    Eggs = System.Convert.ToInt32(GlobalVariables.ds.Tables("TestResults").Rows(j)("TestMoniezid").ToString.Substring(0, 2))
                    Eggs = Eggs + System.Convert.ToInt32(GlobalVariables.ds.Tables("TestResults").Rows(j)("TestMoniezid").ToString.Substring(2, 2))
                    strEggs = AddEggString(str, strEggs, Eggs, "..Moniezid  ")
                End If

                If Convert.ToInt16(GlobalVariables.ds.Tables("TestResults").Rows(j)("TestEmac")) > 0 Then
                    Eggs = System.Convert.ToInt32(GlobalVariables.ds.Tables("TestResults").Rows(j)("TestEmac").ToString.Substring(0, 2))
                    Eggs = Eggs + System.Convert.ToInt32(GlobalVariables.ds.Tables("TestResults").Rows(j)("TestEmac").ToString.Substring(2, 2))
                    strEggs = AddEggString(str, strEggs, Eggs, "..E-mac  ")
                End If

                If Convert.ToInt16(GlobalVariables.ds.Tables("TestResults").Rows(j)("TestEivitaesis")) > 0 Then
                    Eggs = System.Convert.ToInt32(GlobalVariables.ds.Tables("TestResults").Rows(j)("TestEivitaesis").ToString.Substring(0, 2))
                    Eggs = Eggs + System.Convert.ToInt32(GlobalVariables.ds.Tables("TestResults").Rows(j)("TestEivitaesis").ToString.Substring(2, 2))
                    strEggs = AddEggString(str, strEggs, Eggs, "..E-ivitaesis  ")
                End If

                If Convert.ToInt16(GlobalVariables.ds.Tables("TestResults").Rows(j)("TestEalpacae")) > 0 Then
                    Eggs = System.Convert.ToInt32(GlobalVariables.ds.Tables("TestResults").Rows(j)("TestEalpacae").ToString.Substring(0, 2))
                    Eggs = Eggs + System.Convert.ToInt32(GlobalVariables.ds.Tables("TestResults").Rows(j)("TestEalpacae").ToString.Substring(2, 2))
                    strEggs = AddEggString(str, strEggs, Eggs, "..E-alpacae  ")
                End If

                If Convert.ToInt16(GlobalVariables.ds.Tables("TestResults").Rows(j)("TestElamae")) > 0 Then
                    Eggs = System.Convert.ToInt32(GlobalVariables.ds.Tables("TestResults").Rows(j)("TestElamae").ToString.Substring(0, 2))
                    Eggs = Eggs + System.Convert.ToInt32(GlobalVariables.ds.Tables("TestResults").Rows(j)("TestElamae").ToString.Substring(2, 2))
                    strEggs = AddEggString(str, strEggs, Eggs, "..E-lamae  ")
                End If

                If Convert.ToInt16(GlobalVariables.ds.Tables("TestResults").Rows(j)("TestEpunoensis")) > 0 Then
                    Eggs = System.Convert.ToInt32(GlobalVariables.ds.Tables("TestResults").Rows(j)("TestEpunoensis").ToString.Substring(0, 2))
                    Eggs = Eggs + System.Convert.ToInt32(GlobalVariables.ds.Tables("TestResults").Rows(j)("TestEpunoensis").ToString.Substring(2, 2))
                    strEggs = AddEggString(str, strEggs, Eggs, "..E-punoensis  ")
                End If


                str = str.PadRight(30, " ")
                str = str & strEggs
                str = str.PadRight(74, " ")

                Dim wwStr As String = WordWrap(str, 75)
                RichTextBox1.AppendText(wwStr)

                backcolour = Not backcolour
                If backcolour Then
                    RichTextBox1.SelectionBackColor = Color.LightSteelBlue
                Else
                    RichTextBox1.SelectionBackColor = Color.White
                End If

            End If
        Next


        RichTextBox1.SelectionFont = New Font("arial", 10, FontStyle.Italic)
        RichTextBox1.AppendText(vbNewLine & vbNewLine & vbNewLine)
        RichTextBox1.AppendText(vbNewLine & "*************** End of Report *****************")
    End Sub
    Private Sub CheckBox1_Click(sender As Object, e As EventArgs) Handles CheckBox1.Click


        Try

            FormMain.ConnectedDB.UpdateDatabase("UPDATE Alpaca SET OnFarm = '" & CheckBox1.Checked.ToString & "' WHERE Name = '" & GlobalVariables.AlpacaName & "'")
            Me.Refresh()
            Threading.Thread.Sleep(1000)
            FormMain.DataGridView1.Columns.Clear()
            FormMain.DataGridView1.DataSource = Nothing
            GlobalVariables.ds.Dispose()
            FormMain.loadFormWithAnimals()
            FormMain.DataGridViewRefresh()


        Catch ex As Exception
            Dim st As New StackTrace(True)
            st = New StackTrace(ex, True)
            GlobalVariables.Linenumber = st.GetFrame(0).GetFileLineNumber().ToString()
            Dim f As New StackFrame
            FormMain.ErrorHandler(ex, System.Reflection.MethodBase.GetCurrentMethod().Name, GlobalVariables.Linenumber, f)
        End Try
    End Sub
    Private Function AddEggString(str As String, strEggs As String, Eggs As Integer, name As String)
        Dim strLength As Integer = (Len(str) + Len(strEggs) + Len(name))
        strEggs = strEggs & Eggs.ToString & name

        AddEggString = strEggs
    End Function
    Function WordWrap(ByVal Text As String, Optional ByVal maxLengthOfALine As Integer = 77)


        Dim startingPosition As Integer = 0
        Dim endingPosition As Integer = 0
        Dim lineLength As Integer = maxLengthOfALine
        Dim line As String
        Dim SplitLine As String = ""

        Try
            While startingPosition < Text.Length
                endingPosition = Text.Length

                'If the string is less then the max length no need to split
                If endingPosition - startingPosition < maxLengthOfALine Then
                    If endingPosition = -1 Then
                        line = Text.Substring(startingPosition, Text.Length - startingPosition)
                        SplitLine = SplitLine & line.PadRight(maxLengthOfALine, " ") & vbCrLf
                        Exit While
                    Else
                        line = Text.Substring(startingPosition, endingPosition - startingPosition)
                        startingPosition = endingPosition + 1
                        SplitLine = SplitLine & line.PadRight(maxLengthOfALine, " ") & vbCrLf
                    End If
                Else
                    'String is greater than max length so need to split the line
                    While lineLength + startingPosition <= endingPosition
                        While Text.Substring(startingPosition + lineLength - 1, 1) <> Chr(32)
                            lineLength -= 1
                        End While
                        line = Text.Substring(startingPosition, lineLength)
                        If SplitLine = "" Then
                            SplitLine = SplitLine & line.PadRight(maxLengthOfALine, " ") & vbCrLf
                        Else
                            SplitLine = SplitLine & ("                              " & line).PadRight(maxLengthOfALine, " ") & vbCrLf
                        End If
                        If lineLength < maxLengthOfALine Then
                            startingPosition += lineLength
                        Else
                            startingPosition += lineLength + 1
                        End If
                        'lineLength = maxLengthOfALine
                        lineLength = maxLengthOfALine - 30
                    End While

                    line = Text.Substring(startingPosition, endingPosition - startingPosition)
                    If line <> "" Then
                        SplitLine = SplitLine & ("                              " & line).PadRight(maxLengthOfALine, " ") & vbCrLf
                        startingPosition = endingPosition + 1
                    End If
                End If
            End While
        Catch ex As Exception
            Dim st As New StackTrace(True)
            st = New StackTrace(ex, True)
            GlobalVariables.Linenumber = st.GetFrame(0).GetFileLineNumber().ToString()
            Dim f As New StackFrame
            FormMain.ErrorHandler(ex, System.Reflection.MethodBase.GetCurrentMethod().Name, GlobalVariables.Linenumber, f)
        End Try

        WordWrap = SplitLine
    End Function
    Private Sub FormAlpaca_Shown(sender As Object, e As EventArgs) Handles Me.Shown

        TextBox1.Text = GlobalVariables.AlpacaName
        EnterAninalName = GlobalVariables.AlpacaName


        For j = 0 To GlobalVariables.ds.Tables(0).Rows.Count - 1
            If GlobalVariables.ds.Tables(0).Rows(j)("Name") = GlobalVariables.AlpacaName Then
                If GlobalVariables.ds.Tables(0).Rows(j)("OnFarm") = True Then
                    CheckBox1.Checked = True
                Else
                    CheckBox1.Checked = False
                End If
                Exit For
            End If
        Next
        BuildReport()
        enableEdits = True
        ToolStripButton3.Enabled = True
        ToolStripButton1.Enabled = True
    End Sub
    Public Function GetColour(Colour As String) As String
        Try
            Dim dbProvider As String = "Provider=Microsoft.ACE.OLEDB.12.0;"
            Dim dbSource As String = "Data Source = " & GlobalVariables.PacaManpath
            Using OpenCon = New OleDb.OleDbConnection(dbProvider & dbSource)

                OpenCon.Open()

                Using da = New OleDb.OleDbDataAdapter("SELECT * FROM " & "tblReference_Colour_Code", OpenCon)
                    Dim builder As New OleDb.OleDbCommandBuilder
                    builder = New OleDb.OleDbCommandBuilder(da)
                    da.Fill(GlobalVariables.ds, "dsSColour")
                    For j = 0 To GlobalVariables.ds.Tables("dsSColour").Rows.Count - 1
                        If GlobalVariables.ds.Tables("dsSColour").Rows(j)("Colour_Code") = Colour Then
                            GetColour = GlobalVariables.ds.Tables("dsSColour").Rows(j)("Colour_Code_Name") & "(" & Colour & ")"
                            Exit Function
                        End If
                    Next
                End Using
            End Using
        Catch ex As Exception
        End Try
        GetColour = "Colour: Unknown"
    End Function
    Public Function GetSex(Number As Integer) As String
        Try
            Dim dbProvider As String = "Provider=Microsoft.ACE.OLEDB.12.0;"
            Dim dbSource As String = "Data Source = " & GlobalVariables.PacaManpath
            Using OpenCon = New OleDb.OleDbConnection(dbProvider & dbSource)

                OpenCon.Open()

                Using da = New OleDb.OleDbDataAdapter("SELECT * FROM " & "tblReference_Sex", OpenCon)
                    Dim builder As New OleDb.OleDbCommandBuilder
                    builder = New OleDb.OleDbCommandBuilder(da)
                    da.Fill(GlobalVariables.ds, "dsSex")
                    For j = 0 To GlobalVariables.ds.Tables("dsSex").Rows.Count - 1
                        If GlobalVariables.ds.Tables("dsSex").Rows(j)("Sex") = Number Then
                            GetSex = GlobalVariables.ds.Tables("dsSex").Rows(j)("Sex_Name")
                            Exit Function
                        End If
                    Next
                End Using
            End Using
        Catch ex As Exception
        End Try
        GetSex = "Sex: Unknown"
    End Function
    Private Sub TextBox1_Validated(sender As Object, e As EventArgs) Handles TextBox1.Validated
        Dim rtn As MsgBoxResult
        If enableEdits Then
            rtn = MessageBox.Show("You have made a change to the animals name" & vbNewLine & "Do you which to save the change, This will correct all the" & vbNewLine & "records allocated to the old name.", "Warning", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            If rtn = vbYes Then



                Try
                    FormMain.ConnectedDB.UpdateDatabase("UPDATE TestResults SET TestName = '" & TextBox1.Text & "' WHERE TestName = '" & GlobalVariables.AlpacaName & "'")
                    FormMain.ConnectedDB.UpdateDatabase("UPDATE Alpaca SET Name = '" & TextBox1.Text & "' WHERE Name = '" & GlobalVariables.AlpacaName & "'")

                    Me.Refresh()
                    Threading.Thread.Sleep(500)
                    FormMain.DataGridView1.Columns.Clear()
                    FormMain.DataGridView1.DataSource = Nothing
                    GlobalVariables.ds.Dispose()
                    FormMain.loadFormWithAnimals()
                    FormMain.DataGridViewRefresh()

                Catch ex As Exception
                    Dim st As New StackTrace(True)
                    st = New StackTrace(ex, True)
                    GlobalVariables.Linenumber = st.GetFrame(0).GetFileLineNumber().ToString()
                    Dim f As New StackFrame
                    FormMain.ErrorHandler(ex, System.Reflection.MethodBase.GetCurrentMethod().Name, GlobalVariables.Linenumber, f)
                End Try


                GlobalVariables.AlpacaName = TextBox1.Text
                BuildReport()
            Else
                TextBox1.Text = GlobalVariables.AlpacaName
            End If
        End If
    End Sub
    Private Sub ToolStripButton1_Click(sender As Object, e As EventArgs) Handles ToolStripButton1.Click
        Dim printDlg As New PrintDialog()
        'Dim Combo = Mid(Form1.ComboBox1.SelectedItem.ToString, 1, Len(Form1.ComboBox1.SelectedItem.ToString))
        ' Initialize the print dialog with the number of pages in the document.
        printDlg.AllowSomePages = True
        printDlg.PrinterSettings.MinimumPage = 1

        printDlg.PrinterSettings.FromPage = 1
        printDlg.Document = PrintDocument1
        Dim ReportName As String = GlobalVariables.DbDriveLocation & "Print.rtf"
        RichTextBox1.SaveFile(ReportName)

        ' Landscape()



        Dim psi As New ProcessStartInfo

        psi.UseShellExecute = True
        psi.Verb = "print"
        psi.Arguments = printDlg.PrinterSettings.PrinterName.ToString()
        psi.WindowStyle = ProcessWindowStyle.Hidden

        psi.FileName = ReportName ' Here specify a document to be printed

        Process.Start(psi)


        Close()
    End Sub
    Private Sub ToolStripButton3_Click(sender As Object, e As EventArgs) Handles ToolStripButton3.Click
        Dim ReportName As String = GlobalVariables.DbDriveLocation & "Report.rtf"
        Try
            ToolStripButton3.Enabled = False
            Label3.Visible = True


            Dim str1 As String = RichTextBox1.Rtf
            Using writer As New StreamWriter(ReportName, False)
                writer.WriteLine(str1)
                writer.Flush()
                writer.Close()
            End Using
            Label3.Visible = False
        Catch ex As Exception
            Dim st As New StackTrace(True)
            st = New StackTrace(ex, True)
            GlobalVariables.Linenumber = st.GetFrame(0).GetFileLineNumber().ToString()
            Dim f As New StackFrame
            FormMain.ErrorHandler(ex, System.Reflection.MethodBase.GetCurrentMethod().Name, GlobalVariables.Linenumber, f)
        End Try
        ToolStripButton3.Enabled = True

        GlobalVariables.Email = InputBox("Enter Email address", "Email address", "aig1066@hotmail.co.uk")

        If GlobalVariables.Email = "" Then
            Exit Sub
        End If
        If ValidateEmail(GlobalVariables.Email) = False Then
            MsgBox("E mail address wrong", vbOKOnly, "Error")
            Exit Sub
        End If
        FormMain.Email("Worm count database - Report for " & GlobalVariables.AlpacaName, "See attached file for report for " & GlobalVariables.AlpacaName & " as requested", ReportName)

    End Sub
    Public Function ValidateEmail(ByVal strCheck As String) As Boolean
        Try
            Dim vEmailAddress As New System.Net.Mail.MailAddress(strCheck)
        Catch ex As Exception
            Return False
        End Try
        Return True
    End Function
    Private Sub ToolStripButton2_Click(sender As Object, e As EventArgs) Handles ToolStripButton2.Click
        Me.Close()
    End Sub
    Public Sub TryToFindAnimal(string1 As String)
        Dim AnimalMatch As Boolean = False
        For j = 0 To GlobalVariables.ds.Tables("Animals").Rows.Count - 1
            If string1 = GlobalVariables.ds.Tables("Animals").Rows(j)("Animal_Name") Then
                AnimalMatch = True
            End If


        Next

        For j = 0 To GlobalVariables.ds.Tables("Animals").Rows.Count - 1
            Dim similarity As Single = GetSimilarity(string1, GlobalVariables.ds.Tables("Animals").Rows(j)("Animal_Name"))
            Dim string2 As String = GlobalVariables.ds.Tables("Animals").Rows(j)("Animal_Name")
            If (similarity > 0.5 Or string2.Contains(string1) Or string1.Contains(string2)) And Not AnimalMatch Then

                Dim rst As DialogResult = MessageBox.Show(string1 & " was not found in Alpaca Manager" & Chr(169) & vbNewLine & GlobalVariables.ds.Tables("Animals").Rows(j)("Animal_Name") & " was found instead." & vbNewLine & "do you wish to change all entries to " & GlobalVariables.ds.Tables("Animals").Rows(j)("Animal_Name"), "Help", MessageBoxButtons.YesNo, MessageBoxIcon.Information)
                If rst = MsgBoxResult.Yes Then
                    ChangeRecord("Alpaca", "Name", string1, GlobalVariables.ds.Tables("Animals").Rows(j)("Animal_Name"))
                    ChangeRecord("TestResults", "TestName", string1, GlobalVariables.ds.Tables("Animals").Rows(j)("Animal_Name"))
                    GlobalVariables.AlpacaName = string1
                    TextBox1.Text = string1
                    Exit For
                End If
            End If
        Next
    End Sub
    Public Sub ChangeRecord(table As String, Field As String, string1 As String, string2 As String)
        Dim sql As String = "Update " & table & " Set " & Field & " = '" & string2 & "' Where " & Field & " = '" & string1 & "'"

        For Each dr As DataRow In GlobalVariables.ds.Tables(table).Rows
            If dr.Item(Field) = string1 Then
                FormMain.ConnectedDB.UpdateDatabase(sql)
            End If
        Next
    End Sub
    Public Function GetSimilarity(string1 As String, string2 As String) As Single
        Dim dis As Single = ComputeDistance(string1, string2)
        Dim maxLen As Single = string1.Length
        If maxLen < string2.Length Then
            maxLen = string2.Length
        End If
        If maxLen = 0.0F Then
            Return 1.0F
        Else
            Return 1.0F - dis / maxLen
        End If
    End Function
    Private Function ComputeDistance(s As String, t As String) As Integer
        Dim n As Integer = s.Length
        Dim m As Integer = t.Length
        Dim distance As Integer(,) = New Integer(n, m) {}
        ' matrix
        Dim cost As Integer = 0
        If n = 0 Then
            Return m
        End If
        If m = 0 Then
            Return n
        End If
        'init1

        Dim i As Integer = 0
        While i <= n
            distance(i, 0) = System.Math.Min(System.Threading.Interlocked.Increment(i), i - 1)
        End While
        Dim j As Integer = 0
        While j <= m
            distance(0, j) = System.Math.Min(System.Threading.Interlocked.Increment(j), j - 1)
        End While
        'find min distance

        For i = 1 To n
            For j = 1 To m
                cost = (If(t.Substring(j - 1, 1) = s.Substring(i - 1, 1), 0, 1))
                distance(i, j) = Math.Min(distance(i - 1, j) + 1, Math.Min(distance(i, j - 1) + 1, distance(i - 1, j - 1) + cost))
            Next
        Next
        Return distance(n, m)
    End Function
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        ChangeRecord("Alpaca", "Name", EnterAninalName, TextBox1.Text)
        ChangeRecord("TestResults", "TestName", EnterAninalName, TextBox1.Text)
        GlobalVariables.AlpacaName = EnterAninalName
        TextBox1.Text = EnterAninalName
    End Sub


End Class