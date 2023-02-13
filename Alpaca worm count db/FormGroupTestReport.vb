Imports System.Threading
Imports System.Net.Mail
Imports System.IO
Public Class FormGroupTestReport

    Dim ReportGroup As String
    Dim Animal As String
    Dim Searchfrom As String
    Dim SearchTo As String
    Dim BuildingInProgress As Boolean
    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs)
        Me.Close()
    End Sub
    Private Sub FormAlpaca_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        CenterToScreen()
        Me.Icon = My.Resources.Form


    End Sub
    Private Sub ToolStrip1_MouseHover(sender As Object, e As EventArgs) Handles ToolStrip1.MouseHover
        Me.Cursor = Cursors.Hand '
    End Sub
    Private Sub ToolStrip1_MouseLeave(sender As Object, e As EventArgs) Handles ToolStrip1.MouseLeave
        Me.Cursor = Cursors.Arrow
    End Sub
    Private Sub CheckedListBox1_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles CheckedListBox1.SelectedIndexChanged
        ReportGroup = " "
        For Each indexChecked In CheckedListBox1.CheckedIndices
            ' The indexChecked variable contains the index of the item.
            ReportGroup += Convert.ToString(CheckedListBox1.Items(indexChecked)) & ","
        Next
        BuildReport()
    End Sub
    Private Sub BuildReport()

        Try
            FormMain.ConnectedDB = New DataBaseFunctions()                      'Open the database connection
            GlobalVariables.ds = FormMain.ConnectedDB.PopulateDataSet()         'Put the datainto a dataset

            Dim site As String = GlobalVariables.AlpacaName
            Dim SelectedPicture As String = GlobalVariables.DbDriveLocation & "Header.png"
            Dim Picture As Bitmap = New Bitmap(SelectedPicture)
            Clipboard.SetImage(FormMain.ResizeImage(Picture))
            Dim PictureFormat As DataFormats.Format = DataFormats.GetFormat(DataFormats.Bitmap)
            Dim NumberOfAnimals As Int16 = GlobalVariables.ds.Tables("Alpaca").Rows.Count - 1
            Dim NumberOfTest As Int16 = GlobalVariables.ds.Tables("TestResults").Rows.Count - 1
            Dim WorstAnimal As String = ""
            Dim Total As Int16 = 0
            Dim HighestCount As Int16 = 0
            Dim RateNumber As Int16 = 0
            Dim Group As String = GlobalVariables.ds.Tables("TestGroup").Rows(0)("GroupName")
            Dim strEggs As String = ""
            Dim Eggs As Int16 = 0
            Dim backcolour As Boolean = False

            RichTextBox1.Clear()
            RichTextBox1.ZoomFactor = 1
            RichTextBox1.Text = ""

            Me.Cursor = Cursors.WaitCursor

            If RichTextBox1.CanPaste(PictureFormat) Then
                RichTextBox1.Paste(PictureFormat)
            End If


            RichTextBox1.AppendText(vbNewLine)
            'RichTextBox1.SelectionFont = New Font("arial", 10, FontStyle.Regular)
            RichTextBox1.SelectionColor = Color.Black
            RichTextBox1.AppendText("Report date   " & DateTime.Today & vbNewLine)
            RichTextBox1.AppendText("----------------------------------------------------------------------------" & vbNewLine)
            RichTextBox1.AppendText(vbNewLine & vbNewLine & vbNewLine)

            'Get a list of all the test groups
            For Each dr As DataRow In GlobalVariables.ds.Tables("TestGroup").Rows
                Group = dr.ItemArray(1)
                'Check that the entry is in the group and also the animal is in that report
                If ReportGroup.Contains(Group) Then  'ReportGroup is the checkbox selection on the form
                    'Now get the data
                    RichTextBox1.AppendText("Test period--" & Group & vbNewLine)
                    RichTextBox1.AppendText("Test types (MM)= McMasters  (MS)=Modified Stoll's" & vbNewLine & vbNewLine)
                    RichTextBox1.AppendText("Name              Test date      EPG  OPG  Test  Notes" & vbNewLine)
                    Dim data() As DataRow


                    'Now strip out the dubplicate animal test
                    Dim CountOfRows As Integer
                    Dim dtDistinct As DataTable = GlobalVariables.ds.Tables("TestResults").Clone
                    Dim dtDuplicates = dtDistinct.Clone


                    For Each drow As DataRow In GlobalVariables.ds.Tables("TestResults").Select("TestNumberName='" & Group & "'", "EPGTotal DESC")
                        CountOfRows = GlobalVariables.ds.Tables("TestResults").Compute("Count([TestName])", "TestNumberName='" & Group & "' AND TestName='" & drow.ItemArray(2) & "'")
                        If CountOfRows > 1 Then
                            dtDuplicates.Rows.Add(drow.ItemArray)
                        Else
                            dtDistinct.Rows.Add(drow.ItemArray)
                        End If
                    Next

                    dtDuplicates.DefaultView.Sort = "TestDate DESC"
                    dtDuplicates = dtDuplicates.DefaultView.ToTable

                    Dim dtDuplicatesRow() As DataRow = dtDuplicates.Select()
                    Dim dtHolding = dtDuplicates.Clone
                    dtHolding.Clear()

                    'Now get the latest test from the duplicates and add the record to the distinct
                    Dim dup As Integer = -1
                    Dim LatestTest As Integer = 0
                    For Each eachAnimal As DataRow In dtDuplicates.Select()
                        dup = 0
                        LatestTest = -1
                        For Each DuplicatesList As DataRow In dtDuplicates.Select()

                            If DuplicatesList.ItemArray(2) = eachAnimal.ItemArray(2) Then
                                'Check the test date and it larger add to distinct
                                Dim foundRow() As DataRow
                                foundRow = dtHolding.Select("TestName='" & eachAnimal.ItemArray(2) & "'")
                                If foundRow.Length = 0 Then
                                    If DuplicatesList.ItemArray(1) > eachAnimal.ItemArray(1) Then
                                        LatestTest = dup
                                        dtHolding.Rows.Add(DuplicatesList.ItemArray)

                                    End If
                                End If
                            End If
                            dup = dup + 1
                        Next
                    Next

                    'Now finaly move the holding data into the distinct datatable
                    For Each dtHoldingList As DataRow In dtHolding.Select()
                        dtDistinct.Rows.Add(dtHoldingList.ItemArray)
                    Next
                    dtDistinct.DefaultView.Sort = "EPGTotal DESC"
                    dtDistinct = dtDistinct.DefaultView.ToTable
                    Dim dtDistinctRow() As DataRow = dtDistinct.Select()
                    data = dtDistinctRow.Clone
                    NumberOfTest = data.Length - 1

                    RichTextBox1.AppendText("_____________________________________________________________________________" & vbNewLine)
                    For j = 0 To NumberOfTest
                        Eggs = 0
                        strEggs = ""

                        If Animal.Contains(data(j).ItemArray(2)) Then



                            'Now build the string name animal first the the test name and last the result and type of egg found
                            Dim str As String = "  "
                            str = str & data(j).ItemArray(2)
                            str = str.PadRight(20, " ")
                            str = str & data(j).ItemArray(1).ToString.Substring(0, 10)
                            str = str.PadRight(34, " ")
                            str = str & CInt(data(j).ItemArray(6)).ToString.PadLeft(4)
                            str = str.PadRight(39, " ")
                            str = str & CInt(data(j).ItemArray(7)).ToString.PadLeft(4)
                            str = str.PadRight(45, " ")




                            If data(j).ItemArray(21) = "MM" Then
                                str = str & "MM"
                            Else
                                str = str & "MS"
                            End If
                            str = str.PadRight(51, " ")



                            If Convert.ToInt64(data(j).ItemArray(8)) > 0 Then
                                    Eggs = System.Convert.ToInt32(data(j).ItemArray(8).ToString.Substring(0, 3))
                                    Eggs = Eggs + System.Convert.ToInt32(data(j).ItemArray(8).ToString.Substring(3, 3))
                                    strEggs = AddEggString(str, strEggs, Eggs, "..Trichostrongyles   ")
                                End If

                                If Convert.ToInt64(data(j).ItemArray(9)) > 0 Then
                                    Eggs = System.Convert.ToInt32(data(j).ItemArray(9).ToString.Substring(0, 3))
                                    Eggs = Eggs + System.Convert.ToInt32(data(j).ItemArray(9).ToString.Substring(3, 3))
                                    strEggs = AddEggString(str, strEggs, Eggs, "..Trichurius   ")
                                End If

                                If Convert.ToInt64(data(j).ItemArray(10)) > 0 Then
                                    Eggs = System.Convert.ToInt32(data(j).ItemArray(10).ToString.Substring(0, 3))
                                    Eggs = Eggs + System.Convert.ToInt32(data(j).ItemArray(10).ToString.Substring(3, 3))
                                    strEggs = AddEggString(str, strEggs, Eggs, "..Nematordirus   ")
                                End If

                                If Convert.ToInt64(data(j).ItemArray(11)) > 0 Then
                                    Eggs = System.Convert.ToInt32(data(j).ItemArray(11).ToString.Substring(0, 3))
                                    Eggs = Eggs + System.Convert.ToInt32(data(j).ItemArray(11).ToString.Substring(3, 3))
                                    strEggs = AddEggString(str, strEggs, Eggs, "..Capillarid   ")
                                End If

                                If Convert.ToInt64(data(j).ItemArray(12)) > 0 Then
                                    Eggs = System.Convert.ToInt32(data(j).ItemArray(12).ToString.Substring(0, 3))
                                    Eggs = Eggs + System.Convert.ToInt32(data(j).ItemArray(12).ToString.Substring(3, 3))
                                    strEggs = AddEggString(str, strEggs, Eggs, "..Moniezid   ")
                                End If

                                If Convert.ToInt64(data(j).ItemArray(13)) > 0 Then
                                    Eggs = System.Convert.ToInt32(data(j).ItemArray(13).ToString.Substring(0, 3))
                                    Eggs = Eggs + System.Convert.ToInt32(data(j).ItemArray(13).ToString.Substring(3, 3))
                                    strEggs = AddEggString(str, strEggs, Eggs, "..EPG Unidentifed   ")
                                End If

                                If Convert.ToInt64(data(j).ItemArray(14)) > 0 Then
                                    Eggs = System.Convert.ToInt32(data(j).ItemArray(14).ToString.Substring(0, 3))
                                    Eggs = Eggs + System.Convert.ToInt32(data(j).ItemArray(14).ToString.Substring(3, 3))
                                    strEggs = AddEggString(str, strEggs, Eggs, "..E-mac   ")
                                End If

                                If Convert.ToInt64(data(j).ItemArray(15)) > 0 Then
                                    Eggs = System.Convert.ToInt32(data(j).ItemArray(15).ToString.Substring(0, 3))
                                    Eggs = Eggs + System.Convert.ToInt32(data(j).ItemArray(15).ToString.Substring(3, 3))
                                    strEggs = AddEggString(str, strEggs, Eggs, "..E-Eivitaesis   ")
                                End If

                                If Convert.ToInt64(data(j).ItemArray(16)) > 0 Then
                                    Eggs = System.Convert.ToInt32(data(j).ItemArray(16).ToString.Substring(0, 3))
                                    Eggs = Eggs + System.Convert.ToInt32(data(j).ItemArray(16).ToString.Substring(3, 3))
                                    strEggs = AddEggString(str, strEggs, Eggs, "..E-alpaca   ")
                                End If

                                If Convert.ToInt64(data(j).ItemArray(17)) > 0 Then
                                    Eggs = System.Convert.ToInt32(data(j).ItemArray(17).ToString.Substring(0, 3))
                                    Eggs = Eggs + System.Convert.ToInt32(data(j).ItemArray(17).ToString.Substring(3, 3))
                                    strEggs = AddEggString(str, strEggs, Eggs, "..E-lamae    ")
                                End If

                                If Convert.ToInt64(data(j).ItemArray(18)) > 0 Then
                                    Eggs = System.Convert.ToInt32(data(j).ItemArray(18).ToString.Substring(0, 3))
                                    Eggs = Eggs + System.Convert.ToInt32(data(j).ItemArray(18).ToString.Substring(3, 3))
                                    strEggs = AddEggString(str, strEggs, Eggs, "..E-punoensis    ")
                                End If

                                If Convert.ToInt64(data(j).ItemArray(19)) > 0 Then
                                    Eggs = System.Convert.ToInt32(data(j).ItemArray(19).ToString.Substring(0, 3))
                                    Eggs = Eggs + System.Convert.ToInt32(data(j).ItemArray(19).ToString.Substring(3, 3))
                                    strEggs = AddEggString(str, strEggs, Eggs, "..OPG Unidentifed     ")
                                End If


                                'Change the colour if it is a fail and send the string to the report
                                If backcolour Then
                                    RichTextBox1.SelectionBackColor = Color.LightSteelBlue
                                Else
                                    RichTextBox1.SelectionBackColor = Color.White
                                End If
                                backcolour = Not backcolour

                                str = str & strEggs
                                str = str.PadRight(74, " ")
                                Dim wwStr As String = WordWrap(str, 75)

                                If (data(j).ItemArray(24) < GlobalVariables.pass) And (Convert.ToInt64(data(j).ItemArray(13)) <= 0) Then
                                    RichTextBox1.SelectionColor = Color.Black
                                    RichTextBox1.AppendText(wwStr)
                                Else
                                    RichTextBox1.SelectionColor = Color.Red
                                    RichTextBox1.AppendText(wwStr)
                                End If

                            End If
                    Next
                    backcolour = Not backcolour
                    If backcolour Then
                        RichTextBox1.SelectionBackColor = Color.LightSteelBlue
                    Else
                        RichTextBox1.SelectionBackColor = Color.White
                    End If
                    RichTextBox1.AppendText(vbCrLf & "_____________________________________________________________________________" & vbNewLine & vbNewLine & vbNewLine)

                End If
            Next

            RichTextBox1.AppendText(vbNewLine & "Treat all animals marked in red")
            RichTextBox1.AppendText(vbNewLine & vbNewLine & vbNewLine)
            RichTextBox1.AppendText(vbNewLine & "*************** End of Report *****************")
            Me.Cursor = Cursors.Default
            RichTextBox1.SelectionStart = 1
            RichTextBox1.ScrollToCaret()
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub
    Public Sub SortTestDb(Group As String)
        'Sort the datatable
        FormMain.ConnectedDB = New DataBaseFunctions()              'Open the database connection
        GlobalVariables.ds = FormMain.ConnectedDB.PopulateDataSet()        'Put the datainto a dataset
        Dim hold1 As DataRow = GlobalVariables.ds.Tables("TestResults").NewRow
        Dim comp As New System.Collections.Comparer(System.Globalization.CultureInfo.CurrentCulture)
        Dim current As Integer = 0
        Dim CurrentPlus1 As Integer = 0
        Dim SameAnimal As Integer = 0
        'If the Token is a interger


        Dim intSortField As Int32 = GlobalVariables.ds.Tables("TestResults").Columns.IndexOf("TestEPGTotal")
        'First sort into result order
        For cntOutter As Int32 = 0 To GlobalVariables.ds.Tables("TestResults").Rows.Count - 1
            For cntInner As Int32 = 0 To GlobalVariables.ds.Tables("TestResults").Rows.Count - 2

                'Only do something if the rows are in the group to test.
                If GlobalVariables.ds.Tables("TestResults").Rows(cntInner).Item(5) = Group And GlobalVariables.ds.Tables("TestResults").Rows(cntOutter).Item(5) = Group Then

                    current = Convert.ToInt32(GlobalVariables.ds.Tables("TestResults").Rows(cntInner + 1).Item(intSortField)) + Convert.ToInt32(GlobalVariables.ds.Tables("TestResults").Rows(cntInner + 1).Item(intSortField + 1))
                    CurrentPlus1 = Convert.ToInt32(GlobalVariables.ds.Tables("TestResults").Rows(cntInner).Item(intSortField)) + Convert.ToInt32(GlobalVariables.ds.Tables("TestResults").Rows(cntInner).Item(intSortField + 1))
                    SameAnimal = comp.Compare(GlobalVariables.ds.Tables("TestResults").Rows(cntOutter).Item(2), GlobalVariables.ds.Tables("TestResults").Rows(cntInner).Item(2))


                    If comp.Compare(current, CurrentPlus1) = 1 Then
                        hold1.BeginEdit()
                        hold1.ItemArray = GlobalVariables.ds.Tables("TestResults").Rows(cntInner + 1).ItemArray
                        hold1.EndEdit()
                        GlobalVariables.ds.Tables("TestResults").Rows.RemoveAt(cntInner + 1)
                        GlobalVariables.ds.Tables("TestResults").Rows.InsertAt(hold1, cntInner)
                        hold1 = GlobalVariables.ds.Tables("TestResults").NewRow
                    End If
                End If

            Next cntInner
        Next cntOutter



        'Next sort into anaimal order
        For cntOutter As Int32 = 0 To GlobalVariables.ds.Tables("TestResults").Rows.Count - 1
            For cntInner As Int32 = 0 To GlobalVariables.ds.Tables("TestResults").Rows.Count - 2

                'Only do something if the rows are in the group to test.
                If GlobalVariables.ds.Tables("TestResults").Rows(cntInner).Item(5) = Group And GlobalVariables.ds.Tables("TestResults").Rows(cntOutter).Item(5) = Group Then
                    SameAnimal = comp.Compare(GlobalVariables.ds.Tables("TestResults").Rows(cntOutter).Item(2), GlobalVariables.ds.Tables("TestResults").Rows(cntInner).Item(2))

                    If SameAnimal = 0 And (cntOutter <> cntInner) Then
                        hold1.BeginEdit()
                        hold1.ItemArray = GlobalVariables.ds.Tables("TestResults").Rows(cntInner).ItemArray
                        hold1.EndEdit()
                        GlobalVariables.ds.Tables("TestResults").Rows.RemoveAt(cntInner)
                        GlobalVariables.ds.Tables("TestResults").Rows.InsertAt(hold1, cntOutter)
                        hold1 = GlobalVariables.ds.Tables("TestResults").NewRow
                    End If
                End If

            Next cntInner
        Next cntOutter

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
                endingPosition = Text.Length 'Text.IndexOf(vbCrLf, startingPosition)

                If endingPosition - startingPosition <= maxLengthOfALine Then
                    'If the string is less then the max length
                    If endingPosition = -1 Then
                        line = Text.Substring(startingPosition, Text.Length - startingPosition)
                        line = Trim(line)
                        SplitLine = SplitLine & line.PadRight(maxLengthOfALine, " ") & vbCrLf
                        Exit While

                    Else
                        line = Text.Substring(startingPosition, endingPosition - startingPosition)
                        line = Trim(line)
                        startingPosition = endingPosition + 1
                        SplitLine = SplitLine & line.PadRight(maxLengthOfALine, " ") & vbCrLf
                    End If
                Else
                    While lineLength + startingPosition <= endingPosition

                        While Text.Substring(startingPosition + lineLength, 1) <> Chr(32)
                            lineLength -= 1
                        End While
                        line = Text.Substring(startingPosition, lineLength)
                        line = Trim(line)
                        If SplitLine = "" Then
                            'The first line
                            SplitLine = SplitLine & line.PadRight(maxLengthOfALine, " ") & vbCrLf
                        Else
                            'All lines but not the last line
                            SplitLine = SplitLine & ("                                                 " & line).PadRight(maxLengthOfALine, " ") & vbCrLf
                        End If
                        If lineLength < maxLengthOfALine Then
                            startingPosition += lineLength
                        Else
                            startingPosition += lineLength + 1
                        End If
                        lineLength = maxLengthOfALine - 51
                    End While

                    line = Text.Substring(startingPosition, endingPosition - startingPosition)
                    line = Trim(line)
                    If line <> "" Then
                        'Last line what is left
                        SplitLine = SplitLine & ("                                                 " & line).PadRight(maxLengthOfALine, " ") & vbCrLf
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
    Private Sub CheckedListBox2_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles CheckedListBox2.SelectedIndexChanged
        Animal = " "
        For Each indexChecked In CheckedListBox2.CheckedIndices
            ' The indexChecked variable contains the index of the item.
            Animal += Convert.ToString(CheckedListBox2.Items(indexChecked)) & ","
        Next
        CheckedListBox2.Sorted = True
        BuildReport()

    End Sub
    Private Sub Button4_Click(sender As System.Object, e As System.EventArgs) Handles Button4.Click
        For idx As Integer = 0 To CheckedListBox2.Items.Count - 1
            Me.CheckedListBox2.SetItemCheckState(idx, CheckState.Checked)
        Next
        CheckedListBox2.Sorted = True

        Animal = " "
        For Each indexChecked In CheckedListBox2.CheckedIndices
            ' The indexChecked variable contains the index of the item.
            Animal += Convert.ToString(CheckedListBox2.Items(indexChecked)) & ","
        Next
        BuildReport()
    End Sub
    Private Sub Button5_Click(sender As System.Object, e As System.EventArgs) Handles Button5.Click
        For idx As Integer = 0 To CheckedListBox2.Items.Count - 1
            Me.CheckedListBox2.SetItemCheckState(idx, CheckState.Unchecked)
        Next
        CheckedListBox2.Sorted = True

        Animal = " "
        For Each indexChecked In CheckedListBox2.CheckedIndices
            ' The indexChecked variable contains the index of the item.
            Animal += Convert.ToString(CheckedListBox2.Items(indexChecked)) & ","
        Next
        BuildReport()
    End Sub
    Private Sub FormGroupTestReport_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        If GlobalVariables.EmailPassword <> "" And GlobalVariables.EmailUsername <> "" And GlobalVariables.EmailServer <> "" Then
            ToolStripButton3.Enabled = True
        Else
            ToolStripButton3.Enabled = False
        End If

        CheckedListBox1.Items.Clear()
        ' CheckedListBox1.Items.Add("All")
        FormMain.ConnectedDB = New DataBaseFunctions()              'Open the database connection
        GlobalVariables.ds = FormMain.ConnectedDB.PopulateDataSet()        'Put the datainto a dataset
        For Each GroupRow As DataRow In GlobalVariables.ds.Tables("TestGroup").Rows
            CheckedListBox1.Items.Add(GroupRow.Item(1))
        Next

        'Now set the last entry in the checkboxlist on (Ticked)

        CheckedListBox1.SetItemChecked(0, True)
        ReportGroup = " "
        For Each indexChecked In CheckedListBox1.CheckedIndices
            ' The indexChecked variable contains the index of the item.
            ReportGroup += Convert.ToString(CheckedListBox1.Items(indexChecked)) & ","
        Next

        CheckedListBox2.Items.Clear()

        FormMain.ConnectedDB = New DataBaseFunctions()              'Open the database connection
        GlobalVariables.ds = FormMain.ConnectedDB.PopulateDataSet()        'Put the datainto a dataset
        For Each GroupRow As DataRow In GlobalVariables.ds.Tables("Alpaca").Rows
            If GroupRow.Item(2) = Convert.ToBoolean(GlobalVariables.OnFarm) Then
                CheckedListBox2.Items.Add(GroupRow.Item(1))
            End If
        Next

        'Now set the last entry in the checkboxlist on (Ticked)
        For idx As Integer = 0 To CheckedListBox2.Items.Count - 1
            Me.CheckedListBox2.SetItemCheckState(idx, CheckState.Checked)
        Next

        Animal = " "
        For Each indexChecked In CheckedListBox2.CheckedIndices
            ' The indexChecked variable contains the index of the item.
            Animal += Convert.ToString(CheckedListBox2.Items(indexChecked)) & ","
        Next

        CheckedListBox2.Sorted = True
        ReportGroup = " "
        For Each indexChecked In CheckedListBox1.CheckedIndices
            ' The indexChecked variable contains the index of the item.
            ReportGroup += Convert.ToString(CheckedListBox1.Items(indexChecked)) & ","
        Next
        BuildReport()

    End Sub
    Private Sub ToolStripButton1_Click(sender As Object, e As EventArgs)
        BuildReport()
    End Sub
    Private Sub ToolStripButton3_Click(sender As Object, e As EventArgs) Handles ToolStripButton3.Click
        Dim printDlg As New PrintDialog()
        'Dim Combo = Mid(Form1.ComboBox1.SelectedItem.ToString, 1, Len(Form1.ComboBox1.SelectedItem.ToString))
        ' Initialize the print dialog with the number of pages in the document.
        printDlg.AllowSomePages = True
        printDlg.PrinterSettings.MinimumPage = 1

        printDlg.PrinterSettings.FromPage = 1
        printDlg.Document = PrintDocument1
        Dim ReportName As String = GlobalVariables.DbDriveLocation & "Print.rtf"
        RichTextBox1.SaveFile(ReportName)


        Dim psi As New ProcessStartInfo
        psi.UseShellExecute = True
        psi.Verb = "print"
        psi.Arguments = printDlg.PrinterSettings.PrinterName.ToString()
        psi.WindowStyle = ProcessWindowStyle.Hidden
        psi.FileName = ReportName ' Here specify a document to be printed

        Process.Start(psi)


        Close()
    End Sub
    Private Sub ToolStripButton4_Click(sender As Object, e As EventArgs) Handles ToolStripButton4.Click
        Dim ReportName As String = GlobalVariables.DbDriveLocation & "Print.rtf"
        RichTextBox1.SaveFile(ReportName)

        GlobalVariables.Email = InputBox("Enter Email address", "Email address", "aig1066@hotmail.co.uk")

        If GlobalVariables.Email = "" Then
            Exit Sub
        End If
        If ValidateEmail(GlobalVariables.Email) = False Then
            MsgBox("E mail address wrong", vbOKOnly, "Error")
            Exit Sub
        End If

        FormMain.Email("Worm count database - Farm report", "See attached file for report as requested", ReportName)
    End Sub
    Public Function ValidateEmail(ByVal strCheck As String) As Boolean
        Try
            Dim vEmailAddress As New System.Net.Mail.MailAddress(strCheck)
        Catch ex As Exception
            Return False
        End Try
        Return True
    End Function
    Sub EmailgroundThread()
        Try

            'GlobalVariables.Email = "aig1066@hotmail.co.uk"

            GlobalVariables.EmailSending = True
            Dim Smtp_Server As New SmtpClient
            Dim e_mail As New MailMessage()
            Smtp_Server.UseDefaultCredentials = False
            Smtp_Server.Credentials = New Net.NetworkCredential(GlobalVariables.EmailUsername, GlobalVariables.EmailPassword)
            Smtp_Server.Port = 587
            Smtp_Server.EnableSsl = True
            Smtp_Server.Host = GlobalVariables.EmailServer
            e_mail = New MailMessage()
            e_mail.From = New MailAddress(GlobalVariables.EmailUsername)
            If GlobalVariables.Email = "" Or GlobalVariables.Email = "sample@farm.com" Then
                MsgBox("Email recipient has to been entered, go to settings and correct.", vbOKOnly & vbCritical, "Error")
                Exit Sub
            End If
            e_mail.To.Add(GlobalVariables.Email)
            e_mail.Subject = GlobalVariables.Subject
            e_mail.IsBodyHtml = True

            'e_mail.Body = GlobalVariables.Body
            Dim body As String = File.ReadAllText("C:\WormCountDATA\body.html")
            body = body.Replace("<%Tag01%>", GlobalVariables.AlpacaName)
            e_mail.Body = body


            If GlobalVariables.Attachment <> "" Then
                Dim AttachmentFile As String = GlobalVariables.Attachment
                e_mail.Attachments.Add(New Attachment(AttachmentFile))
            End If
            Smtp_Server.Send(e_mail)

            e_mail.Attachments.Dispose()



        Catch ex As Exception

            Dim st As New StackTrace(True)
            st = New StackTrace(ex, True)
            GlobalVariables.Linenumber = st.GetFrame(0).GetFileLineNumber().ToString()
            Dim f As New StackFrame
            FormMain.ErrorHandler(ex, System.Reflection.MethodBase.GetCurrentMethod().Name, GlobalVariables.Linenumber, f)
            GlobalVariables.EmailSending = False
            Exit Sub
        End Try
        GlobalVariables.EmailSending = False

    End Sub
    Private Sub ToolStripButton2_Click(sender As Object, e As EventArgs) Handles ToolStripButton2.Click
        Close()
    End Sub


End Class