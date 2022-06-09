

Public Class FormGroupTestReport

    Dim ReportGroup As String
    Dim Animal As String

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




    End Sub


    Private Sub BuildReport()

        Try
            FormMain.ConnectedDB = New DataBaseFunctions()              'Open the database connection
            GlobalVariables.ds = FormMain.ConnectedDB.PopulateDataSet()        'Put the datainto a dataset

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
                    RichTextBox1.AppendText("Test period--" & Group & vbNewLine & vbNewLine)
                    RichTextBox1.AppendText("Name              Test date      Total EPG" & vbNewLine)

                    SortTestDb(Group)


                    RichTextBox1.AppendText("_____________________________________________________________________________" & vbNewLine)
                    For j = 0 To GlobalVariables.ds.Tables("TestResults").Rows.Count - 1
                        Total = Convert.ToInt32(GlobalVariables.ds.Tables("TestResults").Rows(j)("TestEPGTotal")) + Convert.ToInt32(GlobalVariables.ds.Tables("TestResults").Rows(j)("TestOPGTotal"))
                        Eggs = 0
                        strEggs = ""



                        'Check the test result from the dataqbase is in the current group date
                        If GlobalVariables.ds.Tables("TestResults").Rows(j)("TestNumberName") = Group Then


                            'Animal is a list of the animals ticked in the checkbox selection in the form.
                            If Animal.Contains(GlobalVariables.ds.Tables("TestResults").Rows(j)("TestName")) Then

                                'Now build the string name animal first the the test name and last the result and type of egg found
                                Dim str As String = "  "
                                'If j > 1 Then
                                'If GlobalVariables.ds.Tables("TestResults").Rows(j)("TestName") = GlobalVariables.ds.Tables("TestResults").Rows(j - 1)("TestName") Then
                                'str = str.PadRight(20, " ")

                                'Else
                                str = str & GlobalVariables.ds.Tables("TestResults").Rows(j)("TestName")
                                str = str.PadRight(20, " ")
                                ' End If
                                'End If

                                str = str & GlobalVariables.ds.Tables("TestResults").Rows(j)("TestDate").ToString.Substring(0, 10)
                                str = str.PadRight(35, " ")

                                str = str & Total.ToString
                                str = str.PadRight(40, " ")


                                If Convert.ToInt16(GlobalVariables.ds.Tables("TestResults").Rows(j)("TestTrichostrongyles")) > 0 Then
                                    Eggs = System.Convert.ToInt32(GlobalVariables.ds.Tables("TestResults").Rows(j)("TestTrichostrongyles").ToString.Substring(0, 2))
                                    Eggs = Eggs + System.Convert.ToInt32(GlobalVariables.ds.Tables("TestResults").Rows(j)("TestTrichostrongyles").ToString.Substring(2, 2))
                                    strEggs = AddEggString(str, strEggs, Eggs, "..Trichostrongyles   ")
                                End If

                                If Convert.ToInt16(GlobalVariables.ds.Tables("TestResults").Rows(j)("TestTrichurius")) > 0 Then
                                    Eggs = System.Convert.ToInt32(GlobalVariables.ds.Tables("TestResults").Rows(j)("TestTrichurius").ToString.Substring(0, 2))
                                    Eggs = Eggs + System.Convert.ToInt32(GlobalVariables.ds.Tables("TestResults").Rows(j)("TestTrichurius").ToString.Substring(2, 2))
                                    strEggs = AddEggString(str, strEggs, Eggs, "..Trichurius   ")
                                End If

                                If Convert.ToInt16(GlobalVariables.ds.Tables("TestResults").Rows(j)("TestNematordirus")) > 0 Then
                                    Eggs = System.Convert.ToInt32(GlobalVariables.ds.Tables("TestResults").Rows(j)("TestNematordirus").ToString.Substring(0, 2))
                                    Eggs = Eggs + System.Convert.ToInt32(GlobalVariables.ds.Tables("TestResults").Rows(j)("TestNematordirus").ToString.Substring(2, 2))
                                    strEggs = AddEggString(str, strEggs, Eggs, "..Nematordirus   ")
                                End If

                                If Convert.ToInt16(GlobalVariables.ds.Tables("TestResults").Rows(j)("TestCapillarid")) > 0 Then
                                    Eggs = System.Convert.ToInt32(GlobalVariables.ds.Tables("TestResults").Rows(j)("TestCapillarid").ToString.Substring(0, 2))
                                    Eggs = Eggs + System.Convert.ToInt32(GlobalVariables.ds.Tables("TestResults").Rows(j)("TestCapillarid").ToString.Substring(2, 2))
                                    strEggs = AddEggString(str, strEggs, Eggs, "..Capillarid   ")
                                End If

                                If Convert.ToInt16(GlobalVariables.ds.Tables("TestResults").Rows(j)("TestMoniezid")) > 0 Then
                                    Eggs = System.Convert.ToInt32(GlobalVariables.ds.Tables("TestResults").Rows(j)("TestMoniezid").ToString.Substring(0, 2))
                                    Eggs = Eggs + System.Convert.ToInt32(GlobalVariables.ds.Tables("TestResults").Rows(j)("TestMoniezid").ToString.Substring(2, 2))
                                    strEggs = AddEggString(str, strEggs, Eggs, "..Moniezid   ")
                                End If

                                If Convert.ToInt16(GlobalVariables.ds.Tables("TestResults").Rows(j)("TestEmac")) > 0 Then
                                    Eggs = System.Convert.ToInt32(GlobalVariables.ds.Tables("TestResults").Rows(j)("TestEmac").ToString.Substring(0, 2))
                                    Eggs = Eggs + System.Convert.ToInt32(GlobalVariables.ds.Tables("TestResults").Rows(j)("TestEmac").ToString.Substring(2, 2))
                                    strEggs = AddEggString(str, strEggs, Eggs, "..E-mac   ")
                                End If

                                If Convert.ToInt16(GlobalVariables.ds.Tables("TestResults").Rows(j)("TestEivitaesis")) > 0 Then
                                    Eggs = System.Convert.ToInt32(GlobalVariables.ds.Tables("TestResults").Rows(j)("TestEivitaesis").ToString.Substring(0, 2))
                                    Eggs = Eggs + System.Convert.ToInt32(GlobalVariables.ds.Tables("TestResults").Rows(j)("TestEivitaesis").ToString.Substring(2, 2))
                                    strEggs = AddEggString(str, strEggs, Eggs, "..E-ivitaesis   ")
                                End If

                                If Convert.ToInt16(GlobalVariables.ds.Tables("TestResults").Rows(j)("TestEalpacae")) > 0 Then
                                    Eggs = System.Convert.ToInt32(GlobalVariables.ds.Tables("TestResults").Rows(j)("TestEalpacae").ToString.Substring(0, 2))
                                    Eggs = Eggs + System.Convert.ToInt32(GlobalVariables.ds.Tables("TestResults").Rows(j)("TestEalpacae").ToString.Substring(2, 2))
                                    strEggs = AddEggString(str, strEggs, Eggs, "..E-alpacae   ")
                                End If

                                If Convert.ToInt16(GlobalVariables.ds.Tables("TestResults").Rows(j)("TestElamae")) > 0 Then
                                    Eggs = System.Convert.ToInt32(GlobalVariables.ds.Tables("TestResults").Rows(j)("TestElamae").ToString.Substring(0, 2))
                                    Eggs = Eggs + System.Convert.ToInt32(GlobalVariables.ds.Tables("TestResults").Rows(j)("TestElamae").ToString.Substring(2, 2))
                                    strEggs = AddEggString(str, strEggs, Eggs, "..E-lamae   ")
                                End If

                                If Convert.ToInt16(GlobalVariables.ds.Tables("TestResults").Rows(j)("TestEpunoensis")) > 0 Then
                                    Eggs = System.Convert.ToInt32(GlobalVariables.ds.Tables("TestResults").Rows(j)("TestEpunoensis").ToString.Substring(0, 2))
                                    Eggs = Eggs + System.Convert.ToInt32(GlobalVariables.ds.Tables("TestResults").Rows(j)("TestEpunoensis").ToString.Substring(2, 2))
                                    strEggs = AddEggString(str, strEggs, Eggs, "..E-punoensis   ")
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

                                If (Total < GlobalVariables.pass) And (Convert.ToInt16(GlobalVariables.ds.Tables("TestResults").Rows(j)("TestEmac")) <= 0) Then
                                    RichTextBox1.SelectionColor = Color.Black
                                    RichTextBox1.AppendText(wwstr)
                                Else
                                    RichTextBox1.SelectionColor = Color.Red
                                    RichTextBox1.AppendText(wwStr)
                                End If
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


            NumberOfTest = GlobalVariables.ds.Tables("TestResults").Rows.Count - 1
            Me.Cursor = Cursors.Default
        Catch ex As Exception

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
                            SplitLine = SplitLine & ("                                      " & line).PadRight(maxLengthOfALine, " ") & vbCrLf
                        End If
                        If lineLength < maxLengthOfALine Then
                            startingPosition += lineLength
                        Else
                            startingPosition += lineLength + 1
                        End If
                        lineLength = maxLengthOfALine - 39
                    End While

                    line = Text.Substring(startingPosition, endingPosition - startingPosition)
                    line = Trim(line)
                    If line <> "" Then
                        'Last line what is left
                        SplitLine = SplitLine & ("                                      " & line).PadRight(maxLengthOfALine, " ") & vbCrLf
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

    End Sub




    Private Sub ToolStripButton1_Click(sender As Object, e As EventArgs) Handles ToolStripButton1.Click
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

    Private Sub ToolStripButton4_Click(sender As Object, e As EventArgs) Handles ToolStripButton4.Click
        Dim ReportName As String = GlobalVariables.DbDriveLocation & "Print.rtf"
        RichTextBox1.SaveFile(ReportName)

        FormMain.Email("Worm count database - Farm report", "See attached file for report as requested", ReportName)
    End Sub

    Private Sub ToolStripButton2_Click(sender As Object, e As EventArgs) Handles ToolStripButton2.Click
        Close()
    End Sub
End Class