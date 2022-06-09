Public Class FormPass

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim AsInt As Integer
        If Integer.TryParse(TextBox1.Text, AsInt) Then
            GlobalVariables.pass = TextBox1.Text
        Else
            MessageBox.Show("Please enter a number", "ERROR")
        End If
        FormMain.DataGridViewRefresh()
        Close()

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Close()
    End Sub

    Private Sub FormNewAnimal_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        CenterToScreen()
        Me.Icon = My.Resources.Form
        TextBox1.Text = GlobalVariables.pass
    End Sub
End Class