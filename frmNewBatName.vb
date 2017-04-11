Public Class frmNewBatName
    Private bOkayed As Boolean
    Public ReadOnly Property Okayed() As Boolean
        Get
            Return bOkayed
        End Get
    End Property

    Private Sub butCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles butCancel.Click
        bOkayed = False
        Me.Close()
    End Sub

    Private Sub butOkay_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles butOkay.Click
        bOkayed = True
        Me.Close()
    End Sub

    Private Sub frmNewBatName_Shown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shown

        txtBatName.Text = ""

        Me.Left = frmMain.Left + frmMain.Width / 2 - Me.Width / 2
        Me.Top = frmMain.Top + frmMain.Height / 2 - Me.Height / 2
    End Sub
End Class