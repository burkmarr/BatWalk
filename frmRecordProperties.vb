Public Class frmRecordProperties
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

    Private Sub cbBat_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cbBat.KeyUp

        'This shifts the focus to the next button
        If e.KeyCode = Keys.Return Then
            txtAbundance.Focus()
            txtComment.SelectionStart = txtAbundance.Text.Length
        End If
    End Sub

    Private Sub txtAbundance_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtAbundance.KeyPress

        If e.KeyChar = Microsoft.VisualBasic.ChrW(Keys.Return) Then

            'Suppresses the windows ding sound when enter is presses in a single-line textbox
            e.Handled = True
            txtComment.Focus()
            txtComment.SelectionStart = txtComment.Text.Length
        End If
    End Sub

    Private Sub txtComment_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtComment.KeyPress

        If e.KeyChar = Microsoft.VisualBasic.ChrW(Keys.Return) Then

            'Suppresses the windows ding sound when enter is presses in a single-line textbox
            e.Handled = True
            butOkay.Focus()
        End If
    End Sub

    Private Sub frmRecordProperties_Shown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shown

        cbBat.Focus()
        cbBat.SelectionStart = cbBat.Text.Length

        Me.Left = frmMain.Left + frmMain.Width / 2 - Me.Width / 2
        Me.Top = frmMain.Top + frmMain.Height / 2 - Me.Height / 2
    End Sub
End Class